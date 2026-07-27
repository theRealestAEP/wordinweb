import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection, DocProvider } from "../src/hub.js";
import { ServerMessage, PROTOCOL_VERSION } from "@wordinweb/collab/server";
import type { InsertTextIntent } from "@wordinweb/collab/server";
import { MetricsObservability, NO_OP_OBSERVABILITY } from "../src/observability.js";

function blankDoc(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  });
}

const provider: DocProvider = { load: () => blankDoc("hi") };

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
}

/** A metrics sink that never logs (so tests stay quiet regardless of env). */
const quietObs = () => new MetricsObservability({ log: false });

describe("observability — refusals/kicks by reason", () => {
  it("a takeover produces a taken-over kick AND refusal (the single-tab regression signal)", async () => {
    const obs = quietObs();
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);
    const zombie = new FakeConn("t1");
    const fresh = new FakeConn("t2");
    await hub.handle(zombie, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    await hub.handle(fresh, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", takeover: true, sinceSeq: 0 });

    const s = obs.snapshot();
    expect(s.kicksByReason["taken-over"]).toBe(1);
    // The kick's refusal frame is tallied too — refused-by-reason tracks the wire.
    expect(s.refusedByReason["taken-over"]).toBe(1);
    // Both sockets said a valid hello; both joined before the incumbent was cut.
    expect(s.counters.hellosAccepted).toBe(2);
  });

  it("a second same-identity hello WITHOUT takeover is refused already-open, no kick", async () => {
    const obs = quietObs();
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);
    const tab1 = new FakeConn("t1");
    const tab2 = new FakeConn("t2");
    await hub.handle(tab1, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    await hub.handle(tab2, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });

    const s = obs.snapshot();
    expect(s.refusedByReason["already-open"]).toBe(1);
    expect(s.kicksByReason).toEqual({}); // nobody was kicked
    expect(s.counters.hellosAccepted).toBe(1); // only tab1 joined
  });

  it("version-mismatch and no-session refusals are tallied by reason", async () => {
    const obs = quietObs();
    const hub = new CollabHub(/*provider*/ null, undefined, undefined, undefined, undefined, obs);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION + 1, docId: "d", clientId: "c", sinceSeq: 0 });
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "ghost", clientId: "c", sinceSeq: 0 });
    const s = obs.snapshot();
    expect(s.refusedByReason["version-mismatch"]).toBe(1);
    expect(s.refusedByReason["no-session"]).toBe(1);
  });
});

describe("observability — submits and broadcasts", () => {
  it("counts accepted submits + broadcasts, and rejected submits by reason", async () => {
    const obs = quietObs();
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });
    // A forged clientId is rejected before sequencing.
    await hub.handle(a, { t: "submit", intent: { ...intent, clientId: "mallory", clientSeq: 9 } });

    const s = obs.snapshot();
    expect(s.counters.submitsAccepted).toBe(1);
    expect(s.counters.broadcasts).toBe(1);
    expect(s.counters.submitsRejected).toBe(1);
    expect(s.refusedByReason["client-id-mismatch"]).toBe(1);
  });
});

describe("observability — room + roster gauges", () => {
  it("room create/evict move the counters and the rooms gauge returns to zero", async () => {
    let now = 0;
    const obs = quietObs();
    const hub = new CollabHub(provider, undefined, undefined, () => now, undefined, obs);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });

    let s = obs.snapshot();
    expect(s.counters.roomsCreated).toBe(1);
    expect(s.gauges.rooms).toBe(1);
    expect(s.gauges.rosterMax).toBe(2); // both alice+bob on the fan-out

    hub.disconnect(a);
    hub.disconnect(b);
    now = 10_000_000;
    hub.sweepRooms();
    s = obs.snapshot();
    expect(s.counters.roomsEvicted).toBe(1);
    expect(s.gauges.rooms).toBe(0);
  });
});

describe("observability — transport open/close balance the live gauge", () => {
  it("opened/closed are counted once per socket via the ws adapter", async () => {
    // Import lazily to keep this alongside the ws-adapter usage.
    const { attachWebSocketServer } = await import("../src/ws.js");
    const obs = quietObs();
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);

    let onConn: ((s: unknown) => void) | null = null;
    attachWebSocketServer({ on: (_e, cb) => { onConn = cb as (s: unknown) => void; } } as never, hub, obs);
    let onClose: (() => void) | null = null;
    const socket = { send: () => {}, on: (e: string, cb: (() => void) & ((d: unknown) => void)) => { if (e === "close") onClose = cb; } };
    onConn!(socket);
    let s = obs.snapshot();
    expect(s.counters.connectionsOpened).toBe(1);
    expect(s.gauges.liveConnections).toBe(1);
    onClose!();
    s = obs.snapshot();
    expect(s.counters.connectionsClosed).toBe(1);
    expect(s.gauges.liveConnections).toBe(0); // balanced
  });
});

describe("observability — structured logging gate + safety", () => {
  it("emits one-line JSON events with a monotonic seq, and only safe fields", async () => {
    const lines: string[] = [];
    let t = 1000;
    const obs = new MetricsObservability({ level: "debug", out: (l) => lines.push(l), now: () => t++ });
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);
    const tab1 = new FakeConn("t1");
    const tab2 = new FakeConn("t2");
    await hub.handle(tab1, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "secret-doc", clientId: "secret-client", sinceSeq: 0 });
    await hub.handle(tab2, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "secret-doc", clientId: "secret-client", takeover: true, sinceSeq: 0 });

    expect(lines.length).toBeGreaterThan(0);
    const records = lines.map((l) => JSON.parse(l) as Record<string, unknown>);
    // Every record carries the shipping schema (ts/level/ev/seq); seq is
    // strictly increasing, which is how a collector detects dropped lines.
    let prev = 0;
    for (const r of records) {
      expect(typeof r.ev).toBe("string");
      expect(typeof r.ts).toBe("string");
      expect(new Date(r.ts as string).toISOString()).toBe(r.ts); // real ISO-8601
      expect(["debug", "info", "warn", "error"]).toContain(r.level);
      expect(typeof r.seq).toBe("number");
      expect(r.seq as number).toBeGreaterThan(prev);
      prev = r.seq as number;
    }
    // ZERO-CUSTODY: no docId/clientId (or their values) ever appear in a line.
    const all = lines.join("\n");
    expect(all).not.toContain("secret-doc");
    expect(all).not.toContain("secret-client");
    // The kick/refusal reason (safe fixed vocabulary) IS present.
    expect(records.some((r) => r.ev === "kick" && r.reason === "taken-over")).toBe(true);
  });

  it("logs nothing at level silent, while counters keep moving", async () => {
    const lines: string[] = [];
    const obs = new MetricsObservability({ level: "silent", out: (l) => lines.push(l) });
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, obs);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "c", sinceSeq: 0 });
    // Counters still move; the log stays silent.
    expect(obs.snapshot().counters.hellosAccepted).toBe(1);
    expect(lines).toEqual([]);
  });

  it("the no-op sink records nothing and never throws (hub default)", async () => {
    expect(NO_OP_OBSERVABILITY.snapshot().counters.hellosAccepted).toBe(0);
    const hub = new CollabHub(provider); // default obs = no-op
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "c", sinceSeq: 0 });
    expect(c.received.some((m) => m.t === "welcome")).toBe(true);
  });
});
