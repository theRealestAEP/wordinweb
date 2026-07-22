import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection, DocProvider } from "../src/hub.js";
import { ServerMessage, PROTOCOL_VERSION } from "../src/protocol.js";
import type { InsertTextIntent } from "@wordinweb/collab/server";

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
  last(): ServerMessage {
    return this.received[this.received.length - 1];
  }
}

describe("CollabHub", () => {
  it("refuses a version-mismatched client at hello", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c1");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION + 1, docId: "d", sinceSeq: 0 });
    expect(c.last()).toEqual({ t: "refused", reason: "version-mismatch" });
  });

  it("welcomes a joining client with a snapshot and the current seq", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c1");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    const w = c.last();
    expect(w.t).toBe("welcome");
    if (w.t === "welcome") {
      expect(w.seq).toBe(0);
      expect(w.snapshot.length).toBeGreaterThan(0);
      expect(w.tail).toEqual([]);
    }
    expect(hub.activeDocs()).toEqual(["d"]);
  });

  it("refuses submit from a connection that has not joined", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c1");
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(c, { t: "submit", intent });
    expect(c.last()).toEqual({ t: "refused", reason: "not-joined" });
  });

  it("broadcasts a submitted intent to every participant in the room", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });

    // Address paragraph 0's run using ids the session assigned (1 = para, 2 = run
    // in deterministic parse order for a single-paragraph doc).
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });

    for (const c of [a, b]) {
      const msg = c.last();
      expect(msg.t).toBe("broadcast");
      if (msg.t === "broadcast") {
        expect(msg.entries).toHaveLength(1);
        expect(msg.entries[0].kind).toBe("applied");
        expect(msg.entries[0].seq).toBe(1);
      }
    }
  });

  it("isolates rooms: a submit on one doc does not reach another", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "docA", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "docB", sinceSeq: 0 });
    const before = b.received.length;
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });
    expect(b.received.length).toBe(before); // b got nothing new
    expect(hub.activeDocs().sort()).toEqual(["docA", "docB"]);
  });

  it("drops a connection on disconnect", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    hub.disconnect(b);
    const beforeB = b.received.length;
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });
    expect(b.received.length).toBe(beforeB); // disconnected b receives nothing
  });
});

import { attachWebSocketServer, WsServer, WsSocket } from "../src/ws.js";

describe("attachWebSocketServer", () => {
  it("bridges ws-style frames to the hub and serializes responses", async () => {
    let onConn: ((s: WsSocket) => void) | null = null;
    const wss: WsServer = { on: (_e, cb) => { onConn = cb as (s: WsSocket) => void; } };
    const hub = new CollabHub(provider);
    attachWebSocketServer(wss, hub);

    const sent: string[] = [];
    let onMsg: ((d: unknown) => void) | null = null;
    const socket: WsSocket = {
      send: (d) => sent.push(d),
      on: (e: "message" | "close", cb: ((d: unknown) => void) & (() => void)) => {
        if (e === "message") onMsg = cb;
      },
    };
    onConn!(socket);
    onMsg!(JSON.stringify({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 }));
    await new Promise((r) => setTimeout(r, 0)); // flush the fire-and-forget handle()
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]).t).toBe("welcome");
  });

  it("ignores a malformed frame without crashing", async () => {
    let onConn: ((s: WsSocket) => void) | null = null;
    const wss: WsServer = { on: (_e, cb) => { onConn = cb as (s: WsSocket) => void; } };
    attachWebSocketServer(wss, new CollabHub(provider));
    let onMsg: ((d: unknown) => void) | null = null;
    const socket: WsSocket = { send: () => {}, on: (e, cb: ((d: unknown) => void) & (() => void)) => { if (e === "message") onMsg = cb; } };
    onConn!(socket);
    expect(() => onMsg!("{ not json")).not.toThrow();
  });
});

import { InMemoryStorage } from "../src/storage.js";

describe("CollabHub persistence + rehydration", () => {
  it("persists sequenced entries and rehydrates a new hub from storage", async () => {
    const storage = new InMemoryStorage();
    const hub1 = new CollabHub(provider, storage);
    const a = new FakeConn("a");
    await hub1.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub1.handle(a, { t: "submit", intent });

    // The entry is durably stored.
    const log = await storage.readLog("d", 0);
    expect(log).toHaveLength(1);
    expect(log[0].kind).toBe("applied");

    // A fresh hub (simulating a restart) rehydrates the doc from the log tail:
    // the joining client's welcome reports the persisted seq.
    const hub2 = new CollabHub(provider, storage);
    const b = new FakeConn("b");
    await hub2.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    const w = b.last();
    expect(w.t).toBe("welcome");
    if (w.t === "welcome") expect(w.seq).toBe(1);
  });

  it("deduplicates a resent intent across the persisted log", async () => {
    const storage = new InMemoryStorage();
    const hub = new CollabHub(provider, storage);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 5, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "X" };
    await hub.handle(a, { t: "submit", intent });
    await hub.handle(a, { t: "submit", intent }); // resend
    const log = await storage.readLog("d", 0);
    expect(log).toHaveLength(1); // deduped
  });
});

import type { PresencePosition } from "../src/protocol.js";

describe("CollabHub presence", () => {
  const pos: PresencePosition = { anchor: { blockId: 1, runId: 2, offset: 3 } };

  it("fans a presence update out to other participants but not the sender", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    const aBefore = a.received.length;
    await hub.handle(a, { t: "presence", position: pos });
    expect(a.received.length).toBe(aBefore); // sender does not echo
    const msg = b.last();
    expect(msg.t).toBe("presence");
    if (msg.t === "presence") {
      expect(msg.participant).toBe("a");
      expect(msg.position).toEqual(pos);
    }
  });

  it("ignores presence from a connection that has not joined", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "presence", position: pos });
    expect(c.received).toHaveLength(0); // ignored, not refused
  });

  it("does not persist presence (storage untouched)", async () => {
    const storage = new InMemoryStorage();
    const hub = new CollabHub(provider, storage);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 });
    await hub.handle(a, { t: "presence", position: pos });
    expect(await storage.readLog("d", 0)).toHaveLength(0);
  });
});
