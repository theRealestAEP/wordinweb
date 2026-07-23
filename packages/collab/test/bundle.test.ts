import { describe, expect, it } from "vitest";
import { CollabConnection } from "../src/connection.js";
import { BundlePersister, InMemoryBundleStore, type DocBundle } from "../src/bundle.js";
import { DocumentSession } from "../src/session.js";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";
import { zipSync, strToU8 } from "fflate";

/** Minimal one-paragraph docx ("hi") — same fixture shape as the batches. */
function blankDocx(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">hi</w:t></w:r></w:p></w:body></w:document>`;
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

/**
 * Bundle persistence + resume (plan doc 12 §4/§5). The "server" here is a
 * real DocumentSession behind a message pump — same authority code as the
 * hub — so resume semantics (dedup replay, epoch change) are pinned against
 * the true sequencer, not a mock.
 */

/** A tiny synchronous hub: one session, fan-out, per-conn delivery. */
function miniServer(genesisId: string) {
  const session = new DocumentSession(DocxDocument.load(blankDocx()));
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        if (msg.t === "hello") {
          const cp = session.checkpoint();
          peer.deliver({
            t: "welcome", docId: msg.docId, seq: cp.seq, snapshot: Buffer.from(cp.docx).toString("base64"),
            sidecar: cp.sidecar, tail: session.entriesSince(cp.seq), genesisId, mode: "plaintext",
          });
        } else if (msg.t === "submit") {
          const entry = session.submit(msg.intent);
          for (const p of peers) p.deliver({ t: "broadcast", entries: [entry] });
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return { session, attach };
}

const text = (conn: CollabConnection): string => {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(conn.doc!.docRoot as never);
};
const ins = (at: number, t: string) =>
  ({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: t }) as never;

describe("BundlePersister throttle (round-4 F8: throttle, not debounce)", () => {
  function harness() {
    const srv = miniServer("g1");
    const conn = new CollabConnection(srv.attach(), "alice");
    conn.join("d");
    const store = new InMemoryBundleStore();
    let now = 0;
    const timers: { at: number; fn: () => void }[] = [];
    const p = new BundlePersister(conn, store, "d", {
      throttleMs: 1000,
      now: () => now,
      setTimer: (fn, ms) => { const t = { at: now + ms, fn }; timers.push(t); return t; },
      clearTimer: (t) => { const i = timers.indexOf(t as never); if (i >= 0) timers.splice(i, 1); },
    });
    const advance = (ms: number) => {
      now += ms;
      for (const t of [...timers].filter((t) => t.at <= now)) { timers.splice(timers.indexOf(t), 1); t.fn(); }
    };
    return { conn, store, p, advance, timers };
  }

  it("a sustained burst coalesces to one leading + one trailing write — never a deferred-forever debounce", () => {
    const { conn, store, p, advance } = harness();
    conn.submit(ins(0, "a"));
    p.notify(); // leading edge: writes immediately
    expect(store.writes).toBe(1);
    for (let i = 1; i <= 30; i++) { conn.submit(ins(i, "x")); p.notify(); advance(20); }
    // 600ms of continuous typing: a debounce would still be waiting; the
    // throttle has armed exactly one trailing write inside the window.
    expect(store.writes).toBe(1);
    advance(1000); // window closes → the trailing write fires
    expect(store.writes).toBe(2);
  });

  it("flush() writes immediately and cancels the trailing timer", async () => {
    const { conn, store, p, advance, timers } = harness();
    conn.submit(ins(0, "a"));
    p.notify();
    conn.submit(ins(1, "b"));
    p.notify(); // arms trailing
    await p.flush(); // pagehide: immediate write, timer cancelled
    expect(store.writes).toBe(2);
    expect(timers.length).toBe(0);
    advance(5000);
    expect(store.writes).toBe(2); // nothing left armed
  });

  it("persists the CONFIRMED state and round-trips byte-identically", async () => {
    const { conn, store, p } = harness();
    conn.submit(ins(0, "hello"));
    await p.flush();
    const b = (await store.get("d"))!;
    expect(b.genesisId).toBe("g1");
    expect(b.clientSeq).toBe(1);
    expect(b.pending).toEqual([]); // miniServer echoed synchronously → confirmed
    // The bundle's bytes reload to the same document.
    const reloaded = DocxDocument.load(b.confirmedBytes);
    expect(serializeXml(reloaded.docRoot)).toBe(serializeXml(conn.doc!.docRoot));
  });
});

describe("resume from bundle (doc 12 §5)", () => {
  /** Build a bundle by working in session 1, then "crash". */
  function workAndBundle(srv: ReturnType<typeof miniServer>): DocBundle {
    const conn = new CollabConnection(srv.attach(), "alice");
    conn.join("d");
    conn.submit(ins(0, "hello"));
    return conn.exportBundle("d")!;
  }

  it("case 1 — same epoch: seamless rejoin, clientSeq watermark restored (new edits are not deduped away)", () => {
    const srv = miniServer("g1");
    const bundle = workAndBundle(srv);
    expect(bundle.clientSeq).toBe(1);

    // Same browser later: fresh connection object, resume from the bundle.
    const conn2 = new CollabConnection(srv.attach(), "alice");
    conn2.resume(bundle);
    expect(conn2.ready).toBe(true);
    expect(text(conn2)).toBe("hellohi");
    // The watermark is the regression: a fresh counter would emit
    // (alice, 1) again and the server would dedup the NEW edit silently.
    conn2.submit(ins(5, "!"));
    expect(text(conn2)).toBe("hello!hi");
    expect(srv.session.seq).toBe(2); // genuinely sequenced, not deduped
  });

  it("case 1 crash-before-ack: pending replays exactly once via server dedup", () => {
    const srv = miniServer("g1");
    const conn = new CollabConnection(srv.attach(), "alice");
    conn.join("d");
    conn.submit(ins(0, "hello"));
    // Simulate crash-before-ack: a pending intent the server never saw.
    const bundle = conn.exportBundle("d")!;
    bundle.pending = [{ kind: "insertText", clientId: "alice", clientSeq: 2, base: 1, at: { blockId: 1, runId: 2, offset: 5 }, text: "!" } as never];
    bundle.clientSeq = 2;

    const conn2 = new CollabConnection(srv.attach(), "alice");
    conn2.resume(bundle);
    expect(text(conn2)).toBe("hello!hi"); // replayed and applied…
    // …and a DUPLICATE resume (second crash, same bundle) does not double it:
    const conn3 = new CollabConnection(srv.attach(), "alice", {});
    conn3.resume(bundle);
    expect(text(conn3)).toBe("hello!hi"); // dedup swallowed the re-replay
    expect(srv.session.seq).toBe(2);
  });

  it("case 2 — different epoch: server state wins, pending is withheld, consumer notified (fork rule)", () => {
    const srv1 = miniServer("g1");
    const bundle = workAndBundle(srv1);
    bundle.pending = [{ kind: "insertText", clientId: "alice", clientSeq: 2, base: 1, at: { blockId: 1, runId: 2, offset: 5 }, text: "OFFLINE" } as never];

    // Someone re-seeded: a NEW session under a new epoch, different content.
    const srv2 = miniServer("g2");
    const other = new CollabConnection(srv2.attach(), "bob");
    other.join("d");
    other.submit(ins(0, "reseeded"));

    let epochChange: [string, string] | null = null;
    const conn2 = new CollabConnection(srv2.attach(), "alice", {
      onEpochChange: (a, b) => (epochChange = [a, b]),
    });
    conn2.resume(bundle);
    expect(epochChange).toEqual(["g1", "g2"]);
    expect(text(conn2)).toBe("reseededhi"); // took the server's state…
    expect(text(conn2)).not.toContain("OFFLINE"); // …and did NOT merge epochs
    expect(srv2.session.seq).toBe(1); // the old-epoch pending never reached the sequencer
  });
});
