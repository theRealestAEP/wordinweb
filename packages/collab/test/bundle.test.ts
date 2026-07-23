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

  it("a sustained burst coalesces to one leading + one trailing write — never a deferred-forever debounce", async () => {
    const { conn, store, p, advance } = harness();
    const settle = () => new Promise((r) => setTimeout(r, 20)); // write() awaits a real async digest
    conn.submit(ins(0, "a"));
    p.notify(); // leading edge: slot claimed synchronously, write enqueued
    await settle(); // let the (real-async) leading write's chain drain
    expect(store.writes).toBe(1);
    for (let i = 1; i <= 30; i++) { conn.submit(ins(i, "x")); p.notify(); advance(20); }
    await settle();
    // 600ms of continuous typing: a debounce would still be waiting; the
    // throttle has armed exactly one trailing write inside the window.
    expect(store.writes).toBe(1);
    advance(1000); // window closes → the trailing write fires
    await settle();
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

describe("fuzz teardown round (doc 12 §10): seeded bursts → server death → re-seed → convergence", () => {
  /** mulberry32 — same seeded PRNG discipline as the react fuzz harness. */
  function rng(seed: number) {
    let a = seed >>> 0;
    return () => {
      a |= 0; a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  for (const seed of [11, 23, 47]) {
    it(`seed ${seed}: pre-kill state survives via a bundle; all parties converge in the new epoch`, () => {
      const rand = rng(seed);
      // Epoch 1: alice + bob trade random drained bursts (the supported
      // one-in-flight discipline — each submit echoes synchronously here).
      const srv1 = miniServer("g_epoch1");
      const alice = new CollabConnection(srv1.attach(), "alice");
      const bob = new CollabConnection(srv1.attach(), "bob");
      alice.join("d");
      bob.join("d");
      const who = [alice, bob];
      for (let i = 0; i < 40; i++) {
        const c = who[Math.floor(rand() * 2)];
        const len = text(c).length;
        c.submit(ins(Math.floor(rand() * (len + 1)), String.fromCharCode(97 + Math.floor(rand() * 26))));
      }
      expect(text(alice)).toBe(text(bob)); // converged pre-kill
      const preKill = Buffer.from(alice.doc!.save());
      const aliceBundle = alice.exportBundle("d")!;
      const bobBundle = bob.exportBundle("d")!;

      // Server dies. Zero custody: srv2 starts knowing NOTHING; alice's
      // bundle re-seeds a NEW epoch (doc 12 §5.3 — the browsers are the
      // recovery machinery).
      const srv2 = new DocumentSession(DocxDocument.load(aliceBundle.confirmedBytes));
      srv2.installSidecar(aliceBundle.confirmedSidecar);
      const peers2: { deliver: (m: ServerMessage) => void }[] = [];
      const attach2 = () => {
        const peer = { deliver: (_m: ServerMessage) => {} };
        peers2.push(peer);
        return {
          send: (msg: ClientMessage) => {
            if (msg.t === "hello") {
              const cp = srv2.checkpoint();
              peer.deliver({ t: "welcome", docId: "d", seq: cp.seq, snapshot: Buffer.from(cp.docx).toString("base64"),
                sidecar: cp.sidecar, tail: srv2.entriesSince(cp.seq), genesisId: "g_epoch2", mode: "plaintext" });
            } else if (msg.t === "submit") {
              const entry = srv2.submit(msg.intent);
              for (const p of peers2) p.deliver({ t: "broadcast", entries: [entry] });
            }
          },
          onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
        };
      };

      // The revived session equals the pre-kill confirmed state byte-for-byte.
      expect(Buffer.from(srv2.checkpoint().docx).equals(preKill)).toBe(true);

      // Alice rejoins her own re-seed (epoch changed for her too — fine);
      // bob resumes with his old-epoch bundle and lands in case 2.
      let bobForked = false;
      const alice2 = new CollabConnection(attach2(), "alice");
      alice2.resume(aliceBundle);
      const bob2 = new CollabConnection(attach2(), "bob", { onEpochChange: () => (bobForked = true) });
      bob2.resume(bobBundle);
      expect(bobForked).toBe(true);
      expect(Buffer.from(bob2.doc!.save()).equals(preKill)).toBe(true); // same content — nothing lost

      // Post-revival editing converges across both + the server.
      for (let i = 0; i < 20; i++) {
        const c = [alice2, bob2][Math.floor(rand() * 2)];
        const len = text(c).length;
        c.submit(ins(Math.floor(rand() * (len + 1)), String.fromCharCode(97 + Math.floor(rand() * 26))));
      }
      expect(text(alice2)).toBe(text(bob2));
      expect(Buffer.from(alice2.doc!.save()).equals(Buffer.from(srv2.checkpoint().docx))).toBe(true);
    });
  }
});

describe("lineage + fast-forward (doc 15 phase 1)", () => {
  /** sha256 hex of bytes (what the persister records per head). */
  async function hashOf(bytes: Uint8Array): Promise<string> {
    const d = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
    let hex = "";
    for (const b of new Uint8Array(d)) hex += b.toString(16).padStart(2, "0");
    return hex;
  }
  /** miniServer whose welcome carries a lineage chain (like a real seed). */
  function lineageServer(genesisId: string, lineage: { genesisId: string; seq: number; docHash: string }[]) {
    const srv = miniServer(genesisId);
    const origAttach = srv.attach;
    return {
      ...srv,
      attach: () => {
        const t = origAttach();
        return {
          send: t.send,
          onMessage: (cb: (m: ServerMessage) => void) =>
            t.onMessage((m) => cb(m.t === "welcome" ? { ...m, lineage } : m)),
        };
      },
    };
  }

  it("a strict descendant fast-forwards: no draft, no fork callback, superseded state bankable", async () => {
    // Epoch 1: alice works; her bundle records head (g1, seq 1, hash).
    const srv1 = miniServer("g1");
    const conn1 = new CollabConnection(srv1.attach(), "alice");
    conn1.join("d");
    conn1.submit(ins(0, "hello"));
    const bundle = conn1.exportBundle("d")!;
    bundle.lineage = [{ genesisId: "g1", seq: bundle.confirmedSeq, docHash: await hashOf(bundle.confirmedBytes) }];

    // Epoch 2 seeded by someone whose chain CONTAINS alice's exact head
    // (they carried on from her state — she is a strict ancestor).
    const srv2 = lineageServer("g2", [
      ...bundle.lineage,
      { genesisId: "g2", seq: 0, docHash: "whatever-newer" },
    ]);
    let fastForwarded: [string, string] | null = null;
    let forked = false;
    const conn2 = new CollabConnection(srv2.attach(), "alice", {
      onFastForward: (a, b) => (fastForwarded = [a, b]),
      onEpochChange: () => (forked = true),
    });
    conn2.resume(bundle);
    expect(fastForwarded).toEqual(["g1", "g2"]);
    expect(forked).toBe(false); // silent adoption — the doc-15 churn fix
    expect(conn2.ready).toBe(true);
  });

  it("a FABRICATED lineage (right epoch id, wrong hash) does NOT fast-forward — falls to the fork path", async () => {
    const srv1 = miniServer("g1");
    const conn1 = new CollabConnection(srv1.attach(), "alice");
    conn1.join("d");
    conn1.submit(ins(0, "hello"));
    const bundle = conn1.exportBundle("d")!;
    bundle.lineage = [{ genesisId: "g1", seq: bundle.confirmedSeq, docHash: await hashOf(bundle.confirmedBytes) }];

    // Mallory claims ancestry with alice's epoch id but not her real hash.
    const srv2 = lineageServer("g2", [{ genesisId: "g1", seq: 99, docHash: "forged" }]);
    let fastForwarded = false;
    let forked = false;
    const conn2 = new CollabConnection(srv2.attach(), "alice", {
      onFastForward: () => (fastForwarded = true),
      onEpochChange: () => (forked = true),
    });
    conn2.resume(bundle);
    expect(fastForwarded).toBe(false);
    expect(forked).toBe(true); // treated as divergence → draft preserved path
  });

  it("no lineage in the welcome (legacy/provider rooms) ⇒ every epoch change is a fork (safe default)", () => {
    const srv1 = miniServer("g1");
    const conn1 = new CollabConnection(srv1.attach(), "alice");
    conn1.join("d");
    conn1.submit(ins(0, "hello"));
    const bundle = conn1.exportBundle("d")!;
    bundle.lineage = [{ genesisId: "g1", seq: 1, docHash: "x" }];
    const srv2 = miniServer("g2"); // welcome carries no lineage
    let forked = false;
    const conn2 = new CollabConnection(srv2.attach(), "alice", { onEpochChange: () => (forked = true) });
    conn2.resume(bundle);
    expect(forked).toBe(true);
  });
});
