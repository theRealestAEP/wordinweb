import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { DocumentSession } from "../src/session.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage, EnvelopeEntry, SealedCheckpoint } from "../src/protocol.js";

/**
 * UNDO FUZZ: two encrypted clients interleave real edits with undos under a
 * seeded PRNG, and the room must still converge byte-identically.
 *
 * Undo is the operation most likely to break convergence, because it is the
 * only one whose target is a point in the PAST: its inverse is computed
 * against state that concurrent edits have already moved. Every other op
 * addresses the present. So the interesting failures — an inverse applied at
 * stale offsets, a rebase that double-counts, an undo of an already-undone
 * action, two clients undoing at once — only appear when undos are mixed into
 * ordinary traffic, which is exactly what this does.
 *
 * Deterministic by seed: a failure replays exactly, and the seed goes in as a
 * permanent regression case (the demo-fuzz convention).
 */

function docxBytes(text: string): Uint8Array {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`),
  });
}

function blindServer(genesisId: string, checkpoint: SealedCheckpoint) {
  const log: EnvelopeEntry[] = [];
  const seen = new Set<string>();
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        if (msg.t === "hello") {
          peer.deliver({ t: "welcome-enc", docId: "d", genesisId, checkpoint, tail: [...log], mode: "encrypted" });
        } else if (msg.t === "submit-enc") {
          const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
          let entry = seen.has(key) ? log.find((e) => `${e.clientId}:${e.clientSeq}` === key) : undefined;
          if (!entry) {
            seen.add(key);
            entry = { ...msg.envelope, seq: log.length === 0 ? checkpoint.seq + 1 : log[log.length - 1].seq + 1 };
            log.push(entry);
          }
          for (const p of peers) p.deliver({ t: "broadcast-enc", entries: [entry!] });
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return { attach, log: () => log };
}

async function seedEncrypted(text: string, genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const session = new DocumentSession(DocxDocument.load(docxBytes(text)));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx), sidecar: cp.sidecar, docHash: "seed",
  });
  return { checkpoint: { seq: 0, ...sealed } };
}

/** mulberry32 — deterministic, so a failing seed replays exactly. */
function rng(seed: number) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

async function until(cond: () => boolean, label: string): Promise<void> {
  for (let i = 0; i < 400; i++) {
    if (cond()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error(`timed out waiting for ${label}`);
}

const textOf = (c: EncryptedCollabConnection): string => {
  const walk = (e: { name: string; text: string; children: unknown[] }): string =>
    (e.name.endsWith(":t") ? e.text : "") + (e.children as never[]).map(walk).join("");
  return walk(c.doc!.docRoot as never);
};
const xmlOf = (c: EncryptedCollabConnection): string => serializeXml(c.doc!.docRoot);

async function room() {
  const docKey = mintDocKey();
  const { checkpoint } = await seedEncrypted("seed", "g1", docKey);
  const srv = blindServer("g1", checkpoint);
  // A press made while this client's own edits are in flight is HELD for the
  // drain (the undo pending gate). Counting the drops keeps the fuzz's
  // non-vacuity claim true: a held undo that expired never became a real one.
  let expired = 0;
  const cb = { onUndoQueue: (s: { expired: boolean }) => { if (s.expired) expired++; } };
  const a = new EncryptedCollabConnection(srv.attach(), "alice", docKey, cb);
  const b = new EncryptedCollabConnection(srv.attach(), "bob", docKey, cb);
  a.join("d");
  b.join("d");
  await until(() => a.ready && b.ready, "both clients to rehydrate");
  return { a, b, expiredUndos: () => expired };
}

describe("undo fuzz: undos interleaved with edits still converge", () => {
  for (const seed of [1, 7, 42]) {
    it(`seed ${seed}: 40 mixed rounds converge byte-identically`, async () => {
      const rand = rng(seed);
      const { a, b, expiredUndos } = await room();
      const clients = [a, b];
      let undone = 0;
      let queued = 0;
      let declined = 0;

      for (let round = 0; round < 40; round++) {
        const c = clients[rand() < 0.5 ? 0 : 1];
        const roll = rand();
        if (roll < 0.35) {
          // Undo — the operation under test.
          const outcome = c.undoLast();
          if (outcome === "undone") undone++;
          else if (outcome === "queued") queued++; // held for this client's drain
          else declined++;
        } else if (roll < 0.75) {
          const len = textOf(c).length;
          const at = Math.floor(rand() * Math.max(1, len));
          c.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: String.fromCharCode(97 + Math.floor(rand() * 26)) } as never);
        } else {
          const len = textOf(c).length;
          if (len > 1) {
            const start = Math.floor(rand() * (len - 1));
            c.submit({ kind: "deleteText", blockId: 1, runId: 2, start, end: start + 1 } as never);
          }
        }
        // Let a few rounds pile up in flight before draining, so undos land
        // amid unacknowledged traffic rather than in a quiet room.
        if (round % 4 === 3) await new Promise((r) => setTimeout(r, 8));
      }

      // SETTLE FIRST. A held undo submits on its client's drain, so equality
      // observed while one is still waiting says nothing — the room would
      // diverge again the moment it ran.
      await until(
        () => a.pendingCount === 0 && b.pendingCount === 0 && a.undoQueued === 0 && b.undoQueued === 0,
        `seed ${seed}: in-flight edits and held undos to settle`,
      );
      // THE INVARIANT: whatever the interleaving produced, the room agrees.
      await until(() => xmlOf(a) === xmlOf(b), `seed ${seed} to converge`);
      expect(xmlOf(a)).toBe(xmlOf(b));

      // NON-VACUITY: undos actually happened, and at least some were real —
      // otherwise this would be an ordinary edit fuzz wearing an undo label.
      // A held press counts only because none of them expired: every one of
      // them reached the mirror and became a genuine undo attempt.
      expect(undone + declined + queued).toBeGreaterThan(0);
      expect(undone + queued).toBeGreaterThan(0);
      expect(expiredUndos(), "no held undo may be dropped in a healthy room").toBe(0);
    });
  }

  it("both clients undoing simultaneously still converge", async () => {
    // The nastiest shape: two inverses in flight at once, each computed
    // against a state the other is about to change.
    const { a, b } = await room();
    a.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "AAA" } as never);
    await until(() => textOf(b).includes("AAA"), "B to see A's edit");
    b.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "BBB" } as never);
    await until(() => textOf(a).includes("BBB"), "A to see B's edit");

    a.undoLast();
    b.undoLast();
    // Either press may have been HELD for its own client's drain (the undo
    // pending gate) — B's echo in particular is not implied by A having seen
    // the edit. Both must have run before the room can be judged.
    await until(
      () => a.undoQueued === 0 && b.undoQueued === 0 && a.pendingCount === 0 && b.pendingCount === 0,
      "both undos to run and settle",
    );
    await until(() => xmlOf(a) === xmlOf(b), "the room to converge after simultaneous undos");
    expect(xmlOf(a)).toBe(xmlOf(b));
    // Both authors' text is gone; the seed survives.
    expect(textOf(a)).toBe("seed");
  });
});
