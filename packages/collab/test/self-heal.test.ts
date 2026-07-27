import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type XmlElement } from "@wordinweb/core";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { DocumentSession } from "../src/session.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage, EnvelopeEntry, SealedCheckpoint } from "../src/protocol.js";

/**
 * SELF-HEAL (checkpoint B6a catch-and-resync): the encrypted connection
 * holds both the optimistic replica AND the canonical mirror, so a drifted
 * replica — the "typist ends one character short of canonical, permanently"
 * defect the stress suite reproduces — is detectable LOCALLY at quiescence
 * and repairable by rebuilding the replica from the mirror, with no network
 * round-trip and no user-visible reload beyond the standard docEpoch
 * remount.
 */

function docxBytes(text: string): Uint8Array {
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

/** Minimal blind sequencer (mirrors e2ee-session.test.ts). */
function blindServer(genesisId: string, checkpoint: SealedCheckpoint) {
  let log: EnvelopeEntry[] = [];
  const seen = new Set<string>();
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  const state = { checkpoint };
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        if (msg.t === "hello") {
          peer.deliver({ t: "welcome-enc", docId: "d", genesisId, checkpoint: state.checkpoint, tail: log.filter((e) => e.seq > state.checkpoint.seq), mode: "encrypted" });
        } else if (msg.t === "submit-enc") {
          const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
          let entry = seen.has(key) ? log.find((e) => `${e.clientId}:${e.clientSeq}` === key) : undefined;
          if (!entry) {
            seen.add(key);
            entry = { ...msg.envelope, seq: log.length === 0 ? state.checkpoint.seq + 1 : log[log.length - 1].seq + 1 };
            log.push(entry);
          }
          for (const p of peers) p.deliver({ t: "broadcast-enc", entries: [entry!] });
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return { attach };
}

async function seedEncrypted(text: string, genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  void keys;
  const bytes = docxBytes(text);
  const session = new DocumentSession(DocxDocument.load(bytes));
  const cp = session.checkpoint();
  const k = await deriveEpochKeys(docKey, genesisId);
  const sealed = await sealCheckpoint(k.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: "seed",
  });
  return { checkpoint: { seq: 0, ...sealed } };
}

const flush = (ms = 25) => new Promise((r) => setTimeout(r, ms));
/** Load-immune condition wait: these tests chain REAL-async WebCrypto work
 * behind REAL timers, so a fixed sleep flakes on a busy machine (seen at
 * load avg 50+ under a concurrent playwright run). Poll the condition on a
 * generous budget — a genuine failure still fails, just at the timeout.
 * Fixed flush() remains ONLY for absence assertions (nothing-should-happen
 * windows), which polling cannot express. */
const until = (cond: () => boolean, label: string) =>
  expect.poll(cond, { timeout: 5000, message: label }).toBe(true);
const text = (c: EncryptedCollabConnection): string => {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(c.doc!.docRoot as never);
};
const ins = (at: number, t: string) => ({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: t }) as never;

/** Chop the last character out of the first non-empty w:t IN THE LIVE
 * REPLICA DOC — the exact shape of the B6a drift (optimistic doc one char
 * short of canonical, canonical history intact). */
function corruptLiveDoc(c: EncryptedCollabConnection): void {
  const chop = (el: XmlElement): boolean => {
    if (el.name.endsWith(":t") && el.text.length > 0) {
      el.text = el.text.slice(0, -1);
      return true;
    }
    for (const child of el.children) if (chop(child)) return true;
    return false;
  };
  chop(c.doc!.docRoot);
}

describe("optimistic-replica self-heal (B6a catch-and-resync)", () => {
  it("a drifted replica is detected at quiescence and rebuilt from the mirror", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hello", "g1", docKey);
    const srv = blindServer("g1", checkpoint);

    const heals: { seq: number; liveHash: string; canonicalHash: string }[] = [];
    const alice = new EncryptedCollabConnection(
      srv.attach(), "alice", docKey,
      { onSelfHeal: (info) => heals.push(info) },
      undefined, /*selfCheckDelayMs*/ 10,
    );
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey, {}, undefined, 10);
    alice.join("d");
    bob.join("d");
    await until(() => alice.ready && bob.ready, "join ready");

    bob.submit(ins(0, "A"));
    await until(() => text(alice) === text(bob) && text(alice).includes("A"), "A converges");

    // Simulate the B6a drift: alice's LIVE doc silently loses a character
    // while the canonical history (her mirror, bob, the log) is intact.
    corruptLiveDoc(alice);
    expect(text(alice)).not.toBe(text(bob));

    // Any subsequent broadcast drives alice to quiescence and re-arms the
    // self-check; the heal must fire and restore byte-identity.
    const epochBefore = alice.docEpoch;
    bob.submit(ins(0, "B"));
    await until(() => alice.selfHeals === 1, "heal fires");

    expect(alice.selfHeals).toBe(1);
    expect(heals).toHaveLength(1);
    expect(heals[0].liveHash).not.toBe(heals[0].canonicalHash);
    expect(alice.docEpoch).toBe(epochBefore + 1); // renderer remounts like any reload
    expect(text(alice)).toBe(text(bob));
    expect(serializeXml(alice.doc!.docRoot)).toBe(serializeXml(bob.doc!.docRoot));

    // And the healed replica keeps working: further edits converge.
    bob.submit(ins(0, "C"));
    await until(() => serializeXml(alice.doc!.docRoot) === serializeXml(bob.doc!.docRoot) && text(alice).includes("C"), "post-heal converges");
  });

  it("a LOST submit is resent by the stuck-pending watchdog (server dedup makes resends safe)", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hello", "g3", docKey);
    const srv = blindServer("g3", checkpoint);

    // A lossy attach: swallow the FIRST submit-enc frame (the swarm-observed
    // failure shape — an op lost between pending-tracking and delivery).
    const lossyAttach = () => {
      const inner = srv.attach();
      let dropped = false;
      return {
        send: (msg: ClientMessage) => {
          if (!dropped && (msg as { t: string }).t === "submit-enc") {
            dropped = true;
            return; // lost in transit
          }
          inner.send(msg);
        },
        onMessage: inner.onMessage,
      };
    };

    const alice = new EncryptedCollabConnection(lossyAttach(), "alice", docKey, {}, undefined, 10);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey, {}, undefined, 10);
    alice.join("d");
    bob.join("d");
    await until(() => alice.ready && bob.ready, "join ready");

    alice.submit(ins(0, "LOST?"));
    // The op is pending with nothing on the wire; the watchdog (5× the 10ms
    // debounce) must resend the transformed copy and the room converges.
    await until(
      () => text(bob).includes("LOST?") && serializeXml(alice.doc!.docRoot) === serializeXml(bob.doc!.docRoot),
      "watchdog resend converges",
    );
    expect(text(alice)).toContain("LOST?");
    expect(alice.selfHeals).toBe(0); // recovered by RESEND, not by a heal
  });

  it("undeliverable pending is dropped and healed after the retry budget (convergence over delivery)", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hello", "g4", docKey);
    const srv = blindServer("g4", checkpoint);

    // Swallow EVERY submit-enc from alice — the op can never be delivered.
    const blackholeAttach = () => {
      const inner = srv.attach();
      return {
        send: (msg: ClientMessage) => {
          if ((msg as { t: string }).t === "submit-enc") return;
          inner.send(msg);
        },
        onMessage: inner.onMessage,
      };
    };

    const heals: unknown[] = [];
    const alice = new EncryptedCollabConnection(blackholeAttach(), "alice", docKey, { onSelfHeal: (i) => heals.push(i) }, undefined, 10);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey, {}, undefined, 10);
    alice.join("d");
    bob.join("d");
    await until(() => alice.ready && bob.ready, "join ready");

    alice.submit(ins(0, "DOOMED"));
    // 3 retries × 50ms windows, then drop + heal from the mirror.
    await until(() => alice.selfHeals === 1, "drop+heal fires");
    expect(heals).toHaveLength(1);
    expect(text(alice)).not.toContain("DOOMED"); // honest drop, not silent divergence
    expect(text(alice)).toBe(text(bob));
    // The healed replica keeps working (via the still-lossy transport, bob's
    // edits flow in fine).
    bob.submit(ins(0, "after"));
    await until(() => serializeXml(alice.doc!.docRoot) === serializeXml(bob.doc!.docRoot) && text(alice).includes("after"), "post-heal converges");
  });

  it("a rate-limited burst recovers via the full-queue re-drive (nothing stranded, nothing healed away)", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hello", "g5", docKey);
    const srv = blindServer("g5", checkpoint);

    // A throttling attach: refuse the 3rd..6th submit attempts with
    // rate-limit (mimicking a drained token bucket), accept everything
    // after — so recovery REQUIRES re-driving more than the front op.
    const throttlingAttach = () => {
      const inner = srv.attach();
      let innerDeliver: ((m: import("../src/protocol.js").ServerMessage) => void) | null = null;
      let submits = 0;
      return {
        send: (msg: ClientMessage) => {
          if ((msg as { t: string }).t === "submit-enc") {
            submits++;
            if (submits >= 3 && submits <= 6) {
              innerDeliver?.({ t: "refused", reason: "rate-limit" } as never);
              return;
            }
          }
          inner.send(msg);
        },
        onMessage: (cb: (m: import("../src/protocol.js").ServerMessage) => void) => {
          innerDeliver = cb;
          inner.onMessage(cb);
        },
      };
    };

    const alice = new EncryptedCollabConnection(throttlingAttach(), "alice", docKey, {}, undefined, 10);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey, {}, undefined, 10);
    alice.join("d");
    bob.join("d");
    await until(() => alice.ready && bob.ready, "join ready");

    // A rapid 5-op burst: ops 3-6 (some first-sends AND their early
    // retries) get throttled; the re-drive (backoff 300ms + redelivery +
    // echoes) must deliver the whole chain.
    for (let i = 0; i < 5; i++) alice.submit(ins(0, `${i}`));
    await until(
      () => [0, 1, 2, 3, 4].every((i) => text(bob).includes(`${i}`)) && serializeXml(alice.doc!.docRoot) === serializeXml(bob.doc!.docRoot),
      "re-driven burst converges",
    );
    expect(text(alice)).toBe(text(bob));
    expect(alice.selfHeals).toBe(0); // throttling never escalates to drop+heal
  });

  it("never fires on healthy traffic (no false positives)", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hello", "g2", docKey);
    const srv = blindServer("g2", checkpoint);
    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey, {}, undefined, 10);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey, {}, undefined, 10);
    alice.join("d");
    bob.join("d");
    await until(() => alice.ready && bob.ready, "join ready");

    for (let i = 0; i < 5; i++) {
      alice.submit(ins(0, `a${i}`));
      bob.submit(ins(0, `b${i}`));
      await flush(40);
    }
    // Absence assertion — no poll can express "no heal EVER fires", so wait
    // for full convergence (the precondition for a false positive) and then
    // let several self-check windows pass before reading the counters.
    await until(() => serializeXml(alice.doc!.docRoot) === serializeXml(bob.doc!.docRoot), "churn converges");
    await flush(120); // several self-check windows pass
    expect(alice.selfHeals).toBe(0);
    expect(bob.selfHeals).toBe(0);
    expect(serializeXml(alice.doc!.docRoot)).toBe(serializeXml(bob.doc!.docRoot));
  });
});
