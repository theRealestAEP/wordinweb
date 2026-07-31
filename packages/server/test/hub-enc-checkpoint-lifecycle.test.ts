import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { CollabHub, Connection } from "../src/hub.js";
import { DocumentSession } from "@wordinweb/collab/server";
import {
  EncryptedCollabConnection,
  deriveEpochKeys,
  sealCheckpoint,
  bytesToB64,
  mintDocKey,
} from "@wordinweb/collab/client";
import type {
  ClientMessage,
  ServerMessage,
  EnvelopeEntry,
  SealedCheckpoint,
  IntentEnvelope,
} from "@wordinweb/collab/client";

/**
 * Audit of the ENCRYPTED checkpoint lifecycle against the REAL hub
 * (packages/server/src/hub.ts) driven by REAL EncryptedCollabConnection
 * clients. Everything is seeded/deterministic (no Date.now, no Math.random on
 * the intent path). Each `it` states whether it proves an INVARIANT HOLDS
 * (passing guard) or REPRODUCES A BUG (deterministic failure).
 */

// ------------------------------------------------------------------ fixtures
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

/** Produce the sealed seed checkpoint a go-live client would upload (seq 0). */
async function makeSeed(text: string, genesisId: string, docKey: string): Promise<SealedCheckpoint> {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const session = new DocumentSession(DocxDocument.load(docxBytes(text)));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: "seed",
  });
  return { seq: 0, ...sealed } as SealedCheckpoint;
}

const ins = (at: number, t: string) =>
  ({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: t }) as never;

const text = (c: EncryptedCollabConnection): string => {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(c.doc!.docRoot as never);
};

/**
 * A transport that bridges a real EncryptedCollabConnection to the real hub.
 * Client->hub is forwarded immediately (the hub's enc paths are synchronous);
 * hub->client can be GATED to reproduce in-flight / out-of-order races. Every
 * inbound server message is recorded for assertions.
 */
class HubLink {
  private cb: (m: ServerMessage) => void = () => {};
  readonly conn: Connection;
  readonly inbound: ServerMessage[] = [];
  /** Hold client->hub SUBMITS so an envelope can be kept in-flight while the
   * hub adopts a checkpoint past its base — the only way to construct a real
   * stale-base race under a single-threaded hub. */
  holdSubmits = false;
  private heldSubmits: ClientMessage[] = [];
  constructor(private hub: CollabHub, id: string, private pending: Promise<unknown>[]) {
    this.conn = {
      id,
      send: (m: ServerMessage) => {
        this.inbound.push(m);
        this.cb(m);
      },
    };
  }
  send(msg: ClientMessage): void {
    if (this.holdSubmits && msg.t === "submit-enc") {
      this.heldSubmits.push(msg);
      return;
    }
    this.pending.push(this.hub.handle(this.conn, msg));
  }
  onMessage(cb: (m: ServerMessage) => void): void {
    this.cb = cb;
  }
  /** Release held submits into the hub, in order, and stop holding. */
  releaseSubmits(): void {
    this.holdSubmits = false;
    for (const m of this.heldSubmits.splice(0)) this.pending.push(this.hub.handle(this.conn, m));
  }
  refusals(): string[] {
    return this.inbound.filter((m) => m.t === "refused").map((m) => (m as { reason: string }).reason);
  }
  welcomes(): Extract<ServerMessage, { t: "welcome-enc" }>[] {
    return this.inbound.filter((m): m is Extract<ServerMessage, { t: "welcome-enc" }> => m.t === "welcome-enc");
  }
}

/** A test harness owning the hub, the pending-promise pump, and a client
 * factory. `flush` drains client crypto microtasks + hub promises. */
function harness() {
  const pending: Promise<unknown>[] = [];
  const hub = new CollabHub(null); // zero-custody: encrypted docs are magic-link tier
  const links: HubLink[] = [];
  let n = 0;
  const refusals: { who: string; reason: string }[] = [];
  const mkClient = (clientId: string, docKey: string) => {
    const link = new HubLink(hub, `sock-${n++}`, pending);
    links.push(link);
    const conn = new EncryptedCollabConnection(link as never, clientId, docKey, {
      onRefused: (reason) => refusals.push({ who: clientId, reason }),
    });
    return { conn, link };
  };
  const flush = async (rounds = 10) => {
    for (let i = 0; i < rounds; i++) {
      await new Promise((r) => setTimeout(r, 5));
      await Promise.all(pending.splice(0));
    }
  };
  return { hub, mkClient, flush, refusals, links };
}

// ============================================================ Scenario 1
describe("Scenario 1: crossing MULTIPLE checkpoints (>=3 adoptions), 2 clients", () => {
  it("converges byte-identically with no client refusal across 3 checkpoint adoptions", async () => {
    const h = harness();
    const docKey = mintDocKey();
    const seed = await makeSeed("seed ", "g1", docKey);
    expect(h.hub.seedEncrypted("d", "g1", seed).ok).toBe(true);

    const { conn: alice } = h.mkClient("alice", docKey); // first joiner = checkpointer
    alice.join("d");
    await h.flush();
    const { conn: bob } = h.mkClient("bob", docKey);
    bob.join("d");
    await h.flush();
    expect(alice.ready && bob.ready).toBe(true);

    // 160 alternating edits → seqs 1..160; checkpoints fire at 50/100/150.
    // A late joiner probes the adopted checkpoint at three points to PROVE
    // three distinct adoptions (checkpoint.seq strictly climbs 50→100→150).
    const adopted: number[] = [];
    const probe = async (label: string) => {
      const { conn: probe, link } = h.mkClient(`probe-${label}`, docKey);
      probe.join("d");
      await h.flush();
      const w = link.welcomes()[0];
      adopted.push(w.checkpoint.seq);
      // The probe must reconstruct the SAME doc from checkpoint + tail.
      expect(serializeXml(probe.doc!.docRoot)).toBe(serializeXml(alice.doc!.docRoot));
      return w;
    };

    for (let i = 1; i <= 160; i++) {
      const who = i % 2 === 0 ? bob : alice;
      who.submit(ins(0, "x"));
      await h.flush(4);
      if (i === 60) await probe("a"); // checkpoint 50 should be adopted
      if (i === 110) await probe("b"); // checkpoint 100 should be adopted
      if (i === 160) {
        const w = await probe("c"); // checkpoint 150 should be adopted
        expect(w.tail.length).toBeLessThanOrEqual(12); // short tail => pruned
      }
    }

    // >=3 distinct adoptions, strictly increasing.
    expect(adopted[0]).toBeGreaterThanOrEqual(50);
    expect(adopted[1]).toBeGreaterThan(adopted[0]);
    expect(adopted[2]).toBeGreaterThan(adopted[1]);
    expect(adopted[2]).toBeGreaterThanOrEqual(150);

    // Byte-identical convergence and ZERO refusals (no sequence-desync, no
    // divergence gossip) throughout the whole multi-checkpoint run.
    expect(serializeXml(alice.doc!.docRoot)).toBe(serializeXml(bob.doc!.docRoot));
    expect(h.refusals).toEqual([]);
  }, 60000);
});

// ============================================================ Scenario 2
describe("Scenario 2: late joiner mid-session right after a checkpoint prune", () => {
  it("a 3rd client joining after the prune reconstructs the same doc and keeps numbering", async () => {
    const h = harness();
    const docKey = mintDocKey();
    const seed = await makeSeed("seed ", "g1", docKey);
    h.hub.seedEncrypted("d", "g1", seed);

    const { conn: alice } = h.mkClient("alice", docKey);
    alice.join("d");
    await h.flush();
    const { conn: bob } = h.mkClient("bob", docKey);
    bob.join("d");
    await h.flush();

    // 52 edits so a checkpoint at 50 is adopted and the log is pruned to a
    // short post-checkpoint tail (seqs 51,52).
    for (let i = 1; i <= 52; i++) {
      (i % 2 === 0 ? bob : alice).submit(ins(0, "y"));
      await h.flush(4);
    }

    // Carol joins right after the prune.
    const { conn: carol, link: carolLink } = h.mkClient("carol", docKey);
    carol.join("d");
    await h.flush();
    const w = carolLink.welcomes()[0];
    expect(w.checkpoint.seq).toBeGreaterThanOrEqual(50); // rehydrated from the ADOPTED cp
    expect(w.tail.length).toBeLessThanOrEqual(4); // pruned tail

    expect(serializeXml(carol.doc!.docRoot)).toBe(serializeXml(alice.doc!.docRoot));

    // Carol edits immediately: her entry must continue numbering from the
    // checkpoint (>52), not restart — all three converge.
    carol.submit(ins(0, "C"));
    await h.flush();
    alice.submit(ins(0, "A"));
    await h.flush();
    expect(text(alice)).toBe(text(carol));
    expect(text(bob)).toBe(text(carol));
    expect(text(carol)).toContain("C");
    expect(text(carol)).toContain("A");
    expect(h.refusals).toEqual([]);
  }, 40000);
});

// ============================================================ Scenario 3
describe("Scenario 3: checkpoint adopted while a stale-base submit is in flight", () => {
  // ── REGRESSION GUARD (was a reproduced, deterministic divergence): the
  //    stale-base retry in enc-connection.ts used to reseal `this.lastSent` —
  //    the ORIGINAL, UNTRANSFORMED intent — with a bumped base. But the
  //    optimistic replica had ALREADY transformed its pending copy of that
  //    intent against the intervening confirmed edit, so the hub's mirror
  //    sequenced the untransformed body verbatim (its base excluded that
  //    edit → no transform) while the retrying client's own-echo FAST PATH
  //    baked the transformed optimistic position into its confirmed baseline:
  //    the retrying client permanently, silently diverged from everyone else.
  //    Reachable in honest, in-order use whenever a checkpoint was adopted
  //    past an in-flight submit whose author was concurrently editing the
  //    same region (the CHECKPOINT_EVERY quiescent boundary).
  //
  //    FIX: the retry now reseals the replica's PENDING (already-transformed)
  //    copy — replica.pendingIntent(clientId, clientSeq) — so the canonical
  //    apply matches the optimistic doc. This test drives the exact race and
  //    asserts full convergence.
  it("a stale-base retry resubmits the TRANSFORMED pending body and converges with the canonical", async () => {
    const h = harness();
    const docKey = mintDocKey();
    const seed = await makeSeed("seed ", "g1", docKey);
    h.hub.seedEncrypted("d", "g1", seed);

    const { conn: alice } = h.mkClient("alice", docKey); // first joiner = checkpointer
    alice.join("d");
    await h.flush();
    const { conn: bob, link: bobLink } = h.mkClient("bob", docKey);
    bob.join("d");
    await h.flush();

    // Drive alice to seq 49 with inserts at offset 0 (bob ingests them all, so
    // bob's confirmedSeq tracks to 49).
    for (let i = 1; i <= 49; i++) {
      alice.submit(ins(0, "a"));
      await h.flush(4);
    }

    // HOLD bob's client->hub submit so his optimistic edit is sealed at base 49
    // and kept in-flight while the hub adopts a checkpoint past it. bob's
    // INBOUND stays open — exactly what an in-order transport delivers: bob
    // sees alice's seq-50 edit (and transforms his pending B against it) BEFORE
    // the stale-base refusal arrives.
    bobLink.holdSubmits = true;
    bob.submit(ins(0, "B")); // offset 0 at base 49; buffered en route to the hub
    await h.flush(4);

    // Alice's 50th edit → seq 50 (insert 'a' at offset 0, concurrent with bob's
    // held B). She ingests it (50 % 50 === 0) and uploads a checkpoint at 50;
    // the log head is exactly 50 (bob's submit is held), so the hub adopts it
    // quiescently and prunes → checkpoint.seq = 50. bob ingests seq 50 too and,
    // because his B is still pending at base 49, the replica SLOW PATH
    // transforms B against alice's concurrent insert → B moves to offset 1.
    alice.submit(ins(0, "a")); // seq 50
    await h.flush(8);
    const { conn: probe, link: probeLink } = h.mkClient("probe", docKey);
    probe.join("d");
    await h.flush();
    expect(probeLink.welcomes()[0].checkpoint.seq).toBe(50); // adoption confirmed

    // Release bob's held submit (base 49) → base 49 < checkpoint.seq 50 →
    // STALE-BASE. enc-connection reseals the replica's PENDING copy of B —
    // already transformed against alice's concurrent seq-50 insert (offset 1)
    // — with base = confirmedSeq (50) and resubmits. The hub sequences it;
    // canonical and optimistic agree on B at offset 1.
    bobLink.releaseSubmits();
    await h.flush(12);

    // The stale-base path was genuinely exercised.
    expect(bobLink.refusals()).toContain("stale-base");

    // CANONICAL (alice + a fresh probe) and bob all agree: the transformed
    // position ("aB…"), byte-identical everywhere. Before the fix bob showed
    // "aB…" while alice/probe showed "Ba…" — permanent silent divergence.
    expect(text(alice).startsWith("aB")).toBe(true);
    expect(text(bob)).toBe(text(alice));
    expect(text(probe)).toBe(text(alice)); // the probe re-derived from checkpoint+tail
    expect(serializeXml(bob.doc!.docRoot)).toBe(serializeXml(alice.doc!.docRoot));
  }, 40000);
});

// ============================================================ Scenario 4
describe("Scenario 4: dedup after prune (double-sequence, hardened)", () => {
  // ── THE FOOTGUN (found by reachability probing, now FIXED in hub.ts).
  //    submit-enc dedups on `seen.has(key)` and then looks the entry up with
  //    `log.find(...)`. Those two structures do NOT have the same lifetime: a
  //    quiescent checkpoint adoption bakes entries into the sealed bytes and
  //    EMPTIES the log, while `seen` keeps every key forever. So for a pruned
  //    key `seen.has` is true and `log.find` is undefined.
  //
  //    ATTACK/BUG SHAPE: resubmit an already-sequenced (clientId, clientSeq)
  //    with its base BUMPED to >= checkpoint.seq. The bump is what makes it
  //    reachable — it clears the stale-base guard, which is what catches the
  //    naive same-base resend. The hub then read the missing `log.find` as
  //    "brand new" and minted a SECOND seq for a key that already owns one.
  //    Blast radius is the whole room, not the sender: every honest mirror
  //    re-derives seqs in arrival order and dedups by the same key, so the
  //    re-broadcast fails its `entry.seq === env.seq` check and hard-refuses
  //    `sequence-desync` (last test below pins that consequence). One buggy or
  //    hostile client could therefore take down every other participant.
  //
  //    FIX (hub.ts submit-enc, pruned-key branch): a key in `seen` but not in
  //    the log is treated as ALREADY SEQUENCED — idempotent no-op, no new seq,
  //    no fan-out. Same observable outcome as the normal-duplicate path (whose
  //    re-broadcast is a dedup no-op at every mirror) minus the wire traffic.
  //    These tests pin the hardened behavior; the comments above keep the
  //    reachability documentation.

  const raw = (
    clientId: string,
    clientSeq: number,
    base: number,
    ct = "Y2lwaGVy",
  ): IntentEnvelope => ({ clientId, clientSeq, base, iv: "aXY=", ciphertext: ct });

  it("INVARIANT: an honest resend of a pruned entry is caught by the stale-base guard (no double-sequence)", async () => {
    // Build via a real hub with real hello wiring (import protocol constants).
    const { PROTOCOL_VERSION, ENGINE_VERSION } = await import("@wordinweb/collab/server");
    const helloMsg = (clientId: string) =>
      ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, engineVersion: ENGINE_VERSION }) as never;
    const hub = new CollabHub(null);
    const CP = { seq: 0, iv: "aXY=", ciphertext: "c2VlZA==" };
    hub.seedEncrypted("d", "g1", CP as never);

    const rec: ServerMessage[] = [];
    const a: Connection = { id: "a", send: (m) => rec.push(m) };
    await hub.handle(a, helloMsg("alice")); // first joiner = checkpointer

    // clientSeq 5 submitted base 0 → seq 1, enters log + seen.
    await hub.handle(a, { t: "submit-enc", envelope: raw("alice", 5, 0) });
    // Adopt a quiescent checkpoint at the log head (seq 1) → prune log, seen keeps "alice:5".
    await hub.handle(a, { t: "checkpoint", checkpoint: { seq: 1, iv: "aXY=", ciphertext: "eA==" } });

    // A NAIVE resend reuses the ORIGINAL base (0). base 0 < checkpoint.seq 1 →
    // stale-base, refused BEFORE the dedup/log.find branch. No new entry.
    rec.length = 0;
    await hub.handle(a, { t: "submit-enc", envelope: raw("alice", 5, 0) });
    const last = rec[rec.length - 1];
    expect(last).toEqual({ t: "refused", reason: "stale-base" });
    // Proof no double-sequence happened: no broadcast-enc emitted.
    expect(rec.some((m) => m.t === "broadcast-enc")).toBe(false);
  });

  it("HARDENED: a resend with a bumped base >= checkpoint.seq hits the pruned key and is a NO-OP (no second seq)", async () => {
    // The reachability probe, now asserting the fix. The bumped base clears the
    // stale-base guard and lands squarely on the `seen.has && !log.find` branch
    // — the branch that used to mint a second seq. It must now sequence
    // nothing, broadcast nothing, and leave the numbering untouched so the
    // NEXT genuine intent gets the seq the double-sequence would have stolen.
    const { PROTOCOL_VERSION, ENGINE_VERSION } = await import("@wordinweb/collab/server");
    const helloMsg = (clientId: string) =>
      ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, engineVersion: ENGINE_VERSION }) as never;
    const hub = new CollabHub(null);
    const CP = { seq: 0, iv: "aXY=", ciphertext: "c2VlZA==" };
    hub.seedEncrypted("d", "g1", CP as never);
    const rec: ServerMessage[] = [];
    const a: Connection = { id: "a", send: (m) => rec.push(m) };
    await hub.handle(a, helloMsg("alice"));

    // clientSeq 5 @ base 0 → seq 1 (in log + seen).
    await hub.handle(a, { t: "submit-enc", envelope: raw("alice", 5, 0) });
    const firstSeq = (rec.find((m) => m.t === "broadcast-enc") as never as { entries: { seq: number }[] }).entries[0].seq;
    expect(firstSeq).toBe(1);
    // Adopt checkpoint at 1 → prune. seen still has "alice:5"; log is empty.
    await hub.handle(a, { t: "checkpoint", checkpoint: { seq: 1, iv: "aXY=", ciphertext: "eA==" } });

    // Resend clientSeq 5 with base bumped to 1 (>= checkpoint.seq) — passes the
    // stale-base guard. seen.has("alice:5") is true, log.find returns undefined
    // (pruned) → the hardened hub treats it as already sequenced.
    rec.length = 0;
    await hub.handle(a, { t: "submit-enc", envelope: raw("alice", 5, 1) });
    // No second seq, and nothing at all on the wire — not a broadcast (which
    // would carry the bogus seq) and not a refusal (an honest reconnect must
    // not be told its own already-sequenced edit failed).
    expect(rec).toEqual([]);

    // Numbering is intact: the next GENUINE intent takes seq 2, the number the
    // double-sequence would have consumed.
    await hub.handle(a, { t: "submit-enc", envelope: raw("alice", 6, 1) });
    const next = rec.find((m) => m.t === "broadcast-enc") as never as { entries: EnvelopeEntry[] };
    expect(next.entries[0].seq).toBe(2);
    expect(next.entries[0].clientSeq).toBe(6);
    expect(firstSeq).toBe(1); // and the original key still owns exactly one seq
  });

  it("HARDENED (live): a pruned-key resend from a real client leaves every mirror quiet — no sequence-desync", async () => {
    // End-to-end version of the guard against REAL EncryptedCollabConnection
    // clients: drive a genuine checkpoint adoption, then replay one of the
    // pruned envelopes verbatim with a bumped base (what a buggy reconnecting
    // client does). Nobody may observe a new seq, and — the property that
    // matters — no honest mirror may hard-refuse `sequence-desync`.
    const h = harness();
    const docKey = mintDocKey();
    const seed = await makeSeed("seed ", "g1", docKey);
    h.hub.seedEncrypted("d", "g1", seed);

    const { conn: alice, link: aliceLink } = h.mkClient("alice", docKey); // checkpointer
    alice.join("d");
    await h.flush();
    const { conn: bob } = h.mkClient("bob", docKey);
    bob.join("d");
    await h.flush();

    // 50 edits → seqs 1..50; alice checkpoints at 50 and the hub adopts it
    // quiescently, pruning every entry (log empty, `seen` keeps all 50 keys).
    for (let i = 1; i <= 50; i++) {
      (i % 2 === 0 ? bob : alice).submit(ins(0, "p"));
      await h.flush(4);
    }
    const before = text(alice);

    // Adoption confirmed: a joiner is served from a checkpoint at 50 with an
    // empty tail, so seqs 1..50 now live ONLY in `seen`.
    const { conn: probe, link: probeLink } = h.mkClient("probe", docKey);
    probe.join("d");
    await h.flush();
    expect(probeLink.welcomes()[0].checkpoint.seq).toBe(50);
    expect(probeLink.welcomes()[0].tail.length).toBe(0);

    // Replay one of ALICE's own pruned envelopes (her real sealed bytes, so it
    // would decrypt if it were ever sequenced) at base = checkpoint.seq.
    const mine = aliceLink.inbound
      .filter((m): m is Extract<ServerMessage, { t: "broadcast-enc" }> => m.t === "broadcast-enc")
      .flatMap((m) => m.entries)
      .find((e) => e.clientId === "alice")!;
    const inboundBefore = aliceLink.inbound.length;
    aliceLink.send({
      t: "submit-enc",
      envelope: { clientId: mine.clientId, clientSeq: mine.clientSeq, base: 50, iv: mine.iv, ciphertext: mine.ciphertext },
    } as ClientMessage);
    await h.flush(6);

    // Silent no-op: alice received nothing back, and no client refused.
    expect(aliceLink.inbound.length).toBe(inboundBefore);
    expect(h.refusals).toEqual([]);
    expect(text(alice)).toBe(before);
    expect(text(bob)).toBe(before);
    expect(text(probe)).toBe(before);

    // The room is still healthy: a fresh joiner rebuilds the same document
    // from checkpoint + tail, and live editing continues to converge.
    const { conn: carol } = h.mkClient("carol", docKey);
    carol.join("d");
    await h.flush();
    expect(text(carol)).toBe(before);
    alice.submit(ins(0, "A"));
    await h.flush(6);
    expect(text(alice)).toContain("A");
    expect(text(bob)).toBe(text(alice));
    expect(text(carol)).toBe(text(alice));
    expect(text(probe)).toBe(text(alice));
    expect(h.refusals).toEqual([]);
  }, 40000);

  it("WHY THE GUARD EXISTS: a double-sequenced re-broadcast makes a live mirror hard-refuse `sequence-desync`", async () => {
    // The blast radius the fix prevents, pinned by INJECTING the frame the
    // unhardened hub used to emit (the hub can no longer produce it, so this
    // fabricates it on the wire). A real EncryptedCollabConnection whose mirror
    // already ingested the ORIGINAL seq of clientSeq K receives the SAME
    // clientSeq under a NEW seq, and its `entry.seq !== env.seq` check fires.
    const h = harness();
    const docKey = mintDocKey();
    const seed = await makeSeed("seed ", "g1", docKey);
    h.hub.seedEncrypted("d", "g1", seed);
    const { conn: alice, link } = h.mkClient("alice", docKey);
    alice.join("d");
    await h.flush();

    // alice submits clientSeq 1 → real seq 1; her mirror records "alice:1"@1.
    alice.submit(ins(0, "Z"));
    await h.flush();
    const firstBc = link.inbound.find((m) => m.t === "broadcast-enc") as never as { entries: EnvelopeEntry[] };
    const orig = firstBc.entries[0];
    expect(orig.seq).toBe(1);

    // Fabricate a re-broadcast: the SAME sealed envelope (same clientId:clientSeq,
    // openable by alice) but stamped with a DIFFERENT seq (2) — exactly what the
    // double-sequence bug puts on the wire. Deliver it straight to alice's mirror.
    const dup: EnvelopeEntry = { ...orig, seq: 2 };
    link.conn.send({ t: "broadcast-enc", entries: [dup] } as ServerMessage);
    await h.flush();

    // The mirror dedups clientSeq 1 → prior entry seq 1, but env.seq is 2 →
    // `sequence-desync` surfaces on the honest client.
    expect(h.refusals.some((r) => r.reason === "sequence-desync")).toBe(true);
  }, 20000);
});
