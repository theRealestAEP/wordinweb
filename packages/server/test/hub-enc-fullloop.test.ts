import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type XmlElement } from "@wordinweb/core";
import { CollabHub, type Connection } from "../src/hub.js";
import { EncryptedCollabConnection, bytesToB64, ClientReplica } from "@wordinweb/collab/client";
import {
  DocumentSession,
  deriveEpochKeys,
  mintDocKey,
  sealCheckpoint,
  type ServerMessage,
  type ClientMessage,
} from "@wordinweb/collab/server";

/**
 * REAL full-loop encrypted convergence harness (team-lead task): the REAL
 * `CollabHub` (with its CHECKPOINT_EVERY=50 prune cycle) wired to TWO REAL
 * `EncryptedCollabConnection` clients over an in-process loopback. No
 * in-test `blindServer`, no plaintext shortcut — this is the exact path the
 * product uses in a two-window E2EE session.
 *
 * The point is to drive the collab layer the way a real editor does (type,
 * press Enter → splitParagraph; fast bursts → multiple pending ops) long
 * enough to cross a real checkpoint, and assert byte-identical convergence
 * against a server-ordered canonical oracle (a late joiner, which carries
 * NO optimistic history — its doc IS the canonical replay).
 */

// ---------------------------------------------------------------------------
// Seed doc + encrypted seeding (mirrors packages/collab/test/e2ee-session).
// ---------------------------------------------------------------------------

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

async function seedEncrypted(text: string, genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const bytes = docxBytes(text);
  const session = new DocumentSession(DocxDocument.load(bytes));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: "seed",
  });
  return { seq: 0, ...sealed };
}

// ---------------------------------------------------------------------------
// In-process loopback: REAL hub <-> REAL encrypted clients.
//
// A client `transport.send` schedules `hub.handle(conn, msg)`; the hub's
// `conn.send` delivers synchronously to the client's message callback (which
// enqueues async decrypt/ingest on the client's own serial queue). `settle`
// drains the hub-handle promises and yields the event loop enough turns for
// the clients' WebCrypto queues to run to completion.
// ---------------------------------------------------------------------------

class Bus {
  pending: Promise<void>[] = [];
  private cbs = new Map<string, (m: ServerMessage) => void>();
  constructor(private hub: CollabHub) {}

  transportFor(connId: string) {
    const conn: Connection = {
      id: connId,
      send: (m: ServerMessage) => {
        this.cbs.get(connId)?.(m);
      },
    };
    return {
      send: (m: ClientMessage) => {
        this.pending.push(this.hub.handle(conn, m));
      },
      onMessage: (cb: (m: ServerMessage) => void) => this.cbs.set(connId, cb),
    };
  }
}

/** Yield the event loop until the hub-handle queue AND the clients' internal
 * crypto queues quiesce. Deterministic: no content depends on timing, only
 * the async plumbing settling. */
async function settle(bus: Bus, rounds = 40): Promise<void> {
  // Unconditional yields: the clients' welcome/ingest work runs on their own
  // WebCrypto promise queues and does not always push to `bus.pending`, so we
  // cannot early-exit on an empty pending list. A fixed turn count is
  // deterministic and each turn is a real setTimeout(0) macrotask, which
  // flushes all pending microtasks (incl. crypto) between turns.
  for (let i = 0; i < rounds; i++) {
    const p = bus.pending;
    bus.pending = [];
    if (p.length) await Promise.all(p);
    await new Promise((r) => setTimeout(r, 0));
  }
  // Load-robustness: under full-suite CPU contention the crypto queues can
  // outlast the fixed rounds (the bus goes quiet before the last decrypt
  // lands and re-fills it). Keep yielding until the bus stays empty for a
  // stretch of consecutive turns, bounded so a genuine hang still fails.
  let quiet = 0;
  for (let i = 0; i < rounds * 10 && quiet < 12; i++) {
    const p = bus.pending;
    bus.pending = [];
    if (p.length) {
      quiet = 0;
      await Promise.all(p);
    } else {
      quiet++;
    }
    await new Promise((r) => setTimeout(r, 0));
  }
}

// ---------------------------------------------------------------------------
// Deterministic driving helpers.
// ---------------------------------------------------------------------------

/** Tiny deterministic PRNG (mulberry32) — no Math.random anywhere. */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

interface RunAddr {
  blockId: number;
  runId: number;
  len: number;
}

/** Walk a connection's LIVE (optimistic) doc and return every addressable
 * (paragraph, first-run, text-length) triple — the exact state the real
 * editor would address a keystroke against. */
function addrs(conn: EncryptedCollabConnection): RunAddr[] {
  const doc = conn.doc as unknown as DocxDocument | null;
  if (!doc) return [];
  const ids = doc.stableIds!;
  const out: RunAddr[] = [];
  const visit = (el: XmlElement): void => {
    if (localName(el.name) === "p") {
      const runEl = el.children.find((c) => localName(c.name) === "r");
      if (runEl) {
        const blockId = ids.idOf(el);
        const runId = ids.idOf(runEl);
        if (blockId !== undefined && runId !== undefined) {
          const tEl = runEl.children.find((c) => localName(c.name) === "t");
          out.push({ blockId, runId, len: tEl?.text.length ?? 0 });
        }
      }
    }
    for (const c of el.children) visit(c);
  };
  for (const root of doc.editableRoots()) visit(root);
  return out;
}

function typeChar(conn: EncryptedCollabConnection, a: RunAddr, off: number): void {
  conn.submit({ kind: "insertText", at: { blockId: a.blockId, runId: a.runId, offset: off }, text: "x" } as never);
}

function pressEnter(conn: EncryptedCollabConnection, a: RunAddr, off: number): void {
  const [nb, nr] = conn.allocIds(2);
  conn.submit({ kind: "splitParagraph", at: { blockId: a.blockId, runId: a.runId, offset: off }, newBlockId: nb, newRunId: nr } as never);
}

/** Plaintext content of a client's live doc (for divergence readouts). */
function textOf(conn: EncryptedCollabConnection): string {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk((conn.doc as never as { docRoot: never }).docRoot);
}

function xmlOf(conn: EncryptedCollabConnection): string {
  return serializeXml((conn.doc as never as { docRoot: XmlElement }).docRoot);
}

/** Wire a client with refusal capture. */
function makeClient(bus: Bus, hub: CollabHub, connId: string, clientId: string, docKey: string) {
  const refusals: string[] = [];
  const conn = new EncryptedCollabConnection(bus.transportFor(connId), clientId, docKey, {
    onRefused: (r) => refusals.push(r),
  });
  return { conn, refusals };
}

async function joinFresh(bus: Bus, docKey: string, connId: string, clientId: string) {
  const c = makeClient(bus, undefined as never, connId, clientId, docKey);
  c.conn.join("d");
  await settle(bus);
  return c;
}

// ---------------------------------------------------------------------------
// Tests.
// ---------------------------------------------------------------------------

describe("REAL hub + REAL encrypted clients, full loop (repro)", () => {
  it("A) two clients type + press Enter, interleaved, across a checkpoint — converge, no refusal", async () => {
    const hub = new CollabHub(null);
    const docKey = mintDocKey();
    const seed = await seedEncrypted("hello world", "g1", docKey);
    expect(hub.seedEncrypted("d", "g1", seed).ok).toBe(true);
    const bus = new Bus(hub);

    const alice = makeClient(bus, hub, "ca", "alice", docKey);
    const bob = makeClient(bus, hub, "cb", "bob", docKey);
    alice.conn.join("d");
    bob.conn.join("d");
    await settle(bus);
    expect(alice.conn.ready && bob.conn.ready).toBe(true);

    const rand = prng(1234);
    // Interleave one-in-flight edits, ~every 6th an Enter, long enough to
    // cross seq 50 (checkpoint) and keep going past the prune.
    for (let step = 0; step < 70; step++) {
      const who = step % 2 === 0 ? alice.conn : bob.conn;
      const list = addrs(who);
      const a = list[Math.floor(rand() * list.length)];
      const off = a.len > 0 ? Math.floor(rand() * (a.len + 1)) : 0;
      if (step % 6 === 5) pressEnter(who, a, off);
      else typeChar(who, a, off);
      await settle(bus);
    }

    // Server-ordered canonical oracle: a late joiner replays checkpoint+tail
    // with NO optimistic history — its doc is the pure canonical.
    const carol = await joinFresh(bus, docKey, "cc", "carol");

    const noRefusal = [...alice.refusals, ...bob.refusals, ...carol.refusals];
    expect(noRefusal, `refusals seen: ${JSON.stringify(noRefusal)}`).toEqual([]);
    expect(textOf(alice.conn)).toBe(textOf(bob.conn));
    expect(xmlOf(alice.conn), "alice vs bob diverged").toBe(xmlOf(bob.conn));
    expect(xmlOf(alice.conn), "alice diverged from server-ordered canonical (carol)").toBe(xmlOf(carol.conn));
  }, 60000);

  it("B) fast concurrent bursts (multi-pending inserts + Enter) across a checkpoint — converge, no refusal", async () => {
    const hub = new CollabHub(null);
    const docKey = mintDocKey();
    const seed = await seedEncrypted("hello world", "g1", docKey);
    expect(hub.seedEncrypted("d", "g1", seed).ok).toBe(true);
    const bus = new Bus(hub);

    const alice = makeClient(bus, hub, "ca", "alice", docKey);
    const bob = makeClient(bus, hub, "cb", "bob", docKey);
    alice.conn.join("d");
    bob.conn.join("d");
    await settle(bus);
    expect(alice.conn.ready && bob.conn.ready).toBe(true);

    const rand = prng(98765);
    // Simultaneous bursts: within a chunk, BOTH clients fire several ops with
    // NO settle, so each holds multiple un-confirmed pending (incl. an Enter)
    // AT THE SAME TIME. When the echoes/broadcasts then flow, a remote entry
    // lands on a client while its own split is still an un-confirmed deferred
    // own-echo (ClientReplica.confirmedTail) — the real "both people typing
    // fast, someone hits Enter" case. Settle only BETWEEN chunks, and run
    // past seq 50 so a checkpoint prunes mid-stream.
    for (let chunk = 0; chunk < 16; chunk++) {
      // Alice opens with an Enter first (so the split is the earliest pending,
      // most likely to sit in confirmedTail when a remote interleaves), then
      // both clients keep typing into the same unsettled window.
      const aList = addrs(alice.conn);
      const aa = aList[Math.floor(rand() * aList.length)];
      pressEnter(alice.conn, aa, aa.len > 0 ? Math.floor(rand() * (aa.len + 1)) : 0);
      for (let k = 0; k < 3; k++) {
        const al = addrs(alice.conn);
        const a1 = al[Math.floor(rand() * al.length)];
        typeChar(alice.conn, a1, a1.len > 0 ? Math.floor(rand() * (a1.len + 1)) : 0);
        const bl = addrs(bob.conn);
        const b1 = bl[Math.floor(rand() * bl.length)];
        if (k === 1) pressEnter(bob.conn, b1, b1.len > 0 ? Math.floor(rand() * (b1.len + 1)) : 0);
        else typeChar(bob.conn, b1, b1.len > 0 ? Math.floor(rand() * (b1.len + 1)) : 0);
      }
      await settle(bus);
    }
    await settle(bus);

    const carol = await joinFresh(bus, docKey, "cc", "carol");

    const noRefusal = [...alice.refusals, ...bob.refusals, ...carol.refusals];
    expect(noRefusal, `refusals seen: ${JSON.stringify(noRefusal)}`).toEqual([]);
    expect(
      xmlOf(alice.conn),
      `alice vs bob diverged\nA:${textOf(alice.conn)}\nB:${textOf(bob.conn)}`,
    ).toBe(xmlOf(bob.conn));
    expect(xmlOf(alice.conn), "alice diverged from server-ordered canonical (carol)").toBe(xmlOf(carol.conn));
  }, 60000);
});

// ---------------------------------------------------------------------------
// C) Root-cause isolation, synchronous & crypto-free: reproduce the exact
//    ClientReplica reconciliation ordering that diverges, with no hub, no
//    transport, no async. Pins the failure to `restoreConfirmed` replaying a
//    `confirmedTail` splitParagraph via applyIntent WITHOUT the
//    doc.refresh()+assignFromRoots that every OTHER apply site performs after
//    a split (session.ts:172-175, replica.ts advanceConfirmed:249 /
//    replayPending:269 / loadCanonical). The stale `doc.sections` then makes
//    `buildRunMap` (apply.ts) miss the split's new run, so the interleaved
//    remote edit addressing that run silently fails to apply and is dropped.
// ---------------------------------------------------------------------------

describe("root cause: ClientReplica.restoreConfirmed replays a confirmedTail split without refresh", () => {
  it("a burst [Enter, type] + an interleaved remote insert into the split's new run silently diverges from canonical", () => {
    const bytes = docxBytes("hello"); // one paragraph: blockId 1, runId 2, w:t "hello"
    const authority = new DocumentSession(DocxDocument.load(bytes)); // server-side canonical
    const clientUT = new ClientReplica(bytes); // the client whose reconciliation we probe
    const observer = new ClientReplica(bytes); // passive replica = pure canonical oracle

    // Client A's fast burst: press Enter after "he", then type — TWO pending
    // in flight at once (the multi-pending typing burst).
    const split = { kind: "splitParagraph", clientId: "A", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, newBlockId: 900, newRunId: 901 } as never;
    const typeA = { kind: "insertText", clientId: "A", clientSeq: 2, base: 0, at: { blockId: 1, runId: 2, offset: 1 }, text: "Z" } as never;
    clientUT.submitLocal(split);
    clientUT.submitLocal(typeA);

    // The server sequences the split first and echoes it back. clientUT still
    // has `typeA` pending, so this own-echo is DEFERRED into confirmedTail
    // (replica.ts:217-233) — NOT folded via a refreshing path.
    const eSplit = authority.submit(split);
    clientUT.receive([eSplit]);
    observer.receive([eSplit]);

    // Now a REMOTE insert from client B lands, addressed into the ORIGINAL run
    // at an offset that the split moved into the NEW run (901). Canonically it
    // transforms onto run 901 — the server does exactly that.
    const remoteB = { kind: "insertText", clientId: "B", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 4 }, text: "Q" } as never;
    const eRemote = authority.submit(remoteB);
    // This is the trigger: clientUT has `typeA` pending + a remote entry, so it
    // hits the slow path -> restoreConfirmed replays confirmedTail=[split]
    // WITHOUT refresh, leaving doc.sections stale, so applying eRemote (into
    // run 901) finds no such run in buildRunMap and drops it.
    clientUT.receive([eRemote]);
    observer.receive([eRemote]);

    // Finally clientUT's own typeA is sequenced and echoed back.
    const eTypeA = authority.submit(typeA);
    clientUT.receive([eTypeA]);
    observer.receive([eTypeA]);

    const canonical = serializeXml(authority.doc.docRoot);
    const observed = serializeXml(observer.doc.docRoot);
    const underTest = serializeXml(clientUT.doc.docRoot);

    // Sanity: the oracle is sound — a passive replica tracks canonical exactly.
    expect(observed, "passive observer replica must equal canonical").toBe(canonical);
    expect(canonical, "canonical must contain the remote insert").toContain("Q");

    // THE REPRO (fails today): the client that reconciled through the
    // confirmedTail-split path MUST also equal canonical, but does not — the
    // remote "Q" was silently dropped because eRemote addressed the split's
    // new run (901), which was absent from buildRunMap's stale doc.sections.
    expect(underTest, `clientUT diverged; its content dropped the remote insert -> ${underTest}`).toBe(canonical);
  });
});
