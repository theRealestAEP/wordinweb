#!/usr/bin/env node
/**
 * BIG-DOC COLLAB BENCHMARK — engine layer (no DOM, no layout).
 *
 * Reproduces the "500-page document is unusably slow in the collab editor"
 * report at the collab-engine level: a synthetic document of N paragraphs, a
 * run of K keystrokes through the REAL encrypted client stack
 * (EncryptedCollabConnection + blind sequencer + BundlePersister), timed
 * against the same K keystrokes applied through the local canonical apply
 * (the non-collab editing model cost).
 *
 * Run:   node scripts/bench-collab-bigdoc.mjs [--paras=3000] [--keys=200]
 * Needs: npm run build -w @wordinweb/core -w @wordinweb/collab   (uses dist)
 *
 * Output: STRESS-METRIC lines (perf-report.mjs can ingest them) plus a
 * breakdown of where the collab milliseconds go (doc.save / load / refresh /
 * docHash / checkpoint seals), measured by patching the shared core classes.
 */

import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import {
  EncryptedCollabConnection,
  BundlePersister,
  InMemoryBundleStore,
  ClientReplica,
  applyIntentScoped,
  resyncScope,
  deriveEpochKeys,
  sealCheckpoint,
  mintDocKey,
  bytesToB64,
  docHash,
} from "@wordinweb/collab/client";
import { DocumentSession } from "@wordinweb/collab/server";

const arg = (name, dflt) => {
  const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
  return hit ? Number(hit.split("=")[1]) : dflt;
};
const PARAS = arg("paras", 3000);
const KEYS = arg("keys", 200);

const metric = (scenario, fields) =>
  console.log(
    `STRESS-METRIC ${scenario} ` +
      Object.entries(fields)
        .map(([k, v]) => `${k}=${typeof v === "number" && !Number.isInteger(v) ? v.toFixed(2) : v}`)
        .join(" "),
  );

function docBytes(paras) {
  const para = (i) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at considerable length about it. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
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

/** Address of the LAST paragraph's run (the user report: editing at the tail
 * of a 500-page doc), via the given doc's stable-id table. */
function tailAddr(doc, ids) {
  const blocks = doc.sections[0].blocks;
  const para = blocks[blocks.length - 1];
  const run = para.children.find((c) => c.src && c.content?.some?.((x) => x.kind === "text"));
  const t = run.content.find((c) => c.kind === "text").srcT;
  return { blockId: ids.idOf(para.src), runId: ids.idOf(run.src), len: t.text.length };
}

/* ------------------------- instrumentation ------------------------------ */
const tallies = new Map(); // name -> {n, ms}
function tally(name, ms) {
  const t = tallies.get(name) ?? { n: 0, ms: 0 };
  t.n++;
  t.ms += ms;
  tallies.set(name, t);
}
function patch(obj, key, name) {
  const orig = obj[key];
  obj[key] = function (...args) {
    const t0 = performance.now();
    const out = orig.apply(this, args);
    if (out && typeof out.then === "function") {
      return out.then(
        (value) => {
          tally(name, performance.now() - t0);
          return value;
        },
        (error) => {
          tally(name, performance.now() - t0);
          throw error;
        },
      );
    }
    tally(name, performance.now() - t0);
    return out;
  };
  return () => (obj[key] = orig);
}
function resetTallies() {
  tallies.clear();
}
function snapshotTallies() {
  return Object.fromEntries([...tallies].map(([k, v]) => [k, { n: v.n, ms: +v.ms.toFixed(1) }]));
}

/* ------------------------- blind sequencer ------------------------------ */
/** Mirrors hub.ts submit-enc semantics (same shape as the e2ee test hub). */
function blindServer(genesisId, checkpoint) {
  let log = [];
  const seen = new Set();
  const peers = [];
  const state = { checkpoint, gossips: 0, checkpoints: 0 };
  let designated = false;
  const attach = () => {
    const peer = { deliver: () => {} };
    peers.push(peer);
    const first = !designated;
    designated = true;
    return {
      send: (msg) => {
        if (msg.t === "hello") {
          peer.deliver({
            t: "welcome-enc",
            docId: "d",
            genesisId,
            checkpoint: state.checkpoint,
            tail: log.filter((e) => e.seq > state.checkpoint.seq),
            mode: "encrypted",
          });
          if (first) peer.deliver({ t: "checkpointer", active: true });
        } else if (msg.t === "checkpoint") {
          state.checkpoints++;
          const head = log.length === 0 ? state.checkpoint.seq : log[log.length - 1].seq;
          if (msg.checkpoint.seq === head && msg.checkpoint.seq > state.checkpoint.seq) {
            state.checkpoint = msg.checkpoint;
            log = [];
          }
        } else if (msg.t === "submit-enc") {
          if (msg.envelope.base < state.checkpoint.seq) {
            peer.deliver({ t: "refused", reason: "stale-base" });
            return;
          }
          const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
          let entry = seen.has(key) ? log.find((e) => `${e.clientId}:${e.clientSeq}` === key) : undefined;
          if (!entry) {
            seen.add(key);
            entry = { ...msg.envelope, seq: log.length === 0 ? state.checkpoint.seq + 1 : log[log.length - 1].seq + 1 };
            log.push(entry);
          }
          for (const p of peers) p.deliver({ t: "broadcast-enc", entries: [entry] });
        } else if (msg.t === "gossip") {
          state.gossips++;
          for (const p of peers) p.deliver({ t: "gossip", iv: msg.iv, ciphertext: msg.ciphertext });
        }
      },
      onMessage: (cb) => {
        peer.deliver = cb;
      },
    };
  };
  return { attach, state };
}

async function until(cond, label, ms = 60000) {
  const t0 = Date.now();
  for (;;) {
    if (cond()) return;
    if (Date.now() - t0 > ms) throw new Error(`timeout: ${label}`);
    await new Promise((r) => setTimeout(r, 1));
  }
}

/* ------------------------------ phases ---------------------------------- */

async function primitives(bytes) {
  let t0 = performance.now();
  const doc = DocxDocument.load(bytes);
  const loadMs = performance.now() - t0;

  const ids = doc.enableStableIds();
  t0 = performance.now();
  const saved = doc.save();
  const saveMs = performance.now() - t0;

  t0 = performance.now();
  const sidecar = ids.exportSidecar(doc.editableRoots());
  const sidecarMs = performance.now() - t0;

  t0 = performance.now();
  await docHash(doc);
  const hashMs = performance.now() - t0;

  t0 = performance.now();
  doc.refresh();
  const refreshMs = performance.now() - t0;

  // Mirror-ingest cost of one keystroke (a standalone DocumentSession is the
  // same class the connection's mirror runs).
  const session = new DocumentSession(DocxDocument.load(bytes));
  const addr = tailAddr(session.doc, session.ids);
  t0 = performance.now();
  session.submit({
    kind: "insertText",
    clientId: "m",
    clientSeq: 1,
    base: 0,
    at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len },
    text: "x",
  });
  const mirrorSubmitMs = performance.now() - t0;

  // Checkpoint build+seal (what the designated checkpointer pays every
  // CHECKPOINT_EVERY seqs).
  const keys = await deriveEpochKeys(mintDocKey(), "g_bench");
  t0 = performance.now();
  const cp = session.checkpoint();
  await sealCheckpoint(keys.kContent, "d", "g_bench", cp.seq, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: await docHash(session.doc),
  });
  const checkpointMs = performance.now() - t0;

  // Persister export cost, quiescent (tail folded from live doc: save +
  // sidecar) and mid-burst (load baseline + replay + save).
  const replica = new ClientReplica(bytes, sidecar);
  replica.receive([
    {
      seq: 1,
      kind: "applied",
      intent: {
        kind: "insertText",
        clientId: "p",
        clientSeq: 1,
        base: 0,
        at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len },
        text: "y",
      },
    },
  ]);
  t0 = performance.now();
  replica.exportBundleState();
  const exportQuiescentMs = performance.now() - t0;

  const replica2 = new ClientReplica(bytes, sidecar);
  replica2.receive([
    {
      seq: 1,
      kind: "applied",
      intent: {
        kind: "insertText",
        clientId: "p",
        clientSeq: 1,
        base: 0,
        at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len },
        text: "y",
      },
    },
  ]);
  replica2.submitLocal({
    kind: "insertText",
    clientId: "p2",
    clientSeq: 1,
    base: 1,
    at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len + 1 },
    text: "z",
  });
  t0 = performance.now();
  replica2.exportBundleState();
  const exportMidBurstMs = performance.now() - t0;

  metric("bigdoc-primitives", {
    paragraphs: PARAS,
    bytes: bytes.length,
    savedBytes: saved.length,
    loadMs,
    saveMs,
    sidecarMs,
    docHashMs: hashMs,
    refreshMs,
    mirrorSubmitMs,
    checkpointSealMs: checkpointMs,
    exportBundleQuiescentMs: exportQuiescentMs,
    exportBundleMidBurstMs: exportMidBurstMs,
  });
}

/** K keystrokes through the canonical apply on a plain doc — the engine-side
 * model cost of local (non-collab) editing. */
function localBaseline(bytes) {
  const doc = DocxDocument.load(bytes);
  const ids = doc.enableStableIds();
  const addr = tailAddr(doc, ids);
  const t0 = performance.now();
  for (let i = 0; i < KEYS; i++) {
    const res = applyIntentScoped(doc, ids, {
      kind: "insertText",
      clientId: "l",
      clientSeq: i + 1,
      base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len + i },
      text: "x",
    });
    if (!res.applied) throw new Error(`local apply ${i} failed`);
    resyncScope(doc, ids, res);
  }
  const total = performance.now() - t0;
  metric("bigdoc-local-typing", { paragraphs: PARAS, keystrokes: KEYS, totalMs: total, msPerKey: total / KEYS });
}

/** K keystrokes through the real encrypted stack, one-in-flight, with the
 * bundle persister notified per change (as the react hook wires it). */
async function collabTyping(bytes, { persist }) {
  const docKey = mintDocKey();
  const genesisId = "g_bench";
  const keys = await deriveEpochKeys(docKey, genesisId);
  const seedSession = new DocumentSession(DocxDocument.load(bytes));
  const cp = seedSession.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: "seed",
  });
  const srv = blindServer(genesisId, { seq: 0, ...sealed });

  const store = new InMemoryBundleStore();
  let persister = null;
  const conn = new EncryptedCollabConnection(
    srv.attach(),
    "typist",
    docKey,
    { onChange: () => persister?.notify() },
    undefined,
  );
  if (persist) persister = new BundlePersister(conn, store, "d", {});
  conn.join("d");
  await until(() => conn.ready, "welcome");

  const addr = tailAddr(conn.doc, conn.doc.stableIds);
  resetTallies();
  // The collab dist bundles its own copy of @wordinweb/core (tsup noExternal),
  // so patch the BUNDLED classes — reachable through live instances and the
  // package's own exports (client.js/server.js share one chunk).
  const BundledDoc = conn.doc.constructor;
  const BundledIds = conn.doc.stableIds.constructor;
  const unpatch = [
    patch(BundledDoc.prototype, "save", "doc.save"),
    patch(BundledDoc.prototype, "saveAsync", "doc.saveAsync"),
    patch(BundledDoc.prototype, "refresh", "doc.refresh"),
    patch(BundledDoc, "load", "DocxDocument.load"),
    patch(BundledIds.prototype, "exportSidecar", "ids.exportSidecar"),
    patch(BundledIds.prototype, "assignFromRoots", "ids.assignFromRoots"),
    patch(DocumentSession.prototype, "submit", "mirror.submit"),
    patch(DocumentSession.prototype, "checkpoint", "mirror.checkpoint"),
    patch(ClientReplica.prototype, "receive", "replica.receive"),
    patch(ClientReplica.prototype, "exportBundleState", "replica.exportBundleState"),
    patch(ClientReplica.prototype, "exportBundleStateAsync", "replica.exportBundleStateAsync"),
  ];

  const perKey = [];
  const perKeyAllocMb = [];
  const heapStartMb = process.memoryUsage().heapUsed / 1e6;
  const t0 = performance.now();
  for (let i = 0; i < KEYS; i++) {
    const h0 = process.memoryUsage().heapUsed;
    const k0 = performance.now();
    conn.submit({
      kind: "insertText",
      at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len + i },
      text: "x",
    });
    await until(() => conn.pendingCount === 0, `echo ${i}`);
    perKey.push(performance.now() - k0);
    perKeyAllocMb.push(Math.max(0, (process.memoryUsage().heapUsed - h0) / 1e6));
  }
  const total = performance.now() - t0;
  // Let the trailing persister write + self-check land inside the patched
  // window so their cost is attributed.
  await new Promise((r) => setTimeout(r, 1500));
  const settleMs = performance.now() - t0 - total;
  unpatch.forEach((u) => u());

  const sorted = [...perKey].sort((a, b) => a - b);
  const pct = (p) => sorted[Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length))];
  const scenario = persist ? "bigdoc-collab-typing" : "bigdoc-collab-typing-nopersist";
  const heapEndMb = process.memoryUsage().heapUsed / 1e6;
  const positiveAllocMb = perKeyAllocMb.reduce((sum, value) => sum + value, 0);
  const allocSorted = [...perKeyAllocMb].sort((a, b) => a - b);
  metric(scenario, {
    paragraphs: PARAS,
    keystrokes: KEYS,
    totalMs: total,
    msPerKey: total / KEYS,
    p50: pct(50),
    p90: pct(90),
    p99: pct(99),
    max: sorted[sorted.length - 1],
    settleMs,
    persistWrites: store.writes,
    gossips: srv.state.gossips,
    checkpoints: srv.state.checkpoints,
    sendFailures: conn.sendFailures,
    selfHeals: conn.selfHeals,
    heapStartMB: heapStartMb,
    heapEndMB: heapEndMb,
    heapGrowthMB: heapEndMb - heapStartMb,
    positiveAllocMB: positiveAllocMb,
    allocMBPerKeyMedian: allocSorted[Math.floor(allocSorted.length / 2)],
    allocMBPerKeyMax: Math.max(...perKeyAllocMb),
    allocMBps: positiveAllocMb / (total / 1000),
  });
  const t = snapshotTallies();
  for (const [name, v] of Object.entries(t)) {
    metric(`${scenario}-where`, { op: name.replace(/\s+/g, "_"), calls: v.n, totalMs: v.ms, msPerKey: v.ms / KEYS });
  }
}

const bytes = docBytes(PARAS);
await primitives(bytes);
localBaseline(bytes);
await collabTyping(bytes, { persist: false });
await collabTyping(bytes, { persist: true });
