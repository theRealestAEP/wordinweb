#!/usr/bin/env node
/**
 * GO-LIVE PHASE BENCHMARK — the "Make collaborative chugs on a 500-page
 * document" report, measured phase by phase.
 *
 * Reproduces the demo's exact go-live pipeline (local-editor startCollab →
 * goLiveEncrypted → sealSeed → PUT) against the REAL zero-custody server,
 * and times each phase separately:
 *
 *   save     api.save()            — serialise the edited doc to docx bytes
 *   stretch  stretchShareCode      — PBKDF2-SHA256, 600k iterations
 *   derive   deriveEpochKeys       — HKDF → 3 AES-GCM keys
 *   parse    DocxDocument.load     — sealSeed re-parses the bytes it was given
 *   hash     docHash               — canonical serialise + SHA-256
 *   media    mediaAddressesOf      — media address map for the checkpoint
 *   seal     sealCheckpoint        — b64(docx) + JSON + pad + AES-GCM + b64
 *   upload   PUT /docs/:id         — loopback HTTP (a LOWER bound on real nets)
 *
 * Run:   node scripts/bench-golive.mjs [--paras=18000] [--port=12399]
 * Needs: npm run build -w @wordinweb/core -w @wordinweb/collab -w @wordinweb/server
 *
 * ~37 paragraphs lay out to a page (e2e/bigdoc-typing.spec.ts: 2200 ≈ 60
 * pages), so the default 18000 ≈ the reported 500-page document.
 *
 * Output: STRESS-METRIC lines in the perf-report.mjs shape.
 */

import { spawn } from "node:child_process";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import {
  deriveEpochKeys,
  stretchShareCode,
  sealCheckpoint,
  mintDocKey,
  bytesToB64,
  docHash,
  mediaAddressesOf,
} from "@wordinweb/collab/client";

const arg = (name, dflt) => {
  const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
  return hit ? Number(hit.split("=")[1]) : dflt;
};
const PARAS = arg("paras", 18000);
const PORT = arg("port", 12399);
/** Embedded media (random bytes, incompressible), in MiB. A text-only 500
 * pager zips to ~1–2 MB; the "tens of megabytes" documents are the ones
 * carrying images, which is what this simulates. */
const MEDIA_MB = arg("mediaMB", 0);

const metric = (scenario, fields) =>
  console.log(
    `STRESS-METRIC ${scenario} ` +
      Object.entries(fields)
        .map(([k, v]) => `${k}=${typeof v === "number" && !Number.isInteger(v) ? v.toFixed(2) : v}`)
        .join(" "),
  );

function docBytes(paras) {
  // Seeded LCG word soup: repetitive prose zips ~40:1 and would understate a
  // real document's bytes by an order of magnitude. Varied words keep the
  // compressed size in the realistic ~1.5 MB range for 500 text pages.
  let seed = 42;
  const rnd = () => ((seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff);
  const word = () => {
    let w = "";
    const len = 3 + Math.floor(rnd() * 8);
    for (let i = 0; i < len; i++) w += String.fromCharCode(97 + Math.floor(rnd() * 26));
    return w;
  };
  const para = (i) => {
    let text = `Paragraph ${i}:`;
    for (let j = 0; j < 18; j++) text += ` ${word()}`;
    return `<w:p><w:r><w:t xml:space="preserve">${text}. </w:t></w:r></w:p>`;
  };
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  const parts = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  };
  if (MEDIA_MB > 0) {
    const blob = new Uint8Array(MEDIA_MB * 1024 * 1024);
    for (let i = 0; i < blob.length; i += 65536) {
      crypto.getRandomValues(blob.subarray(i, Math.min(i + 65536, blob.length)));
    }
    parts["word/media/image1.png"] = blob;
  }
  return zipSync(parts);
}

function randHex(bytes) {
  const b = new Uint8Array(bytes);
  crypto.getRandomValues(b);
  return [...b].map((x) => x.toString(16).padStart(2, "0")).join("");
}

/* --------------------------- real server -------------------------------- */
const server = spawn("node", ["packages/server/dist/cli.js"], {
  env: { ...process.env, PORT: String(PORT) },
  stdio: ["ignore", "ignore", "ignore"],
});
const httpBase = `http://localhost:${PORT}`;
for (let i = 0; ; i++) {
  try {
    const r = await fetch(`${httpBase}/healthz`);
    if (r.ok) break;
  } catch {
    /* not up yet */
  }
  if (i > 100) throw new Error("server never became healthy");
  await new Promise((r) => setTimeout(r, 100));
}

/* ------------------------------ phases ---------------------------------- */
const now = () => performance.now();
const t = {};

// The document the user has been editing locally (parse cost is NOT part of
// go-live — the doc is already open on screen when the button is clicked).
const editedDoc = DocxDocument.load(docBytes(PARAS));

let t0 = now();
const docx = editedDoc.save(); // ← startCollab: bytes = api.save()
t.saveMs = now() - t0;

const docId = `d_${randHex(16)}`;
const docKey = mintDocKey();
const genesisId = `g_${randHex(16)}`;

t0 = now();
const stretched = await stretchShareCode("redwood", docId);
t.stretchMs = now() - t0;

t0 = now();
const keys = await deriveEpochKeys(docKey, genesisId, stretched);
t.deriveMs = now() - t0;

t0 = now();
const doc = DocxDocument.load(docx); // ← sealSeed's re-parse of the bytes
t.parseMs = now() - t0;

t0 = now();
const hash = await docHash(doc);
t.hashMs = now() - t0;

t0 = now();
const mediaMeta = mediaAddressesOf(doc);
t.mediaMs = now() - t0;

t0 = now();
const sealed = await sealCheckpoint(keys.kContent, docId, genesisId, 0, {
  docx: bytesToB64(docx),
  sidecar: null,
  docHash: hash,
  mediaMeta,
});
t.sealMs = now() - t0;

t0 = now();
const codeVerifier = btoa(String.fromCharCode(...stretched));
let uploadStatus;
try {
  const res = await fetch(`${httpBase}/docs/${docId}`, {
    method: "PUT",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ encrypted: { genesisId, checkpoint: { seq: 0, ...sealed } }, codeVerifier }),
  });
  uploadStatus = res.status;
} catch {
  // The server hangs up mid-body on an over-cap request (EPIPE) — that IS
  // the finding for oversized documents: they cannot go live at all.
  uploadStatus = "refused-connection";
}
t.uploadMs = now() - t0;

const total = Object.values(t).reduce((a, b) => a + b, 0);
metric("golive-phases", {
  paragraphs: PARAS,
  pagesApprox: Math.round(PARAS / 37),
  mediaMB: MEDIA_MB,
  docxBytes: docx.byteLength,
  sealedBytes: sealed.ciphertext.length,
  uploadStatus,
  ...t,
  totalMs: total,
});

server.kill();
process.exit(0);
