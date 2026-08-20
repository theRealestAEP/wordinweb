#!/usr/bin/env node
/**
 * What one Enter costs on a big document, in measure calls, against what the
 * break cache retains for it. `--cap` sets the break-cache size so the
 * heap/latency tradeoff can be swept; `--split` moves the edit, which sets how
 * much of the document the relay re-lays behind it.
 *
 * The tradeoff is a cliff, not a curve. Measured at 3,500 paragraphs, splitting
 * at 1,750 (the relay re-lays 1,757 blocks):
 *   cap 60,000 (default)      361 calls   180 MB
 *   cap 2,000                 361 calls   104 MB
 *   cap 1,000             382,897 calls    52 MB
 * The relay walks its blocks once, front to back, with no reuse inside that
 * sweep, so any cap below the number of blocks it re-lays evicts each entry
 * just before it is asked for. Splitting at 100 instead (3,405 blocks re-laid)
 * puts cap 2,000 on the wrong side of the same cliff: 737,044 calls, against
 * 754,090 for a full layout of the whole document.
 *
 * Imports TypeScript source directly, so run it through tsx:
 *   npx tsx --expose-gc scripts/bench-split-cost.mjs [--paras=] [--cap=] [--split=]
 */
import { zipSync, strToU8 } from "fflate";

if (typeof globalThis.gc !== "function") {
  console.error("run with --expose-gc");
  process.exit(1);
}
const heapMB = () => {
  globalThis.gc();
  globalThis.gc();
  return process.memoryUsage().heapUsed / 1048576;
};
const fmt = (n) => n.toFixed(1);
const arg = (name, dflt) => {
  const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
  return hit ? Number(hit.split("=")[1]) : dflt;
};
const PARAS = arg("paras", 3500);
const CHARS = arg("chars", 600);
const CAP = arg("cap", 0); // 0 = leave default
const SPLIT = arg("split", -1); // block index to split at; -1 = midpoint

function docBytes(paras, chars) {
  const sentence =
    "The quick brown fox jumps over the lazy dog while the committee deliberates at length. ";
  const body = [];
  for (let i = 0; i < paras; i++) {
    let text = `Paragraph ${i}: `;
    while (text.length < chars) text += sentence;
    const third = Math.floor(text.length / 3);
    body.push(
      `<w:p><w:r><w:t xml:space="preserve">${text.slice(0, third)}</w:t></w:r>` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text.slice(third, 2 * third)}</w:t></w:r>` +
        `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${text.slice(2 * third)}</w:t></w:r></w:p>`,
    );
  }
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
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join("")}</w:body></w:document>`,
    ),
  });
}

const core = await import("../packages/core/src/index.ts");
const engine = await import("../packages/core/src/layout/engine.ts");
const inline = await import("../packages/core/src/layout/inline.ts");
const { DocxDocument, layoutDocument, ApproxMeasurer, clearBreakCache } = core;
if (CAP > 0 && inline.__setBreakCacheMax) inline.__setBreakCacheMax(CAP);

/** ApproxMeasurer that counts width() calls — the causal unit of layout work. */
let measureCalls = 0;
class CountingMeasurer extends ApproxMeasurer {
  width(text, font, letterSpacing = 0) {
    measureCalls++;
    return super.width(text, font, letterSpacing);
  }
}

function splitAt(doc, idx) {
  const body = doc.docRoot.children.find((c) => c.name.endsWith("body"));
  const paras = body.children.filter((c) => c.name.endsWith(":p"));
  const before = paras[idx];
  const beforeIndex = body.children.indexOf(before);
  const clone = JSON.parse(JSON.stringify(before));
  const findT = (el) => {
    if (el.name.endsWith(":t")) return el;
    for (const c of el.children) {
      const hit = findT(c);
      if (hit) return hit;
    }
    return null;
  };
  const t0 = findT(before);
  const t1 = findT(clone);
  const full = t0.text;
  t0.text = full.slice(0, 20);
  t1.text = full.slice(20);
  body.children.splice(beforeIndex + 1, 0, clone);
  doc.reparseDirectBodyParagraphSplit(before, clone);
  return clone;
}

const bytes = docBytes(PARAS, CHARS);
const base = heapMB();
const doc = DocxDocument.load(bytes);
const parsed = heapMB();
const measurer = new CountingMeasurer();
const first = layoutDocument(doc, { measurer, windowModel: true });
const laid = heapMB();
const mountCalls = measureCalls;

// One Enter in the middle of the document.
const splitIdx = SPLIT >= 0 ? SPLIT : Math.floor(PARAS / 2);
const clone = splitAt(doc, splitIdx);
measureCalls = 0;
layoutDocument(doc, { measurer, windowModel: true, prev: first, dirtyHint: clone });
const splitCalls = measureCalls;
const stats = { ...engine.__incrStats };

const afterLaid = heapMB();
clearBreakCache(measurer);
const afterClear = heapMB();

console.log(
  `SPLIT-COST paras=${PARAS} cap=${CAP || "default"} split=${splitIdx} pages=${first.pages.length} ` +
    `mountCalls=${mountCalls} splitCalls=${splitCalls} ` +
    `blocksLaid=${stats.blocksLaid} converged=${stats.convergedBlock} fb='${stats.fallbackReason}' ` +
    `parsedMB=${fmt(parsed - base)} layoutMB=${fmt(laid - parsed)} breakCacheMB=${fmt(afterLaid - afterClear)}`,
);
