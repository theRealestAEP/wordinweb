#!/usr/bin/env node
/**
 * Heap attribution for the line-break cache on a big document.
 *
 * bench-heap-bigdoc.mjs reports the total; this says how much of that total is
 * the measurer's break cache, by laying out a ~500-page document with the
 * windowed model and then dropping the cache and re-measuring. Measured at
 * 3,500 paragraphs: 180 MB of a 192 MB layout heap. The page window already
 * shrank the item model to ~12 MB, so the break cache is now nearly all of it.
 *
 * Capping the cache smaller is not free — bench-split-cost.mjs measures the
 * relay cost that pays for those entries.
 *
 * Imports TypeScript source directly, so run it through tsx:
 *   npx tsx --expose-gc scripts/bench-cache-heap.mjs [--paras=3500] [--chars=600]
 */
import { zipSync, strToU8 } from "fflate";

if (typeof globalThis.gc !== "function") {
  console.error("run with node --expose-gc");
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
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join("")}</w:body></w:document>`;
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

const core = await import("../packages/core/src/index.ts");
const { DocxDocument, layoutDocument, ApproxMeasurer, clearBreakCache } = core;

const bytes = docBytes(PARAS, CHARS);
const base = heapMB();
const doc = DocxDocument.load(bytes);
const parsed = heapMB();
const measurer = new ApproxMeasurer();
let result = layoutDocument(doc, { measurer, windowModel: true });
const laid = heapMB();

// The break cache is a module-level WeakMap keyed by measurer; dropping its
// entries for this measurer isolates exactly what it retains.
clearBreakCache(measurer);
const afterBreakClear = heapMB();

const pages = result.pages.length;
console.log(
  `CACHE-HEAP paras=${PARAS} pages=${pages} ` +
    `parsedMB=${fmt(parsed - base)} layoutMB=${fmt(laid - parsed)} ` +
    `breakCacheMB=${fmt(laid - afterBreakClear)} ` +
    `layoutWithoutBreakCacheMB=${fmt(afterBreakClear - parsed)}`,
);
void result;
