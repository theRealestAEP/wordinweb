#!/usr/bin/env node
/**
 * BIG-DOCUMENT HEAP BENCHMARK — jsdom, real DocxView, staged attribution.
 *
 * Measures what a large document COSTS IN MEMORY, which nothing else tracks:
 * a ~500-page text document holds ~150x its own size in JS heap today, all of
 * it in the layout page model (one TextItem per word and per space, ~550 B
 * each). This instrument reports the model cost per page so a regression — or
 * an improvement — shows up as a number in internal/perf.
 *
 * Retained stages use full GC. Peak stages sample reachable heap at each
 * window boundary before the next page range is released:
 *   parsedMB   — DocxDocument.load (XML tree + package)
 *   layoutMB   — layoutDocument page model on top of the parsed doc
 *   layoutPeakMB — highest sampled layout growth above the parsed document
 *   mountedMB  — full DocxView mount (renderer DOM + editor + React) in jsdom
 *   openPeakHeapMB — highest sampled total heap during the full app mount
 *   undoMB     — growth after `undoKeys` separate (non-coalesced) checkpoints
 *
 * Run:
 *   node --expose-gc scripts/bench-heap-bigdoc.mjs --paras=3500 --chars=600
 *
 * Same WW_PKG contract as bench-local-typing.mjs: point it at any built
 * wordinweb package (this tree's packages/react/dist or a published version)
 * to compare builds.
 *
 * jsdom caveat: mountedMB includes jsdom's JS-object DOM for the mounted page
 * window, which Chrome keeps in native memory instead. parsedMB and layoutMB
 * are engine-only and transfer to the browser as-is; they are the numbers the
 * 150x problem lives in.
 */

import { JSDOM } from "jsdom";
import { zipSync, strToU8 } from "fflate";

const arg = (name, dflt) => {
  const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
  return hit ? Number(hit.split("=")[1]) : dflt;
};
const PARAS = arg("paras", 3500);
const CHARS = arg("chars", 600);
const UNDO_KEYS = arg("undoKeys", 5);

if (typeof globalThis.gc !== "function") {
  console.error("bench-heap-bigdoc: run with node --expose-gc (heap numbers need forced GC)");
  process.exit(1);
}
const heapMB = () => {
  globalThis.gc();
  globalThis.gc();
  return process.memoryUsage().heapUsed / 1048576;
};
const fmt = (n) => n.toFixed(1);
let peakBaseBytes = 0;
let peakBytes = 0;
const sampleHeap = () => {
  globalThis.gc();
  globalThis.gc();
  peakBytes = Math.max(peakBytes, process.memoryUsage().heapUsed);
};
const resetPeak = () => {
  peakBaseBytes = process.memoryUsage().heapUsed;
  peakBytes = peakBaseBytes;
};
const peakGrowthMB = () => (peakBytes - peakBaseBytes) / 1048576;
globalThis.__dxwPerf = { heapSample: sampleHeap };

/* ------------------------- jsdom environment ---------------------------- */
const dom = new JSDOM("<!doctype html><html><body></body></html>", {
  pretendToBeVisual: true,
  url: "http://localhost/",
});
const { window } = dom;
for (const k of [
  "document", "window", "navigator", "HTMLElement", "HTMLCanvasElement", "Element", "Node",
  "MouseEvent", "KeyboardEvent", "CustomEvent", "Event", "Range", "getComputedStyle",
  "requestAnimationFrame", "cancelAnimationFrame", "ResizeObserver", "MutationObserver",
  "IntersectionObserver", "CSS", "DOMParser", "XMLSerializer", "location", "history",
  "customElements", "HTMLInputElement", "HTMLTextAreaElement", "SVGElement", "Text",
]) {
  if (window[k] === undefined) continue;
  try {
    globalThis[k] = window[k];
  } catch {
    Object.defineProperty(globalThis, k, { value: window[k], configurable: true });
  }
}
globalThis.IS_REACT_ACT_ENVIRONMENT = true;

function fakeCtx() {
  return {
    font: "",
    measureText(text) {
      const size = Number(/([\d.]+)px/.exec(this.font)?.[1]) || 12;
      return { width: size * 0.5 * text.length };
    },
    fillText() {}, clearRect() {}, setTransform() {}, scale() {}, translate() {},
    save() {}, restore() {}, beginPath() {}, fill() {}, drawImage() {},
  };
}
window.HTMLCanvasElement.prototype.getContext = function () { return fakeCtx(); };
globalThis.OffscreenCanvas = class { constructor(w = 0, h = 0) { this.width = w; this.height = h; } getContext() { return fakeCtx(); } };
if (!window.document.elementFromPoint) window.document.elementFromPoint = () => null;
if (!window.Range.prototype.getBoundingClientRect) {
  window.Range.prototype.getBoundingClientRect = () => ({ left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0, x: 0, y: 0, toJSON() {} });
}
if (!window.Element.prototype.scrollIntoView) window.Element.prototype.scrollIntoView = () => {};
if (!window.ResizeObserver) {
  globalThis.ResizeObserver = window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
}

/* ------------------------------ doc bytes ------------------------------- */
function docBytes(paras, chars) {
  const sentence =
    "The quick brown fox jumps over the lazy dog while the committee deliberates at length. ";
  const body = [];
  for (let i = 0; i < paras; i++) {
    let text = `Paragraph ${i}: `;
    while (text.length < chars) text += sentence;
    // Three runs per paragraph (plain / bold / italic), like a real document.
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

/* -------------------------------- run ----------------------------------- */
const pkg = process.env.WW_PKG ?? "wordinweb";
const corePkg = process.env.WW_CORE ?? pkg;
const [{ createElement }, { act }, { createRoot }, mod, core] = await Promise.all([
  import("react"),
  import("react"),
  import("react-dom/client"),
  import(pkg),
  import(corePkg),
]);
const DocxView = mod.DocxView;
if (!DocxView) throw new Error(`no DocxView export in ${pkg}`);
if (!core.DocxDocument || !core.layoutDocument) throw new Error(`no DocxDocument/layoutDocument export in ${corePkg}`);

/** Version-independent deterministic measurer (mirrors core's ApproxMeasurer)
 * so the engine-only stages compare identically across builds that may not
 * export one. */
function approxWidth(text, font, letterSpacing = 0) {
  let w = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 32;
    let em;
    if (ch === " ") em = 0.25;
    else if ("iIljtf.,;:!|'".includes(ch)) em = 0.28;
    else if ("mwMW".includes(ch)) em = 0.85;
    else if (ch >= "A" && ch <= "Z") em = 0.66;
    else if (ch >= "0" && ch <= "9") em = 0.5;
    else if (code > 0x2e80) em = 1.0;
    else em = 0.5;
    if (font.bold) em *= 1.05;
    w += em * font.size;
  }
  return w + letterSpacing * text.length;
}
const benchMeasurer = {
  width: approxWidth,
  metrics: (font) => ({ ascent: font.size * 0.9, descent: font.size * 0.25, lineHeight: font.size * 1.15 }),
};

const h0 = heapMB();
const bytes = docBytes(PARAS, CHARS);
const docxMB = bytes.length / 1048576;

/* Engine-only stages first: these transfer to the browser unchanged. */
const doc = await core.DocxDocument.load(bytes.slice());
const h1 = heapMB();
resetPeak();
let layout = core.layoutDocument(doc, { measurer: benchMeasurer, windowModel: true });
sampleHeap();
const layoutPeakMB = peakGrowthMB();
const h2 = heapMB();
const pages = layout.totalPages;
let items = 0;
for (const p of layout.pages) items += p.items.length;
const retainedPages = layout._window?.retainedPages().size ?? pages;
const fontSamples = layout._fontSamples?.length ?? 0;
const relayPageCount = layout._window ? 12 : 0;
const relayStart = Math.floor((pages - relayPageCount) / 2);
const relayPages = Array.from({ length: relayPageCount }, (_, index) => relayStart + index);
const retainedBeforeRelay = layout._window?.retainedPages();
globalThis.__dxwPerf.heapSample = undefined;
const relayStarted = performance.now();
layout._window?.materialize(relayPages);
const relayMs = performance.now() - relayStarted;
if (retainedBeforeRelay) layout._window?.releaseExcept(retainedBeforeRelay);
globalThis.__dxwPerf.heapSample = sampleHeap;
layout = null;

/* Full app mount on its own copy of the bytes. */
const h3 = heapMB();
resetPeak();
Object.defineProperty(window.HTMLElement.prototype, "clientHeight", { get() { return 1200; }, configurable: true });
const container = window.document.createElement("div");
window.document.body.appendChild(container);
const root = createRoot(container);
const tick = async (ms = 5) => { await act(async () => { await new Promise((r) => setTimeout(r, ms)); }); };
await act(async () => {
  root.render(
    createElement(DocxView, {
      source: docBytes(PARAS, CHARS),
      editable: true,
      style: { height: "500px" },
    }),
  );
});
for (let i = 0; i < 400 && !container.querySelector(".dxw-page"); i++) await tick();
if (!container.querySelector(".dxw-page")) throw new Error("document never painted");
const busy = () => !!container.querySelector("[data-dxw-layout-busy]");
for (let i = 0; i < 6000 && busy(); i++) await tick(5);
sampleHeap();
const mountedPeakMB = peakGrowthMB();
const openPeakHeapMB = peakBytes / 1048576;
const h4 = heapMB();

/* Undo growth: separate checkpoints (typing bursts >1s apart each snapshot
 * the editable roots — on a big document that is the second-order leak). */
const span = container.querySelector(".dxw-page span") ?? container.querySelector(".dxw-page");
await act(async () => {
  const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
  span.dispatchEvent(new window.MouseEvent("mousedown", o));
  span.dispatchEvent(new window.MouseEvent("mouseup", o));
});
await tick(10);
const target = () =>
  (container.contains(window.document.activeElement) ? window.document.activeElement : container.querySelector("textarea")) ?? container;
const realNow = Date.now;
let skew = 0;
Date.now = () => realNow() + skew;
const h5 = heapMB();
for (let i = 0; i < UNDO_KEYS; i++) {
  skew += 2000; // defeat the 1s coalesce window: every keystroke = one snapshot
  await act(async () => {
    target().dispatchEvent(new window.KeyboardEvent("keydown", { key: "Z", bubbles: true, cancelable: true }));
    await new Promise((r) => setTimeout(r, 0));
  });
  for (let t = 0; t < 2000 && busy(); t++) await tick(5);
}
Date.now = realNow;
const h6 = heapMB();

console.log(
  `STRESS-METRIC heap-bigdoc pkg=${pkg === "wordinweb" ? "worktree" : "custom"} paragraphs=${PARAS} pages=${pages} ` +
    `docxMB=${fmt(docxMB)} items=${items} itemsPerPage=${Math.round(items / pages)} ` +
    `retainedPages=${retainedPages} fontSamples=${fontSamples} ` +
    `parsedMB=${fmt(h1 - h0)} layoutMB=${fmt(h2 - h1)} layoutPeakMB=${fmt(layoutPeakMB)} ` +
    `layoutKBPerPage=${fmt(((h2 - h1) * 1024) / pages)} mountedMB=${fmt(h4 - h3)} ` +
    `mountedPeakMB=${fmt(mountedPeakMB)} openPeakHeapMB=${fmt(openPeakHeapMB)} ` +
    `relayPages=${relayPageCount} relayMs=${fmt(relayMs)} ` +
    `relayMsPerPage=${fmt(relayPageCount > 0 ? relayMs / relayPageCount : 0)} ` +
    `undoKeys=${UNDO_KEYS} undoMB=${fmt(h6 - h5)} ` +
    `heapTotalMB=${fmt(heapMB())}`,
);
await act(async () => { root.unmount(); });
process.exit(0);
