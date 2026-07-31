#!/usr/bin/env node
/**
 * LOCAL (non-collab) TYPING BENCHMARK — jsdom, real DocxView + DocxEditor.
 *
 * Measures per-keystroke latency of the plain editable DocxView on a large
 * document. Runs against ANY build of the `wordinweb` react package, so the
 * same instrument can compare this tree with the published 0.1.22:
 *
 *   # this tree (build first: npm run build -w @wordinweb/core -w wordinweb)
 *   node scripts/bench-local-typing.mjs --paras=3000 --keys=30
 *
 *   # published 0.1.22 (see internal/perf notes): create a scratch dir with
 *   # wordinweb@0.1.22 + react 18 installed, then
 *   WW_PKG=/abs/path/to/scratch/node_modules/wordinweb/dist/index.js \
 *     node /abs/path/to/scripts/bench-local-typing.mjs --paras=3000 --keys=30
 *
 * jsdom has no real layout, so absolute ms understate a browser — but the
 * SHAPE (is a keystroke O(document)?) and the per-phase breakdown
 * (__dxwPerf.samples: refresh / layout / render) transfer.
 */

import { JSDOM } from "jsdom";
import { zipSync, strToU8 } from "fflate";

const arg = (name, dflt) => {
  const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
  return hit ? Number(hit.split("=")[1]) : dflt;
};
const PARAS = arg("paras", 3000);
const KEYS = arg("keys", 30);
const CLICK_EVERY = arg("clickEvery", 0); // 0 = click once, then type straight
// Bookmark every paragraph (the TOC-heading shape). Pre-fix, typing in a
// bookmarked paragraph fell out of the reparse fast path into a full
// doc.refresh() + whole-document relayout per keystroke.
const BOOKMARKS = arg("bookmarks", 0);

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

// Shims mirrored from packages/react/test/setup.ts
globalThis.__measureCalls = 0;
function fakeCtx() {
  return {
    font: "",
    measureText(text) {
      globalThis.__measureCalls++;
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
function docBytes(paras) {
  const para = (i) =>
    `<w:p>${BOOKMARKS ? `<w:bookmarkStart w:id="${i}" w:name="_Toc${1000 + i}"/>` : ""}` +
    `<w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at length. </w:t></w:r>` +
    `${BOOKMARKS ? `<w:bookmarkEnd w:id="${i}"/>` : ""}</w:p>`;
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

/* -------------------------------- run ----------------------------------- */
const pkg = process.env.WW_PKG ?? "wordinweb";
const [{ createElement }, ReactAct, { createRoot }, mod] = await Promise.all([
  import("react"),
  import("react"),
  import("react-dom/client"),
  import(pkg),
]);
const { act } = ReactAct;
const DocxView = mod.DocxView;
if (!DocxView) throw new Error(`no DocxView export in ${pkg}`);

globalThis.__dxwPerf = { samples: [] };
// jsdom has no layout: clientHeight is 0 everywhere, which makes the
// virtualizer mount only the tail window. Give every element a viewport so
// the first pages are mounted and the click hit-test has a page to land on.
Object.defineProperty(window.HTMLElement.prototype, "clientHeight", { get() { return 1200; }, configurable: true });
const container = window.document.createElement("div");
window.document.body.appendChild(container);
const root = createRoot(container);

const tick = async (ms = 5) => { await act(async () => { await new Promise((r) => setTimeout(r, ms)); }); };
const settleUntil = async (cond, budgetMs) => {
  const t0 = performance.now();
  for (;;) {
    if (cond()) return performance.now() - t0;
    if (performance.now() - t0 > budgetMs) return performance.now() - t0;
    await tick(2);
  }
};

const t0 = performance.now();
let renderError = null;
await act(async () => {
  root.render(
    createElement(DocxView, {
      source: docBytes(PARAS),
      editable: true,
      style: { height: "500px" },
      onError: (e) => { renderError = e; },
    }),
  );
});
for (let i = 0; i < 400 && !container.querySelector(".dxw-page"); i++) await tick();
if (!container.querySelector(".dxw-page")) {
  if (renderError) throw renderError;
  throw new Error(`document never painted: ${container.innerHTML.slice(0, 500)}`);
}
const mountMs = performance.now() - t0;
const pages = container.querySelectorAll(".dxw-page").length;

// Place the caret in the first page's first span (a caret jump, like a user click).
const clickSpan = async (idx = 0) => {
  // Virtualization mounts only a window of pages (in jsdom: the tail); click
  // a span on a MOUNTED page — the last ones, matching "edit at the tail".
  const spans = [...container.querySelectorAll(".dxw-page span")];
  const span = spans[Math.min(idx, spans.length - 1)] ?? container.querySelector(".dxw-page");
  await act(async () => {
    const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    span.dispatchEvent(new window.MouseEvent("mousedown", o));
    span.dispatchEvent(new window.MouseEvent("mouseup", o));
  });
};
await clickSpan(0);
await tick(10);
if (process.env.WW_DEBUG) {
  const ae = window.document.activeElement;
  console.log("DEBUG activeElement:", ae?.tagName, ae?.className);
  console.log("DEBUG caretEl:", container.querySelectorAll("[class*=caret]").length,
    "spans:", container.querySelectorAll(".dxw-page span").length,
    "mountedPages:", [...container.querySelectorAll(".dxw-page")].filter((p) => p.childElementCount > 0).length);
}

const target = () =>
  (container.contains(window.document.activeElement) ? window.document.activeElement : container.querySelector("textarea")) ?? container;

const busy = () => !!container.querySelector("[data-dxw-layout-busy]");
let busySeen = 0;

const perKey = [];
const perKeyMeasures = [];
const perKeyAllocMb = [];
const typingHeapStartMb = process.memoryUsage().heapUsed / 1e6;
for (let i = 0; i < KEYS; i++) {
  if (CLICK_EVERY && i > 0 && i % CLICK_EVERY === 0) await clickSpan((i * 7) % 40);
  const m0 = globalThis.__measureCalls;
  const h0 = process.memoryUsage().heapUsed;
  const k0 = performance.now();
  await act(async () => {
    target().dispatchEvent(new window.KeyboardEvent("keydown", { key: "Z", bubbles: true, cancelable: true }));
    await new Promise((r) => setTimeout(r, 0));
  });
  // Settle: run timers/rAF until no background layout is pending.
  await settleUntil(() => !busy(), 30_000);
  if (busy()) busySeen++;
  perKey.push(performance.now() - k0);
  perKeyMeasures.push(globalThis.__measureCalls - m0);
  // Positive delta only: GC between keystrokes makes the raw delta negative;
  // clamping keeps the median honest about allocation per keystroke.
  perKeyAllocMb.push(Math.max(0, (process.memoryUsage().heapUsed - h0) / 1e6));
}

const typed = (container.textContent?.match(/Z/g) ?? []).length;
if (typed < KEYS) console.log(`WARN only ${typed}/${KEYS} keystrokes landed`);
const sorted = [...perKey].sort((a, b) => a - b);
const pct = (p) => sorted[Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length))];
const sum = perKey.reduce((a, b) => a + b, 0);
const samples = globalThis.__dxwPerf.samples ?? [];
const phase = (k) => samples.reduce((a, s) => a + (s[k] ?? 0), 0);
const fmt = (n) => (Number.isInteger(n) ? n : n.toFixed(2));
const allocSorted = [...perKeyAllocMb].sort((a, b) => a - b);
const typingHeapEndMb = process.memoryUsage().heapUsed / 1e6;
const positiveAllocMb = perKeyAllocMb.reduce((total, value) => total + value, 0);
console.log(
  `STRESS-METRIC local-typing pkg=${pkg === "wordinweb" ? "worktree" : "custom"} paragraphs=${PARAS} pages=${pages} ` +
    `bookmarks=${BOOKMARKS} keystrokes=${KEYS} mountMs=${fmt(mountMs)} totalMs=${fmt(sum)} msPerKey=${fmt(sum / KEYS)} ` +
    `p50=${fmt(pct(50))} p90=${fmt(pct(90))} p99=${fmt(pct(99))} max=${fmt(sorted[sorted.length - 1])} busySeen=${busySeen} landed=${typed} ` +
    `measurePerKeyMedian=${[...perKeyMeasures].sort((a, b) => a - b)[Math.floor(perKeyMeasures.length / 2)]} ` +
    `measurePerKeyMax=${Math.max(...perKeyMeasures)} ` +
    `allocMbPerKeyMedian=${fmt(allocSorted[Math.floor(allocSorted.length / 2)])} ` +
    `allocMbPerKeyMax=${fmt(Math.max(...perKeyAllocMb))} ` +
    `heapStartMB=${fmt(typingHeapStartMb)} heapEndMB=${fmt(typingHeapEndMb)} ` +
    `heapGrowthMB=${fmt(typingHeapEndMb - typingHeapStartMb)} ` +
    `positiveAllocMB=${fmt(positiveAllocMb)} allocMBps=${fmt(positiveAllocMb / (sum / 1000))}`,
);
if (samples.length) {
  console.log(
    `STRESS-METRIC local-typing-phases commits=${samples.length} refreshMs=${fmt(phase("refresh"))} ` +
      `layoutMs=${fmt(phase("layout"))} renderMs=${fmt(phase("render"))} totalMs=${fmt(phase("total"))}`,
  );
}
await act(async () => { root.unmount(); });
process.exit(0);
