#!/usr/bin/env node
/**
 * REMOTE-EDIT REPAINT BENCHMARK — jsdom, real CollabEditor watcher + a second
 * client typing through the real hub (the owner scenario: user X watches a
 * huge document while user Y edits far away).
 *
 * Measures what ONE remote keystroke costs the WATCHING client: canvas
 * measureText calls (the causal O(document) work), wall ms to painted, and
 * whether the input-blocking background layout ever engaged. A structural
 * remote edit (splitParagraph = Enter) is measured separately — it genuinely
 * reflows forward and is allowed to cost more; it is reported, not averaged
 * away.
 *
 * Run:   node scripts/bench-remote-typing.mjs [--paras=2400] [--keys=30] [--splits=8]
 *        node scripts/bench-remote-typing.mjs --forceGlobal=1   # pre-fix behavior
 * Needs: npm run build -w @wordinweb/core -w @wordinweb/collab -w @wordinweb/server -w wordinweb
 *
 * --forceGlobal=1 replaces the session's takeRenderScope with a constant
 * document scope — exactly the pre-fix repaint path — so the before/after
 * comparison runs on the same instrument.
 */

import { JSDOM } from "jsdom";
import { zipSync, strToU8 } from "fflate";

const arg = (name, dflt) => {
  const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
  return hit ? Number(hit.split("=")[1]) : dflt;
};
const PARAS = arg("paras", 2400);
const KEYS = arg("keys", 30);
const SPLITS = arg("splits", 8);
const FORCE_GLOBAL = arg("forceGlobal", 0) === 1;

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

// Shims mirrored from packages/react/test/setup.ts + measureText counting.
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
// Give every element a viewport so the first pages mount and the edited
// paragraph's DOM exists (jsdom clientHeight is otherwise 0 everywhere).
Object.defineProperty(window.HTMLElement.prototype, "clientHeight", { get() { return 1200; }, configurable: true });

/* ------------------------------ doc bytes ------------------------------- */
function docBytes(paras) {
  const para = (i) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at length. </w:t></w:r></w:p>`;
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
const [{ createElement, act }, { createRoot }, reactMod, collabMod, serverMod, coreMod, collabClientMod] = await Promise.all([
  import("react"),
  import("react-dom/client"),
  import("wordinweb"),
  import("wordinweb/collab"),
  import("@wordinweb/server"),
  import("@wordinweb/core"),
  import("@wordinweb/collab/client"),
]);
const { DocxView } = reactMod;
const { useCollab } = collabMod;
const { CollabHub } = serverMod;
const { serializeXml } = coreMod;
const { CollabConnection, createWebSocketTransport } = collabClientMod;

// A thin watcher: useCollab + DocxView, the same subset CollabEditor passes.
// Owning the wiring here is what lets --forceGlobal swap the scope feed for
// the pre-fix constant-doc behavior on an otherwise identical instrument.
const PLACEHOLDER_SOURCE = new Uint8Array(0); // stable identity — a fresh one per render re-runs the load effect
function Watcher({ createSocket }) {
  const session = useCollab({ url: "ws://x", docId: "big", clientId: "watcher", createSocket });
  if (!session.ready || !session.doc) return createElement("div", null, "connecting");
  return createElement(DocxView, {
    source: PLACEHOLDER_SOURCE, // placeholder; the live collab doc renders
    editable: true,
    collab: {
      submit: session.submit,
      submitOp: session.submitOp,
      presence: session.presence,
      allocIds: session.allocIds,
      doc: session.doc,
      renderSignal: session.renderVersion,
      takeRenderScope: FORCE_GLOBAL ? () => ({ kind: "doc" }) : session.takeRenderScope,
      setPresence: session.setPresence,
    },
  });
}

const hub = new CollabHub({ load: () => docBytes(PARAS) });
function factoryFor() {
  let n = 0;
  return (_url) => {
    const ls = [];
    const conn = { id: `b${n++}`, send: (m) => ls.forEach((l) => l({ data: JSON.stringify(m) })) };
    let opened = false;
    return {
      send: (d) => { void hub.handle(conn, JSON.parse(d)); },
      addEventListener: (t, cb) => { if (t === "message") ls.push(cb); else if (!opened) { opened = true; cb(); } },
    };
  };
}
const factory = factoryFor();

const tick = async (ms = 5) => { await act(async () => { await new Promise((r) => setTimeout(r, ms)); }); };
async function until(cond, label, tries = 2000) {
  for (let i = 0; i < tries; i++) {
    if (cond()) return;
    await tick(2);
  }
  throw new Error(`timeout: ${label}`);
}

globalThis.__dxwPerf = {};
const container = window.document.createElement("div");
window.document.body.appendChild(container);
const root = createRoot(container);

const t0 = performance.now();
await act(async () => {
  root.render(createElement(Watcher, { createSocket: factory }));
});
await until(() => !!container.querySelector(".dxw-page"), "watcher paints");
const mountMs = performance.now() - t0;
const mountMeasures = globalThis.__measureCalls;
const pages = container.querySelectorAll(".dxw-page").length;

// Track input-blocking background layout windows on the watcher.
let busySeen = 0;
new window.MutationObserver((records) => {
  for (const r of records) {
    if (r.target.hasAttribute?.("data-dxw-layout-busy")) busySeen++;
  }
}).observe(container, { subtree: true, attributes: true, attributeFilter: ["data-dxw-layout-busy"] });

// The remote typist edits paragraph 2 — page 1 of a ~100-page document.
const editor = new CollabConnection(createWebSocketTransport(factory("ws://x")), "typist");
editor.join("big");
await until(() => editor.ready, "typist welcome");
const para = editor.doc.sections[0].blocks[2];
const run = para.children[0];
const srcT = run.content.find((c) => c.kind === "text").srcT;
const ids = editor.doc.stableIds;
const addr = { blockId: ids.idOf(para.src), runId: ids.idOf(run.src), len: srcT.text.length };

// The watcher's live doc, through the React fiber tree (same technique as
// packages/react/test — the session is a prop of DocxView).
function watcherDoc() {
  const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"));
  const stack = [container[key]];
  let guard = 0;
  while (stack.length && guard++ < 20000) {
    const f = stack.pop();
    if (!f) continue;
    const c = f.memoizedProps?.collab;
    if (c && c.presence && typeof c.presence === "object") return c.doc;
    if (f.child) stack.push(f.child);
    if (f.sibling) stack.push(f.sibling);
  }
  throw new Error("watcher session not found");
}

/* ---- remote TEXT keystrokes (the hot path) ---- */
const perKeyMs = [];
const perKeyMeasures = [];
for (let k = 0; k < KEYS; k++) {
  const marker = `X${k}Q`;
  const m0 = globalThis.__measureCalls;
  const k0 = performance.now();
  await act(async () => {
    editor.submit({
      kind: "insertText",
      at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len + "X0Q".length * k },
      text: marker,
    });
  });
  await until(() => !!container.textContent?.includes(marker), `watcher paints ${marker}`);
  perKeyMs.push(performance.now() - k0);
  perKeyMeasures.push(globalThis.__measureCalls - m0);
}

/* ---- remote STRUCTURAL edits (Enter far from the watcher) ---- */
const perSplitMs = [];
const perSplitMeasures = [];
for (let s = 0; s < SPLITS; s++) {
  const [newBlockId, newRunId] = editor.allocIds(2);
  const before = watcherDoc().sections[0].blocks.length;
  const m0 = globalThis.__measureCalls;
  const k0 = performance.now();
  await act(async () => {
    editor.submit({
      kind: "splitParagraph",
      at: { blockId: addr.blockId, runId: addr.runId, offset: 2 },
      newBlockId,
      newRunId,
    });
  });
  await until(() => watcherDoc().sections[0].blocks.length === before + 1, `watcher applies split ${s}`);
  await tick(80); // let any queued background layout land inside the window
  perSplitMs.push(performance.now() - k0);
  perSplitMeasures.push(globalThis.__measureCalls - m0);
}
await tick(120);

/* ---- convergence: watcher == typist == fresh joiner (server canonical) ---- */
const late = new CollabConnection(createWebSocketTransport(factory("ws://x")), "late");
late.join("big");
await until(() => late.ready, "late welcome");
const watcherXml = serializeXml(watcherDoc().docRoot);
const converged = watcherXml === serializeXml(editor.doc.docRoot) && watcherXml === serializeXml(late.doc.docRoot);

const stats = (arr) => {
  const sorted = [...arr].sort((a, b) => a - b);
  const pct = (p) => sorted[Math.min(sorted.length - 1, Math.floor((p / 100) * sorted.length))];
  return { p50: pct(50), p90: pct(90), max: sorted[sorted.length - 1], sum: arr.reduce((a, b) => a + b, 0) };
};
const fmt = (n) => (Number.isInteger(n) ? n : n.toFixed(2));
const mode = FORCE_GLOBAL ? "remote-typing-forced-global" : "remote-typing";
const ms = stats(perKeyMs);
const mm = stats(perKeyMeasures);
console.log(
  `STRESS-METRIC ${mode} paragraphs=${PARAS} pages=${pages} keystrokes=${KEYS} mountMs=${fmt(mountMs)} ` +
    `mountMeasures=${mountMeasures} msPerKey=${fmt(ms.sum / KEYS)} p50=${fmt(ms.p50)} p90=${fmt(ms.p90)} max=${fmt(ms.max)} ` +
    `measurePerKeyMedian=${mm.p50} measurePerKeyMax=${mm.max} busySeen=${busySeen} converged=${converged}`,
);
const sms = stats(perSplitMs);
const smm = stats(perSplitMeasures);
console.log(
  `STRESS-METRIC ${mode}-structural paragraphs=${PARAS} splits=${SPLITS} msPerSplit=${fmt(sms.sum / SPLITS)} ` +
    `p50=${fmt(sms.p50)} max=${fmt(sms.max)} measurePerSplitMedian=${smm.p50} measurePerSplitMax=${smm.max} busySeen=${busySeen} converged=${converged}`,
);
if (!converged) {
  console.error("DIVERGED: watcher/typist/late replicas are not byte-identical");
  process.exit(2);
}
await act(async () => { root.unmount(); });
process.exit(0);
