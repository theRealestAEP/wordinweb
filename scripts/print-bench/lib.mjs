// Shared driver for the print benches: the old renderForPrint path (as
// DiscoveryTool's chrome.ts drives it) and the new streaming path, both in
// headless Chromium through Playwright, plus memory sampling.
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { buildSync } from "esbuild";
import { chromium } from "playwright";

/**
 * The print bundle, built from this checkout's packages/core as DiscoveryTool's
 * vendor/wordinweb/rebuild.sh builds print.js: a browser IIFE, global wordinwebPrint.
 */
export const BUNDLE = buildSync({
  stdin: {
    contents: 'export { renderForPrint, openPrintStream } from "./render/print.js";',
    resolveDir: fileURLToPath(new URL("../../packages/core/src/", import.meta.url)),
    loader: "ts",
  },
  bundle: true,
  format: "iife",
  globalName: "wordinwebPrint",
  platform: "browser",
  write: false,
}).outputFiles[0].text;

/** In-page helpers: a stable hash of each rendered page's DOM. */
const PAGE_HASHES = `(() => {
  const hash = (s) => {
    let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (let i = 0; i < s.length; i++) { const ch = s.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  };
  const norm = (html) => html.replace(/blob:[^"'\\s)]+/g, "blob:");
  return [...document.querySelectorAll(".dxw-page")].map((el) => hash(norm(el.outerHTML)));
})()`;

const PAGE_HTML = (index) => `(() => {
  const el = document.querySelectorAll(".dxw-page")[${index}];
  return el ? el.outerHTML.replace(/blob:[^"'\\s)]+/g, "blob:") : null;
})()`;

export async function launch() {
  return chromium.launch({ headless: true, args: ["--disable-gpu", "--no-first-run"] });
}

/** Peak renderer RSS (MB) and JS heap (MB) sampled while `work` runs. */
export async function sampled(browser, page, work, trace = null) {
  const session = await page.context().newCDPSession(page);
  await session.send("Performance.enable");
  const browserSession = await browser.newBrowserCDPSession();
  let peakRss = 0;
  let peakHeap = 0;
  let stopped = false;
  const t0 = performance.now();
  const sample = async () => {
    let heap = 0;
    let heapTotal = 0;
    let rss = 0;
    try {
      const { metrics } = await session.send("Performance.getMetrics");
      heap = (metrics.find((m) => m.name === "JSHeapUsedSize")?.value ?? 0) / 1048576;
      heapTotal = (metrics.find((m) => m.name === "JSHeapTotalSize")?.value ?? 0) / 1048576;
      peakHeap = Math.max(peakHeap, heap);
    } catch {}
    try {
      // Chromium names its own processes; the page's renderer is the only one.
      const { processInfo } = await browserSession.send("SystemInfo.getProcessInfo");
      const renderers = processInfo.filter((p) => p.type === "renderer").map((p) => String(p.id));
      if (renderers.length === 0) return;
      const out = execFileSync("ps", ["-o", "rss=", "-p", renderers.join(",")], { encoding: "utf8" });
      for (const line of out.split("\n")) {
        const kb = Number(line.trim());
        if (kb > 0) rss = Math.max(rss, kb / 1024);
      }
      peakRss = Math.max(peakRss, rss);
    } catch {}
    trace?.push({ t: Math.round(performance.now() - t0), rssMB: Math.round(rss), heapMB: Math.round(heap), heapTotalMB: Math.round(heapTotal) });
  };
  const timer = setInterval(() => { if (!stopped) void sample(); }, 250);
  try {
    const result = await work();
    await sample();
    return { result, peakRssMB: Math.round(peakRss), peakHeapMB: Math.round(peakHeap) };
  } finally {
    stopped = true;
    clearInterval(timer);
    await session.detach().catch(() => {});
  }
}

/** The old path: renderForPrint over the whole document, then (as DiscoveryTool
 * prints) the pages in slices of `batch`. `onPages(hashes)` sees every page's
 * hash once rendered; `print` prints each slice to PDF when set. */
/** Both paths see the same wall clock, so DATE and TIME fields agree. */
const FIXED_TIME = new Date("2026-10-15T05:46:50");

export async function runOld(page, bytes, { batch = 25, print = false, onPages } = {}) {
  await page.goto("about:blank");
  await page.clock.setFixedTime(FIXED_TIME);
  await page.addScriptTag({ content: BUNDLE });
  const base64 = Buffer.from(bytes).toString("base64");
  const t0 = performance.now();
  const { pageCount } = await page.evaluate(`(async () => {
    const bytes = Uint8Array.from(atob("${base64}"), (c) => c.charCodeAt(0));
    const root = document.body.appendChild(document.createElement("div"));
    return await wordinwebPrint.renderForPrint(bytes, root);
  })()`);
  const layoutMs = Math.round(performance.now() - t0);
  const hashes = await page.evaluate(PAGE_HASHES);
  onPages?.(hashes);
  let printMs = 0;
  if (print) {
    const t1 = performance.now();
    await page.evaluate(`window.printPages = [...document.querySelectorAll(".dxw-page")]; for (const p of window.printPages) p.remove();`);
    for (let first = 0; first < Math.max(pageCount, 1); first += batch) {
      await page.evaluate(`document.querySelector(".dxw-pages").replaceChildren(...window.printPages.slice(${first}, ${first + batch}));`);
      await page.pdf({ preferCSSPageSize: true, printBackground: true });
    }
    printMs = Math.round(performance.now() - t1);
  }
  return { pageCount, hashes, layoutMs, printMs, pageHtml: (i) => page.evaluate(PAGE_HTML(i)) };
}

/** Serve `bytes` at http://docx.local/doc.docx with Range support, and an empty page at /. */
export async function serve(page, bytes) {
  await page.route("http://docx.local/**", async (route) => {
    const url = new URL(route.request().url());
    if (url.pathname === "/") return route.fulfill({ status: 200, contentType: "text/html", body: "<!doctype html><html><head></head><body></body></html>" });
    if (url.pathname !== "/doc.docx") return route.fulfill({ status: 404, body: "" });
    const range = /bytes=(\d+)-(\d+)/.exec(route.request().headers()["range"] ?? "");
    if (!range) return route.fulfill({ status: 200, contentType: "application/octet-stream", body: Buffer.from(bytes) });
    const start = Number(range[1]);
    const end = Math.min(Number(range[2]), bytes.length - 1);
    return route.fulfill({
      status: 206,
      headers: { "Content-Range": `bytes ${start}-${end}/${bytes.length}`, "Accept-Ranges": "bytes", "Content-Type": "application/octet-stream" },
      body: Buffer.from(bytes.subarray(start, end + 1)),
    });
  });
}

/** The new path: openPrintStream over the served file, a batch at a time.
 * `onBatch(first, hashes, batchApi)` sees each mounted batch before release. */
export async function runNew(page, bytes, { batch = 25, print = false, onBatch } = {}) {
  await serve(page, bytes);
  await page.goto("http://docx.local/");
  await page.clock.setFixedTime(FIXED_TIME);
  await page.addScriptTag({ content: BUNDLE });
  const t0 = performance.now();
  const started = await page.evaluate(`(async () => {
    window.__print = wordinwebPrint.openPrintStream("http://docx.local/doc.docx", document.body, { batchPages: ${batch} });
    return await window.__print.start();
  })()`);
  const countMs = Math.round(performance.now() - t0);
  const hashes = [];
  let printMs = 0;
  let renderMs = 0;
  let batches = 0;
  for (;;) {
    const t1 = performance.now();
    const b = await page.evaluate(`window.__print.nextBatch()`);
    renderMs += performance.now() - t1;
    if (!b) break;
    batches++;
    const batchHashes = await page.evaluate(PAGE_HASHES);
    if (batchHashes.length !== b.count) throw new Error(`batch at ${b.first}: ${batchHashes.length} pages mounted, ${b.count} announced`);
    hashes.push(...batchHashes);
    if (onBatch) await onBatch(b.first, batchHashes, { pageHtml: (i) => page.evaluate(PAGE_HTML(i - b.first)) });
    if (print) {
      const t2 = performance.now();
      await page.pdf({ preferCSSPageSize: true, printBackground: true });
      printMs += performance.now() - t2;
    }
    await page.evaluate(`window.__print.release()`);
  }
  return { ...started, hashes, batches, countMs, renderMs: Math.round(renderMs), printMs: Math.round(printMs) };
}

/** Where two normalized HTML strings first differ, with context. */
export function firstDiff(a, b) {
  if (a === b) return null;
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  const from = Math.max(0, i - 120);
  return { at: i, old: a.slice(from, i + 160), new: b.slice(from, i + 160) };
}
