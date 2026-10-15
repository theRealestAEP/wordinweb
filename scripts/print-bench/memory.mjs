// One print in a fresh headless Chromium, with its memory sampled: the peak
// renderer RSS and JS heap while laying out and printing to PDF, the peak
// RSS of the browser process, and of this Node process (which serves the
// file's bytes, as DiscoveryTool does).
//   node memory.mjs --mode old|new [--batch 25] [--print 0|1] [--trace samples.json] <docx>
// Prints one JSON line. The new mode serves the file from disk, a range at a
// time; the old mode needs its bytes whole.
import { readFileSync, writeFileSync } from "node:fs";
import { basename } from "node:path";
import { launch, runNew, runOld, sampled } from "./lib.mjs";

const args = process.argv.slice(2);
const opt = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const value = args[i + 1];
  args.splice(i, 2);
  return value;
};
const mode = opt("mode", "new");
const batch = Number(opt("batch", "25"));
const print = opt("print", "1") === "1";
const traceFile = opt("trace", null);
const file = args[0];

const browser = await launch();
const page = await browser.newPage();
const t0 = performance.now();
const trace = traceFile ? [] : null;
let out;
try {
  const { result, peakRssMB, peakHeapMB, browserRssMB } = await sampled(
    browser,
    page,
    () => (mode === "old" ? runOld(page, new Uint8Array(readFileSync(file)), { batch, print }) : runNew(page, file, { batch, print })),
    trace,
  );
  out = {
    name: basename(file),
    mode,
    pages: result.pageCount,
    peakRssMB,
    peakHeapMB,
    browserRssMB,
    nodeRssMB: Math.round(process.resourceUsage().maxRSS / 1024),
    images: result.images,
    imagesDrawn: result.imagesDrawn,
    ms: Math.round(performance.now() - t0),
    layoutMs: result.layoutMs ?? result.countMs,
    renderMs: result.renderMs,
    printMs: result.printMs,
  };
} catch (error) {
  out = { name: basename(file), mode, error: String(error && error.message ? error.message : error).slice(0, 200), nodeRssMB: Math.round(process.resourceUsage().maxRSS / 1024), ms: Math.round(performance.now() - t0) };
}
if (traceFile) writeFileSync(traceFile, JSON.stringify(trace));
await browser.close().catch(() => {});
console.log(JSON.stringify(out));
