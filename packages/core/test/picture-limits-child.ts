// node --expose-gc vite-node picture-limits-child.ts CONTROL TARGET: render
// the document CONTROL (the same document with no picture) until the process
// is warm, then CONTROL and TARGET once each, each render after a full
// garbage collection, sampling RSS, array buffers and external memory every
// 5 ms. Prints, as JSON, TARGET's pages and unsupported pictures, its time,
// how far each of its peaks rose over CONTROL's, and the largest typed array
// each render made. Typed arrays live outside the V8 heap, which
// --max-old-space-size does not cap, so array buffers and external memory
// are measured as well as RSS. The samples are as good as their timing: a
// font file each render reads (about 55 MB) is in some samples and not
// others. The largest array is exact (picture-limits.test.ts).
import { Writable } from "node:stream";

// The largest typed array or ArrayBuffer the render makes: every
// constructor is wrapped before the renderer is imported, so the modules it
// loads (fflate's included) make theirs through the wrappers. A picture over
// the budget must never make one of its bitmap's size.
let largest = 0;
const typed = globalThis as unknown as Record<string, new (...args: unknown[]) => { byteLength: number }>;
for (const name of ["ArrayBuffer", "Uint8Array", "Uint8ClampedArray", "Int8Array", "Uint16Array", "Int16Array", "Uint32Array", "Int32Array", "Float32Array", "Float64Array"]) {
  typed[name] = new Proxy(typed[name], {
    construct(target, args, newTarget) {
      const made = Reflect.construct(target, args, newTarget) as { byteLength: number };
      largest = Math.max(largest, made.byteLength);
      return made;
    },
  });
}
const { fileSource, renderDocxToPdf } = await import("../src/pdf/render.js");

const MB = 1024 * 1024;
const [control, target] = process.argv.slice(2);
const gc = (globalThis as { gc?: () => void }).gc as () => void;

interface Peaks {
  rss: number;
  arrayBuffers: number;
  external: number;
}

async function measure(path: string): Promise<{ peaks: Peaks; largest: number; pages: number; unsupportedImages: string[]; ms: number }> {
  gc();
  largest = 0;
  const peaks: Peaks = { rss: 0, arrayBuffers: 0, external: 0 };
  const sample = (): void => {
    const usage = process.memoryUsage();
    peaks.rss = Math.max(peaks.rss, usage.rss);
    peaks.arrayBuffers = Math.max(peaks.arrayBuffers, usage.arrayBuffers);
    peaks.external = Math.max(peaks.external, usage.external);
  };
  const sampler = setInterval(sample, 5);
  const began = performance.now();
  const result = await renderDocxToPdf(fileSource(path), new Writable({ write: (_chunk, _encoding, done) => done() }));
  const ms = performance.now() - began;
  clearInterval(sampler);
  sample();
  return { peaks, largest, pages: result.pages, unsupportedImages: result.unsupportedImages, ms };
}

await measure(control);
const before = await measure(control);
const after = await measure(target);
const rise = (key: keyof Peaks): number => Math.round((after.peaks[key] - before.peaks[key]) / MB);
console.log(
  JSON.stringify({
    unsupportedImages: after.unsupportedImages,
    pages: after.pages,
    controlPages: before.pages,
    rssRiseMB: rise("rss"),
    arrayBuffersRiseMB: rise("arrayBuffers"),
    externalRiseMB: rise("external"),
    peakRssMB: Math.round(after.peaks.rss / MB),
    largestArrayMB: Math.round(after.largest / MB),
    controlLargestArrayMB: Math.round(before.largest / MB),
    ms: Math.round(after.ms),
  }),
);
