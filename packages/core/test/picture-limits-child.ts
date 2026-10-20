// vite-node picture-limits-child.ts PATH: render the .docx at PATH and print,
// as JSON, the pictures it reported unsupported, its pages, the peak RSS and
// the peak array buffers and external memory in MB (typed arrays live
// outside the V8 heap, which --max-old-space-size does not cap), and the
// time. A child process of its own, so the peaks are this render's
// (picture-limits.test.ts).
import { Writable } from "node:stream";
import { fileSource, renderDocxToPdf } from "../src/pdf/render.js";

const MB = 1024 * 1024;
const [path] = process.argv.slice(2);
let arrayBuffers = 0;
let external = 0;
const sample = (): void => {
  const usage = process.memoryUsage();
  arrayBuffers = Math.max(arrayBuffers, usage.arrayBuffers);
  external = Math.max(external, usage.external);
};
const sampler = setInterval(sample, 5);
const began = performance.now();
const sink = new Writable({ write: (_chunk, _encoding, done) => done() });
const result = await renderDocxToPdf(fileSource(path), sink);
clearInterval(sampler);
sample();
console.log(
  JSON.stringify({
    unsupportedImages: result.unsupportedImages,
    pages: result.pages,
    peakRssMB: Math.round((process.resourceUsage().maxRSS * 1024) / MB),
    arrayBuffersMB: Math.round(arrayBuffers / MB),
    externalMB: Math.round(external / MB),
    ms: Math.round(performance.now() - began),
  }),
);
