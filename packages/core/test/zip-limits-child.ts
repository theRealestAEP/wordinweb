// vite-node zip-limits-child.ts PATH WAY: open the .docx at PATH one WAY
// (render, load or estimate) and print, as JSON, the error it gives and how
// far the peak RSS rose above the RSS at the start, in MB. A child process
// of its own, so the peak is this one opening's (zip-limits.test.ts).
import { readFileSync } from "node:fs";
import { Writable } from "node:stream";
import { DocxDocument } from "../src/docx.js";
import { estimateCompareMemory } from "../src/edit/compare/memory.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";

const [path, way] = process.argv.slice(2);
const bytes = new Uint8Array(readFileSync(path));
const start = process.memoryUsage().rss;
const began = performance.now();
let error: { name: string; code?: string; message: string } | null = null;
try {
  if (way === "render") {
    const sink = new Writable({ write: (_chunk, _encoding, done) => done() });
    await renderDocxToPdf(bytesSource(bytes), sink);
  } else if (way === "load") DocxDocument.load(bytes);
  else if (way === "estimate") estimateCompareMemory(bytes, bytes);
  else throw new Error(`no way ${way}`);
} catch (caught) {
  const e = caught as Error & { code?: string };
  error = { name: e.name, code: e.code, message: e.message };
}
const grewMB = Math.round((process.resourceUsage().maxRSS * 1024 - start) / (1024 * 1024));
console.log(JSON.stringify({ error, grewMB, ms: Math.round(performance.now() - began) }));
