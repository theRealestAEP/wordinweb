// vite-node zip-limits-child.ts ORIGINAL REVISED OUT: compare the two
// .docx files and print, as JSON, the error compareFiles gives and how far
// the peak RSS rose above the RSS at the start, in MB. A child process of
// its own, so the peak is this compare's (compare-files.test.ts).
import { compareFiles } from "../src/compare.js";
import { CompareError } from "../src/errors.js";

const [original, revised, out] = process.argv.slice(2);
const start = process.memoryUsage().rss;
const began = performance.now();
let error: { code: string; message: string } | null = null;
try {
  await compareFiles(original, revised, out, { author: "A" });
} catch (caught) {
  const e = caught as CompareError;
  error = { code: e.code, message: e.message };
}
const grewMB = Math.round((process.resourceUsage().maxRSS * 1024 - start) / (1024 * 1024));
console.log(JSON.stringify({ error, grewMB, ms: Math.round(performance.now() - began) }));
