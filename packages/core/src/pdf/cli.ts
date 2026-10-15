// Render a .docx to PDF without a browser:
//   node packages/core/dist/pdf/cli.js <in.docx> <out.pdf>
import { createWriteStream } from "node:fs";
import { renderDocxToPdf } from "./render.js";

const [input, output] = process.argv.slice(2);
if (!input || !output) {
  console.error("usage: docx-to-pdf <in.docx> <out.pdf>");
  process.exit(2);
}
const t0 = performance.now();
const result = await renderDocxToPdf(input, createWriteStream(output));
const ms = Math.round(performance.now() - t0);
const rss = Math.round(process.resourceUsage().maxRSS / 1024);
console.log(JSON.stringify({ input, output, pages: result.pages, ms, maxRssMB: rss, fonts: result.fonts, missingFonts: result.missingFonts, unsupportedImages: result.unsupportedImages, undrawn: result.undrawn }));
