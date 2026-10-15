#!/usr/bin/env node
// Render a .docx to PDF without a browser.
import { createWriteStream, rmSync } from "node:fs";
import { PdfRenderError, PdfRenderErrorCode } from "./errors.js";
import { renderDocxToPdf } from "./render.js";

const USAGE = `Usage: docx-pdf IN.docx OUT.pdf [--comments]

Write OUT.pdf: IN.docx laid out and drawn as Word prints it, with the
fonts embedded, no browser involved. Prints one JSON line about the run.

Options:
  --comments    print the review comments beside each page, as Word's markup
  -h, --help    show this help

Exit codes: 0 done; 1 failed; 2 usage; 3 unsupported_script (text in a script
the renderer cannot shape); 4 not_docx; 5 damaged.`;

const EXIT: Record<PdfRenderErrorCode, number> = { failed: 1, unsupported_script: 3, not_docx: 4, damaged: 5 };

const args = process.argv.slice(2);
if (args.includes("-h") || args.includes("--help")) {
  console.log(USAGE);
  process.exit(0);
}
const comments = args.includes("--comments");
const files = args.filter((arg) => arg !== "--comments");
const unknown = files.find((arg) => arg.startsWith("-"));
const [input, output, extra] = files;
if (unknown || !input || !output || extra) {
  console.error(`docx-pdf: ${unknown ? `unknown option ${unknown}` : extra ? `unexpected argument ${extra}` : "give the document and the PDF to write"}\nRun docx-pdf --help for the options.`);
  process.exit(2);
}
const t0 = performance.now();
try {
  const result = await renderDocxToPdf(input, createWriteStream(output), { comments });
  const ms = Math.round(performance.now() - t0);
  const rss = Math.round(process.resourceUsage().maxRSS / 1024);
  console.log(JSON.stringify({ input, output, pages: result.pages, ms, maxRssMB: rss, fonts: result.fonts, missingFonts: result.missingFonts, unsupportedImages: result.unsupportedImages }));
} catch (error) {
  const code: PdfRenderErrorCode = error instanceof PdfRenderError ? error.code : "failed";
  console.error(`docx-pdf: ${code}: ${error instanceof Error ? error.message : String(error)}`);
  // No half-written PDF is left behind.
  rmSync(output, { force: true });
  process.exitCode = EXIT[code];
}
