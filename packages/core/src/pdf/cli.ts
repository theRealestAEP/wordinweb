#!/usr/bin/env node
// Render a .docx to PDF without a browser.
import { createWriteStream, rmSync } from "node:fs";
import { PdfRenderError, PdfRenderErrorCode } from "./errors.js";
import { renderDocxToPdf } from "./render.js";

const USAGE = `Usage: docx-pdf IN.docx OUT.pdf [--comments] [--markup] [--hide-authors] [--fonts DIR]...

Write OUT.pdf: IN.docx laid out and drawn as Word prints it, with the
fonts embedded, no browser involved. Prints one JSON line about the run.
A Strict Open XML Document (ISO/IEC 29500 Strict) is read as Transitional.
A Word 97-2003 .doc (the binary format) is read too, converted on the way
in; the file's bytes decide, not its name.

Options:
  --comments      print the review comments beside each page, as Word's markup
  --markup        print the tracked changes as Word's All Markup: insertions
                  underlined and deletions struck in each author's color, a
                  balloon per change beside the page (with the comments)
  --hide-authors  name no author anywhere: each balloon keeps its date, what
                  changed and the comment text, with no name and no initials,
                  and every author's changes are red
  --fonts DIR     also scan DIR for fonts, ahead of the system's directories
                  (repeatable; WORDINWEB_FONT_DIRS adds directories the same way)
  -h, --help      show this help

Exit codes: 0 done; 1 failed; 2 usage; 3 unsupported_script (text in a script
the renderer cannot shape); 4 not_docx; 5 damaged; 6 unsupported_format (a
password-protected document, an RTF, PDF or OpenDocument file); 7
unsupported_content (content Word prints that the renderer does not read, such
as a w:altChunk: an HTML or RTF part Word imports on open); 8 too_large (the
document is over the size limits).`;

const EXIT: Record<PdfRenderErrorCode, number> = { failed: 1, unsupported_script: 3, not_docx: 4, damaged: 5, unsupported_format: 6, unsupported_content: 7, too_large: 8 };

const args = process.argv.slice(2);
if (args.includes("-h") || args.includes("--help")) {
  console.log(USAGE);
  process.exit(0);
}
const comments = args.includes("--comments");
const markup = args.includes("--markup");
const authors = !args.includes("--hide-authors");
const extraFontDirs: string[] = [];
const files: string[] = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] === "--comments" || args[i] === "--markup" || args[i] === "--hide-authors") continue;
  if (args[i] === "--fonts") {
    if (i + 1 >= args.length) {
      console.error("docx-pdf: --fonts needs a directory\nRun docx-pdf --help for the options.");
      process.exit(2);
    }
    extraFontDirs.push(args[++i]);
    continue;
  }
  files.push(args[i]);
}
const unknown = files.find((arg) => arg.startsWith("-"));
const [input, output, extra] = files;
if (unknown || !input || !output || extra) {
  console.error(`docx-pdf: ${unknown ? `unknown option ${unknown}` : extra ? `unexpected argument ${extra}` : "give the document and the PDF to write"}\nRun docx-pdf --help for the options.`);
  process.exit(2);
}
const t0 = performance.now();
try {
  const result = await renderDocxToPdf(input, createWriteStream(output), { comments, markup, authors, extraFontDirs });
  const ms = Math.round(performance.now() - t0);
  const rss = Math.round(process.resourceUsage().maxRSS / 1024);
  console.log(JSON.stringify({ input, output, pages: result.pages, ms, maxRssMB: rss, fonts: result.fonts, missingFonts: result.missingFonts, missingCharacters: result.missingCharacters, unsupportedImages: result.unsupportedImages }));
} catch (error) {
  const code: PdfRenderErrorCode = error instanceof PdfRenderError ? error.code : "failed";
  console.error(`docx-pdf: ${code}: ${error instanceof Error ? error.message : String(error)}`);
  // No half-written PDF is left behind.
  rmSync(output, { force: true });
  process.exitCode = EXIT[code];
}
