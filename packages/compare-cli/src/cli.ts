#!/usr/bin/env node
import { mkdtempSync, renameSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { compareFiles } from "./compare.js";
import { CompareError } from "./errors.js";

const USAGE = `Usage: docx-compare ORIGINAL.docx REVISED.docx -o RESULT.docx --author NAME [options]

Write RESULT.docx: REVISED.docx with every difference from ORIGINAL.docx as a
tracked change by NAME. Accept all of them to get the revised document;
reject all of them to get the original.

A Strict Open XML Document (ISO/IEC 29500 Strict) is read as Transitional,
and RESULT.docx is written as Transitional either way. Either input may be
a Word 97-2003 .doc (the binary format), or a document of Word for the
Macintosh 1.x-5.x or Word for Windows 1.x/2.0: it is converted to a .docx
working file first, so a .doc can be compared with a .docx. The file's
bytes decide, not its name.

The body, the notes and the binary parts (images, embedded files) can be any
size. The other parts (styles, numbering, headers, footers, charts, custom
XML) are held in memory, at most 3 MB together, so that a compare fits in a
256 MB heap; over that, the compare fails.

Options:
  -o, --out FILE      the result (required)
  --author NAME       who the tracked changes are attributed to (required)
  --date ISO          the date on every tracked change, such as 2026-10-14T09:30:00Z (default: now)
  --no-formatting     do not report formatting-only differences
  --no-whitespace     do not report whitespace-only differences
  --temp-dir DIR      where the working files go (default: the system's)
  --pdf FILE          also write the redline as a PDF: the result with its tracked
                      changes as plain formatting, rendered without a browser
  --pdf-comments      print the review comments beside the PDF's pages
  --hide-authors      name no author in the PDF: each comment keeps its date
                      and text, with no name and no initials
  --quiet             do not print notes about differences that could not be expressed exactly
  -h, --help          show this help

Exit codes: 0 done; 1 failed (the message starts with the reason's code, such as
not_docx, unsupported_format for an RTF, PDF, OpenDocument, Word 6.0/95 or
password-protected file, unsupported_content for content Word prints that the
compare does not read, such as a w:altChunk (an HTML or RTF part Word imports
on open), or unsupported_script); 2 usage.`;

/** A mistake in the command line: reported with exit code 2. */
class UsageError extends Error {}

const VALUED = new Set(["out", "author", "date", "temp-dir", "pdf"]);
const SWITCHES = new Set(["no-formatting", "no-whitespace", "pdf-comments", "hide-authors", "quiet", "help"]);

function parseArgs(argv: string[]): { files: string[]; flags: Record<string, string | boolean> } {
  const files: string[] = [];
  const flags: Record<string, string | boolean> = {};
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (!arg.startsWith("-") || arg === "-") {
      files.push(arg);
      continue;
    }
    const key = arg === "-o" ? "out" : arg === "-h" ? "help" : arg.startsWith("--") ? arg.slice(2) : "";
    if (VALUED.has(key)) {
      if (i + 1 >= argv.length) throw new UsageError(`${arg} needs a value`);
      flags[key] = argv[++i];
    } else if (SWITCHES.has(key)) {
      flags[key] = true;
    } else {
      throw new UsageError(`unknown option ${arg}`);
    }
  }
  return { files, flags };
}

function isFolder(path: string): boolean {
  return !!statSync(path, { throwIfNoEntry: false })?.isDirectory();
}

/** An xsd:dateTime as Word writes it: date, time, and Z or an offset. */
const ISO_DATE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$/;

async function main(): Promise<void> {
  const { files, flags } = parseArgs(process.argv.slice(2));
  if (flags.help) {
    console.log(USAGE);
    return;
  }
  const [original, revised, ...extra] = files;
  if (!original || !revised) throw new UsageError("give the original and the revised document");
  if (extra.length > 0) throw new UsageError(`unexpected argument ${extra[0]}`);
  if (typeof flags.out !== "string") throw new UsageError("-o RESULT.docx is required");
  if (typeof flags.author !== "string" || flags.author.trim() === "") throw new UsageError("--author NAME is required");
  if (typeof flags.date === "string" && !ISO_DATE.test(flags.date)) {
    throw new UsageError(`--date ${flags.date} is not a date and time such as 2026-10-14T09:30:00Z`);
  }

  // The result is written beside its final name and renamed into place, so a
  // failed or interrupted run leaves no half-written file, and the result may
  // replace an input. The working files are removed on Ctrl-C too.
  const out = flags.out;
  if (!isFolder(dirname(resolve(out)))) throw new Error(`${out}: the folder for the result does not exist`);
  const tempDir = typeof flags["temp-dir"] === "string" ? flags["temp-dir"] : tmpdir();
  if (!isFolder(tempDir)) throw new Error(`${tempDir}: the --temp-dir folder does not exist`);
  const partial = `${out}.${process.pid}.partial`;
  const pdf = typeof flags.pdf === "string" ? flags.pdf : undefined;
  if (pdf && !isFolder(dirname(resolve(pdf)))) throw new Error(`${pdf}: the folder for the PDF does not exist`);
  if (flags["pdf-comments"] && !pdf) throw new UsageError("--pdf-comments needs --pdf FILE");
  if (flags["hide-authors"] && !pdf) throw new UsageError("--hide-authors needs --pdf FILE");
  if (pdf && [original, revised, out].some((path) => resolve(path) === resolve(pdf))) throw new UsageError("--pdf FILE must not be an input or the result");
  const pdfPartial = pdf ? `${pdf}.${process.pid}.partial` : undefined;
  const work = mkdtempSync(join(tempDir, "docx-compare-"));
  const cleanUp = (): void => {
    rmSync(partial, { force: true });
    if (pdfPartial) rmSync(pdfPartial, { force: true });
    rmSync(work, { recursive: true, force: true });
  };
  const stop = (code: number) => () => {
    cleanUp();
    process.exit(code);
  };
  process.once("SIGINT", stop(130));
  process.once("SIGTERM", stop(143));
  try {
    await compareFiles(original, revised, partial, {
      author: flags.author,
      date: typeof flags.date === "string" ? flags.date : undefined,
      formatting: !flags["no-formatting"],
      whitespace: !flags["no-whitespace"],
      tempDir: work,
      pdf: pdfPartial ? { path: pdfPartial, comments: !!flags["pdf-comments"], authors: !flags["hide-authors"] } : undefined,
      onNote: flags.quiet ? undefined : (note) => console.error(`note: ${note.code}: ${note.detail}`),
    });
    renameSync(partial, out);
    if (pdf && pdfPartial) renameSync(pdfPartial, pdf);
  } finally {
    cleanUp();
  }
}

main().catch((error: unknown) => {
  if (error instanceof UsageError) {
    console.error(`docx-compare: ${error.message}\nRun docx-compare --help for the options.`);
    process.exitCode = 2;
    return;
  }
  const code = error instanceof CompareError ? `${error.code}: ` : "";
  console.error(`docx-compare: ${code}${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
});
