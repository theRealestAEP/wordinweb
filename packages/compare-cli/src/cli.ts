#!/usr/bin/env node
import { compareFiles } from "./compare.js";

const USAGE = `Usage: docx-compare ORIGINAL.docx REVISED.docx -o RESULT.docx --author NAME [options]

Write RESULT.docx: REVISED.docx with every difference from ORIGINAL.docx as a
tracked change by NAME. Accept all of them to get the revised document;
reject all of them to get the original.

Options:
  -o, --out FILE      the result (required)
  --author NAME       who the tracked changes are attributed to (required)
  --date ISO          the date on every tracked change (default: now)
  --no-formatting     do not report formatting-only differences
  --no-whitespace     do not report whitespace-only differences
  --temp-dir DIR      where the working files go (default: the system's)
  --quiet             do not print notes about differences that could not be expressed exactly
  -h, --help          show this help`;

function parseArgs(argv: string[]): { files: string[]; flags: Record<string, string | boolean> } {
  const files: string[] = [];
  const flags: Record<string, string | boolean> = {};
  const valued = new Set(["out", "author", "date", "temp-dir"]);
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const key = arg === "-o" ? "out" : arg === "-h" ? "help" : arg.startsWith("--") ? arg.slice(2) : null;
    if (key === null) files.push(arg);
    else if (valued.has(key)) {
      if (i + 1 >= argv.length) throw new Error(`${arg} needs a value`);
      flags[key] = argv[++i];
    } else flags[key] = true;
  }
  return { files, flags };
}

async function main(): Promise<void> {
  const { files, flags: args } = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(USAGE);
    return;
  }
  const [original, revised, ...extra] = files;
  if (!original || !revised || extra.length > 0 || typeof args.out !== "string" || typeof args.author !== "string") {
    console.error(USAGE);
    process.exitCode = 2;
    return;
  }
  await compareFiles(original, revised, args.out, {
    author: args.author,
    date: typeof args.date === "string" ? args.date : undefined,
    formatting: !args["no-formatting"],
    whitespace: !args["no-whitespace"],
    tempDir: typeof args["temp-dir"] === "string" ? args["temp-dir"] : undefined,
    onNote: args.quiet ? undefined : (note) => console.error(`note: ${note.code}: ${note.detail}`),
  });
}

main().catch((error: unknown) => {
  console.error(`docx-compare: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
});
