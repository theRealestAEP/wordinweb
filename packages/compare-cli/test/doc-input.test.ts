import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { collectRevisions } from "../../core/src/edit/suggest.js";
import { AUTHOR, DATE } from "../../core/test/compare-helpers.js";
import { compareFiles } from "../src/compare.js";
import { CompareError } from "../src/errors.js";

// The Word-made .doc files core's doc-reader tests use, with their .docx originals.
const fixtures = join(fileURLToPath(new URL(".", import.meta.url)), "..", "..", "core", "test", "fixtures", "doc");
const dir = mkdtempSync(join(tmpdir(), "compare-doc-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const options = { author: AUTHOR, date: DATE };

async function errorOf(original: string, revised: string): Promise<CompareError> {
  try {
    await compareFiles(original, revised, join(dir, "x.docx"), options);
  } catch (error) {
    return error as CompareError;
  }
  throw new Error("compared");
}

describe("compareFiles with a Word 97-2003 .doc", () => {
  it("compares a .doc with itself as a .docx: no revisions", async () => {
    const doc = join(fixtures, "parity-columns--v03.doc");
    const out = join(dir, "same.docx");
    const { revisions } = await compareFiles(doc, doc, out, options);
    expect(Object.values(revisions).reduce((a, b) => a + b, 0)).toBe(0);
    const result = DocxDocument.load(new Uint8Array(readFileSync(out)));
    expect(collectRevisions(result).length).toBe(0);
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
  });

  it("compares a .doc with a .docx, either way round", async () => {
    const doc = join(fixtures, "parity-tables--v00.doc");
    const docx = join(fixtures, "parity-tables--v00.docx");
    const a = await compareFiles(doc, docx, join(dir, "a.docx"), options);
    const b = await compareFiles(docx, doc, join(dir, "b.docx"), options);
    // The .doc holds the original's tracked deletions as plain text until the
    // review marks are read, so the two differ by exactly those; each way
    // round sees the same count.
    expect(a.revisions.insertion + a.revisions.deletion).toBe(b.revisions.insertion + b.revisions.deletion);
    expect(DocxDocument.load(new Uint8Array(readFileSync(join(dir, "a.docx")))).sections[0].blocks.some((block) => block.type === "table")).toBe(true);
  });

  it("refuses the formats it does not read, naming each", async () => {
    const good = join(fixtures, "parity-columns--v03.docx");
    writeFileSync(join(dir, "text.doc"), "{\\rtf1\\ansi Hello}");
    writeFileSync(join(dir, "print.doc"), "%PDF-1.4 ...");
    const empty = new Uint8Array(512);
    empty.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    empty[30] = 9;
    empty[32] = 6;
    empty.set([0xfe, 0xff, 0xff, 0xff], 48);
    empty.set([0xfe, 0xff, 0xff, 0xff], 68);
    writeFileSync(join(dir, "empty.doc"), empty);
    const cases: [string, RegExp][] = [
      [join(dir, "text.doc"), /text\.doc: an RTF file/],
      [join(dir, "print.doc"), /print\.doc: a PDF file/],
      [join(dir, "empty.doc"), /empty\.doc: an OLE compound file with no Word document/],
    ];
    for (const [path, message] of cases) {
      const error = await errorOf(path, good);
      expect([error.code, error.file]).toEqual(["unsupported_format", "original"]);
      expect(error.message).toMatch(message);
    }
    const revised = await errorOf(good, join(dir, "text.doc"));
    expect([revised.code, revised.file]).toEqual(["unsupported_format", "revised"]);
  });
});
