import * as CFB from "cfb";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { docToDocx } from "../../core/src/doc/hooks.js";
import { compareDocuments } from "../../core/src/edit/compare/index.js";
import { collectRevisions } from "../../core/src/edit/suggest.js";
import { AUTHOR, DATE } from "../../core/test/compare-helpers.js";
import { makeDocx, wrapDocument } from "../../core/test/helpers.js";
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

  it("compares a Word 95 .doc with itself, and with its own .docx", async () => {
    const doc = join(fixtures, "word95-025814.doc");
    const same = await compareFiles(doc, doc, join(dir, "word95-same.docx"), options);
    expect(Object.values(same.revisions).reduce((a, b) => a + b, 0)).toBe(0);
    const docx = join(dir, "word95.docx");
    writeFileSync(docx, docToDocx(new Uint8Array(readFileSync(doc))));
    const against = await compareFiles(doc, docx, join(dir, "word95-against.docx"), options);
    expect(Object.values(against.revisions).reduce((a, b) => a + b, 0)).toBe(0);
    const result = DocxDocument.load(new Uint8Array(readFileSync(join(dir, "word95-against.docx"))));
    expect(result.sections[0].blocks.some((block) => block.type === "table")).toBe(true);
  });

  it("refuses the formats it does not read, naming each", async () => {
    const good = join(fixtures, "parity-columns--v03.docx");
    writeFileSync(join(dir, "text.doc"), "{\\rtf1\\ansi Hello}");
    writeFileSync(join(dir, "print.doc"), "%PDF-1.4 ...");
    const cfb = CFB.utils.cfb_new();
    CFB.utils.cfb_add(cfb, "Workbook", new Uint8Array(64));
    writeFileSync(join(dir, "empty.doc"), new Uint8Array(CFB.write(cfb, { type: "array" }) as number[]));
    const cases: [string, RegExp][] = [
      [join(dir, "text.doc"), /text\.doc: an RTF file/],
      [join(dir, "print.doc"), /print\.doc: a PDF file/],
      [join(dir, "empty.doc"), /empty\.doc: an Excel 97-2003 workbook/],
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

describe("compareFiles with the oldest Word formats", () => {
  // GovDocs1 documents (public domain): Word for the Macintosh 5.0 and Word for Windows 2.0.
  const oldword = join(fixtures, "..", "oldword");

  it("compares each with itself: no revisions", async () => {
    for (const name of ["macword5-govdocs-000790.doc", "winword2-govdocs-474644.doc"]) {
      const path = join(oldword, name);
      const { revisions } = await compareFiles(path, path, join(dir, "old-same.docx"), options);
      expect(Object.values(revisions).reduce((a, b) => a + b, 0)).toBe(0);
    }
  });

  it("compares one with a .docx", async () => {
    const mac = join(oldword, "macword5-govdocs-000271.doc");
    const docx = join(fixtures, "parity-columns--v03.docx");
    const { revisions } = await compareFiles(mac, docx, join(dir, "old-docx.docx"), options);
    expect(revisions.insertion).toBeGreaterThan(0);
    expect(revisions.deletion).toBeGreaterThan(0);
  });
});

describe("comments numbered otherwise", () => {
  // The same document with every comment id raised by 100, as a .doc and the
  // .docx Word makes of it differ, and with a page break Word last rendered.
  const renumbered = (): Uint8Array => {
    const files = unzipSync(new Uint8Array(readFileSync(join(fixtures, "parity-columns--v03.docx"))));
    const raise = (name: string): void => {
      const xml = strFromU8(files[name]).replace(/(<w:(?:comment|commentRangeStart|commentRangeEnd|commentReference)\b[^>]*?\bw:id=")(\d+)"/g, (_, head: string, id: string) => `${head}${Number(id) + 100}"`);
      files[name] = strToU8(xml);
    };
    raise("word/document.xml");
    raise("word/comments.xml");
    files["word/document.xml"] = strToU8(strFromU8(files["word/document.xml"]).replace("<w:r>", "<w:r><w:lastRenderedPageBreak/></w:r><w:r>"));
    return zipSync(files);
  };

  it("gives no change for a comment that is the same in both documents", async () => {
    const original = join(fixtures, "parity-columns--v03.docx");
    const revised = join(dir, "renumbered.docx");
    writeFileSync(revised, renumbered());
    const { revisions } = await compareFiles(original, revised, join(dir, "renumbered-result.docx"), options);
    expect(Object.values(revisions).reduce((a, b) => a + b, 0)).toBe(0);
    const result = compareDocuments(DocxDocument.load(new Uint8Array(readFileSync(original))), DocxDocument.load(renumbered()), options);
    expect(collectRevisions(result).length).toBe(0);
  });
});

describe("a field written whole or in runs", () => {
  const body = (field: string): Uint8Array =>
    makeDocx({ "word/document.xml": wrapDocument(`<w:p><w:r><w:t xml:space="preserve">Author: </w:t></w:r>${field}<w:r><w:t xml:space="preserve"> signed.</w:t></w:r></w:p>`) });
  const result = (text: string): string => `<w:r><w:rPr><w:i/></w:rPr><w:t>${text}</w:t></w:r>`;
  const whole = (text: string): Uint8Array => body(`<w:fldSimple w:instr=" AUTHOR \\* MERGEFORMAT ">${result(text)}</w:fldSimple>`);
  const inRuns = (text: string): Uint8Array =>
    body(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> AUTHOR \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result(text)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`);
  /** The revisions compareDocuments finds, after checking compareFiles agrees on whether there are any. */
  const count = async (a: Uint8Array, b: Uint8Array): Promise<number> => {
    writeFileSync(join(dir, "field-a.docx"), a);
    writeFileSync(join(dir, "field-b.docx"), b);
    const { revisions } = await compareFiles(join(dir, "field-a.docx"), join(dir, "field-b.docx"), join(dir, "field-result.docx"), options);
    const inMemory = collectRevisions(compareDocuments(DocxDocument.load(a), DocxDocument.load(b), options)).length;
    expect(Object.values(revisions).reduce((x, y) => x + y, 0) > 0).toBe(inMemory > 0);
    return inMemory;
  };

  it("is the same field when its instruction and result are", async () => {
    expect(await count(whole("Jordan Rivera"), inRuns("Jordan Rivera"))).toBe(0);
    expect(await count(inRuns("Jordan Rivera"), whole("Jordan Rivera"))).toBe(0);
    // A result that changed is still a change.
    expect(await count(whole("Jordan Rivera"), inRuns("Sam Lee"))).toBeGreaterThan(0);
  });
});
