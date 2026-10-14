import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { compareDocuments } from "../../core/src/edit/compare/index.js";
import { acceptAllRevisions, collectRevisions, rejectAllRevisions } from "../../core/src/edit/suggest.js";
import { makeDocx, wrapDocument } from "../../core/test/helpers.js";
import { AUTHOR, DATE, canonicalBody, para } from "../../core/test/compare-helpers.js";
import { compareFiles } from "../src/compare.js";

const dir = mkdtempSync(join(tmpdir(), "compare-files-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

let pairs = 0;

/** compareFiles on two bodies: the result must review back to either input, and say what compareDocuments says. */
async function check(originalBody: string, revisedBody: string, parts: Record<string, string> = {}): Promise<DocxDocument> {
  const n = pairs++;
  const original = makeDocx({ ...parts, "word/document.xml": wrapDocument(originalBody) });
  const revised = makeDocx({ ...parts, "word/document.xml": wrapDocument(revisedBody) });
  const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`)];
  writeFileSync(paths[0], original);
  writeFileSync(paths[1], revised);
  await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE });
  const result = new Uint8Array(readFileSync(paths[2]));

  const accepted = DocxDocument.load(result);
  acceptAllRevisions(accepted);
  expect(canonicalBody(accepted)).toBe(canonicalBody(DocxDocument.load(revised)));
  expect(collectRevisions(accepted)).toHaveLength(0);
  const rejected = DocxDocument.load(result);
  rejectAllRevisions(rejected);
  expect(canonicalBody(rejected)).toBe(canonicalBody(DocxDocument.load(original)));

  const expected = compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), { author: AUTHOR, date: DATE });
  expect(withoutIds(canonicalBody(DocxDocument.load(result)))).toBe(withoutIds(canonicalBody(expected)));
  return DocxDocument.load(result);
}

/** Revision ids are unique either way, but taken in a different order. */
function withoutIds(xml: string): string {
  return xml.replace(/ w:id="\d+"/g, "");
}

const table = (rows: string[][]): string =>
  `<w:tbl><w:tblPr/><w:tblGrid>${rows[0].map(() => `<w:gridCol w:w="2000"/>`).join("")}</w:tblGrid>${rows
    .map((cells) => `<w:tr>${cells.map((c) => `<w:tc><w:tcPr/>${para(c)}</w:tc>`).join("")}</w:tr>`)
    .join("")}</w:tbl>`;

describe("compareFiles", () => {
  it("a word changed", async () => {
    await check(para("The quick brown fox jumps.") + para("Second."), para("The quick brown cat jumps.") + para("Second."));
  });

  it("identical documents give no revisions", async () => {
    const body = para("One.") + para("Two.") + table([["a", "b"]]);
    const result = await check(body, body);
    expect(collectRevisions(result)).toHaveLength(0);
  });

  it("paragraphs added, removed and edited", async () => {
    const o = ["Alpha one.", "Beta two.", "Gamma three.", "Delta four.", "Epsilon five."].map((t) => para(t)).join("");
    const r = ["Alpha one.", "Gamma three, edited.", "New paragraph.", "Delta four.", "Epsilon five.", "Zeta six."].map((t) => para(t)).join("");
    await check(o, r);
  });

  it("entirely different documents", async () => {
    await check(para("First original.") + para("Second original."), para("Something else.") + para("And more."));
  });

  it("a paragraph split in two and two merged", async () => {
    await check(
      para("Keep this.") + para("One long sentence. And another.") + para("Merge me") + para(" with this.") + para("End."),
      para("Keep this.") + para("One long sentence.") + para(" And another.") + para("Merge me with this.") + para("End."),
    );
  });

  it("a table row changed and a row added", async () => {
    await check(
      para("Before.") + table([["a", "b"], ["c", "d"]]) + para("After."),
      para("Before.") + table([["a", "b"], ["c", "e"], ["f", "g"]]) + para("After."),
    );
  });

  it("a heading's style changed", async () => {
    await check(
      para("Title", { pPr: `<w:pStyle w:val="Heading1"/>` }) + para("Body."),
      para("Title", { pPr: `<w:pStyle w:val="Heading2"/>` }) + para("Body."),
    );
  });

  it("section breaks", async () => {
    const sect = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    await check(
      para("Section one.", { pPr: sect }) + para("Section two."),
      para("Section one, edited.", { pPr: sect }) + para("Inserted.") + para("Section two."),
    );
  });

  it("an input with tracked changes is compared as accepted", async () => {
    const tracked = `<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:ins w:id="7" w:author="x"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="8" w:author="x"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`;
    const n = pairs++;
    const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`)];
    writeFileSync(paths[0], makeDocx({ "word/document.xml": wrapDocument(tracked + para("Other.")) }));
    writeFileSync(paths[1], makeDocx({ "word/document.xml": wrapDocument(para("Kept added") + para("Other, edited.")) }));
    const notes: string[] = [];
    await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE, onNote: (note) => notes.push(note.code) });
    expect(notes).toContain("existing-revisions-accepted");
    const rejected = DocxDocument.load(new Uint8Array(readFileSync(paths[2])));
    rejectAllRevisions(rejected);
    const expected = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(para("Kept added") + para("Other.")) }));
    expect(canonicalBody(rejected)).toBe(canonicalBody(expected));
  });
});
