import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { XmlElement, serializeXml } from "../../core/src/xml.js";
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

  it("a table over a megabyte: changed, removed and added", async () => {
    // Over TABLE_CHUNK_BYTES, so a removed or added table goes through emitDeleted or emitInserted in chunks.
    const big = (name: string, edit = ""): string => table(Array.from({ length: 7500 }, (_, i) => [`${name}${i}a`, i === 4000 ? `${name}${i}b${edit}` : `${name}${i}b`, `${name}${i}c`]));
    await check(
      para("One.") + big("alpha") + para("Two.") + big("beta") + para("Three."),
      para("One.") + para("Two.") + big("beta", " edited") + para("Three.") + big("gamma"),
    );
  }, 60000);

  it("tracked changes in a table are accepted, and a table they empty joins the paragraphs around it", async () => {
    const author = `w:author="x"`;
    const deletedMark = `<w:p><w:pPr><w:rPr><w:del w:id="1" ${author}/></w:rPr></w:pPr><w:del w:id="2" ${author}><w:r><w:delText>Removed.</w:delText></w:r></w:del></w:p>`;
    const deletedRow = `<w:tr><w:trPr><w:del w:id="3" ${author}/></w:trPr><w:tc><w:tcPr/><w:p><w:del w:id="4" ${author}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p></w:tc></w:tr>`;
    const insertedRow = `<w:tr><w:trPr><w:ins w:id="5" ${author}/></w:trPr><w:tc><w:tcPr/><w:p><w:ins w:id="6" ${author}><w:r><w:t>new</w:t></w:r></w:ins></w:p></w:tc></w:tr>`;
    const tblPr = (w: string, change = ""): string => `<w:tblPr><w:tblW w:w="${w}" w:type="dxa"/>${change}</w:tblPr>`;
    const grid = `<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>`;
    const row = (text: string): string => `<w:tr><w:tc><w:tcPr/>${para(text)}</w:tc></w:tr>`;
    // The deleted paragraph mark joins "After." once the table between them is gone.
    const tracked =
      para("Keep.") + deletedMark + `<w:tbl>${tblPr("3000")}${grid}${deletedRow}</w:tbl>` + para("After.") +
      `<w:tbl>${tblPr("5000", `<w:tblPrChange w:id="7" ${author}>${tblPr("3000")}</w:tblPrChange>`)}${grid}${row("a")}${insertedRow}</w:tbl>`;
    const accepted = para("Keep.") + para("After.") + `<w:tbl>${tblPr("5000")}${grid}${row("a")}${row("new")}</w:tbl>`;
    const revised = para("Keep.") + para("After, edited.") + `<w:tbl>${tblPr("5000")}${grid}${row("a")}${row("new, edited")}</w:tbl>`;
    const n = pairs++;
    const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`)];
    writeFileSync(paths[0], makeDocx({ "word/document.xml": wrapDocument(tracked) }));
    writeFileSync(paths[1], makeDocx({ "word/document.xml": wrapDocument(revised) }));
    const notes: string[] = [];
    await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE, onNote: (note) => notes.push(note.code) });
    expect(notes).toContain("existing-revisions-accepted");
    const result = new Uint8Array(readFileSync(paths[2]));
    const rejected = DocxDocument.load(result);
    rejectAllRevisions(rejected);
    expect(canonicalBody(rejected)).toBe(canonicalBody(DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(accepted) }))));
    const acceptedDoc = DocxDocument.load(result);
    acceptAllRevisions(acceptedDoc);
    expect(canonicalBody(acceptedDoc)).toBe(canonicalBody(DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(revised) }))));
  });

  it("a note reference in a removed row gets its own note", async () => {
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const parts = {
      "word/_rels/document.xml.rels":
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`,
      "word/footnotes.xml":
        `<w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="1"><w:p><w:r><w:t xml:space="preserve">A source.</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    };
    const cell = (text: string): string => `<w:tc><w:tcPr/><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:tc>`;
    // The row's cell count changes, so it is struck and reinserted whole: two
    // references to one note, one of them struck, which gets a struck copy of
    // the note. A reject restores the row pointing at the copy, as
    // compareDocuments has it, so the result is checked against that alone.
    const grid = `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`;
    const original = makeDocx({ ...parts, "word/document.xml": wrapDocument(`<w:tbl><w:tblPr/>${grid}<w:tr>${cell("Cited")}<w:tc><w:tcPr/>${para("More")}</w:tc></w:tr></w:tbl>`) });
    const revised = makeDocx({ ...parts, "word/document.xml": wrapDocument(`<w:tbl><w:tblPr/>${grid}<w:tr>${cell("Cited")}</w:tr></w:tbl>`) });
    const n = pairs++;
    const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`)];
    writeFileSync(paths[0], original);
    writeFileSync(paths[1], revised);
    await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE });
    const result = DocxDocument.load(new Uint8Array(readFileSync(paths[2])));
    const expected = compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), { author: AUTHOR, date: DATE });
    expect(withoutIds(canonicalBody(result))).toBe(withoutIds(canonicalBody(expected)));
    const notes = (doc: DocxDocument) => withoutIds(serializeXml(doc.footnotesTree() as XmlElement));
    expect(notes(result)).toBe(notes(expected));
    expect(result.footnotesTree()?.children.filter((c) => c.attrs["w:id"] === "2")).toHaveLength(1);
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
