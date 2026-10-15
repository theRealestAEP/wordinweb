import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { XmlElement, child, localName, parseXml, serializeXml } from "../../core/src/xml.js";
import { compareDocuments } from "../../core/src/edit/compare/index.js";
import { acceptAllRevisions, collectRevisions, rejectAllRevisions } from "../../core/src/edit/suggest.js";
import { makeDocx, wrapDocument } from "../../core/test/helpers.js";
import { AUTHOR, DATE, canonicalBody, para } from "../../core/test/compare-helpers.js";
import { compareFiles } from "../src/compare.js";
import { CompareError } from "../src/errors.js";
import { flatten } from "../src/printable.js";

const dir = mkdtempSync(join(tmpdir(), "compare-files-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

let pairs = 0;

/** compareFiles on two bodies: the result must review back to either input, and say what compareDocuments says. */
async function check(originalBody: string, revisedBody: string, parts: Record<string, string> = {}): Promise<DocxDocument> {
  const n = pairs++;
  const original = makeDocx({ ...parts, "word/document.xml": wrapDocument(originalBody) });
  const revised = makeDocx({ ...parts, "word/document.xml": wrapDocument(revisedBody) });
  const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`), join(dir, `${n}-print.docx`)];
  writeFileSync(paths[0], original);
  writeFileSync(paths[1], revised);
  const { revisions } = await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE, printable: { path: paths[3], beforeSectPr: SUMMARY } });
  const result = new Uint8Array(readFileSync(paths[2]));

  // The counts are collectRevisions' on the result, kind by kind.
  const tally: Record<string, number> = Object.fromEntries(Object.keys(revisions).map((k) => [k, 0]));
  for (const ref of collectRevisions(DocxDocument.load(result))) tally[ref.kind]++;
  expect(revisions).toEqual(tally);
  // The printable copy is the result flattened whole, the summary before the final sectPr.
  const reference = DocxDocument.load(result);
  const body = child(reference.docRoot, "body") as XmlElement;
  flatten(body);
  const sectPr = body.children.findIndex((c) => localName(c.name) === "sectPr");
  body.children.splice(sectPr === -1 ? body.children.length : sectPr, 0, ...parseXml(`<root>${SUMMARY}</root>`).children);
  expect(canonicalBody(DocxDocument.load(new Uint8Array(readFileSync(paths[3]))))).toBe(canonicalBody(reference));

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

const SUMMARY = `<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>Summary</w:t></w:r></w:p>`;

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

  it("footnotes edited, added, removed and tracked are compared as compareDocuments compares them", async () => {
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const rels =
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="${REL}/comments" Target="comments.xml"/></Relationships>`;
    const note = (id: number, body: string): string => `<w:footnote w:id="${id}"><w:p><w:r><w:t xml:space="preserve">${body}</w:t></w:r></w:p></w:footnote>`;
    const notesPart = (notes: string): string =>
      `<w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>${notes}</w:footnotes>`;
    const cited = (text: string, id: number): string =>
      `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`;
    const comments = `<w:comments ${W_NS}><w:comment w:id="0" w:author="x"><w:p><w:r><w:t>A comment.</w:t></w:r></w:p></w:comment></w:comments>`;
    const tracked = `<w:footnote w:id="4"><w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:ins w:id="9" w:author="x"><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:footnote>`;
    const original = makeDocx({
      "word/_rels/document.xml.rels": rels,
      "word/comments.xml": comments,
      "word/footnotes.xml": notesPart(note(1, "Same note.") + note(2, "A note that is edited here.") + note(3, "A note that goes.") + tracked),
      "word/document.xml": wrapDocument(cited("One", 1) + cited("Two", 2) + cited("Three", 3) + cited("Four", 4)),
    });
    const revised = makeDocx({
      "word/_rels/document.xml.rels": rels,
      "word/comments.xml": comments,
      "word/footnotes.xml": notesPart(note(1, "Same note.") + note(2, "A note that was edited here.") + note(5, "A new note.") + note(4, "Kept added")),
      "word/document.xml": wrapDocument(cited("One", 1) + cited("Two", 2) + cited("Five", 5) + cited("Four", 4)),
    });
    const n = pairs++;
    const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`)];
    writeFileSync(paths[0], original);
    writeFileSync(paths[1], revised);
    await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE });
    const bytes = new Uint8Array(readFileSync(paths[2]));
    const result = DocxDocument.load(bytes);
    const expected = compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), { author: AUTHOR, date: DATE });
    expect(withoutIds(canonicalBody(result))).toBe(withoutIds(canonicalBody(expected)));
    const notes = (doc: DocxDocument) => withoutIds(serializeXml(doc.footnotesTree() as XmlElement));
    expect(notes(result)).toBe(notes(expected));
    // The comments are never read: the result carries the revised part as it was.
    expect(strFromU8(unzipSync(bytes)["word/comments.xml"])).toBe(comments);
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

describe("compareFiles errors", () => {
  /** compareFiles' error, which must be a CompareError. */
  const errorOf = async (original: string, revised: string, out: string): Promise<CompareError> => {
    try {
      await compareFiles(original, revised, out, { author: AUTHOR, date: DATE });
    } catch (error) {
      expect(error).toBeInstanceOf(CompareError);
      return error as CompareError;
    }
    throw new Error("compareFiles succeeded");
  };

  const good = join(dir, "good.docx");
  writeFileSync(good, makeDocx({ "word/document.xml": wrapDocument(para("Text.")) }));
  const options = { author: AUTHOR, date: DATE };
  const fails = (original: string, revised: string, out: string) => expect(compareFiles(original, revised, out, options)).rejects;

  it("names a missing input, a folder, and a file that is not a .docx", async () => {
    await fails(join(dir, "missing.docx"), good, join(dir, "x.docx")).toThrow(/missing\.docx: the original file does not exist/);
    await fails(good, dir, join(dir, "x.docx")).toThrow(/the revised file is not a file/);
    writeFileSync(join(dir, "text.docx"), "not a zip");
    await fails(join(dir, "text.docx"), good, join(dir, "x.docx")).toThrow(/text\.docx: not a \.docx file$/);
    writeFileSync(join(dir, "ole.docx"), Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]));
    await fails(join(dir, "ole.docx"), good, join(dir, "x.docx")).toThrow(/a \.doc file or a password-protected document/);
  });

  it("names the input and part whose XML cannot be read", async () => {
    const broken = join(dir, "broken.docx");
    writeFileSync(broken, makeDocx({ "word/document.xml": wrapDocument(para("Text.")).replace("</w:body>", "<!-- never closed") }));
    await fails(good, broken, join(dir, "x.docx")).toThrow(/broken\.docx: word\/document\.xml: unterminated comment/);
  });

  it("gives each failure a code and the input at fault", async () => {
    const xlsx = join(dir, "book.docx");
    writeFileSync(xlsx, makeDocx({ "word/document.xml": `<w:workbook ${'xmlns:w="w"'}/>` }));
    const broken = join(dir, "broken2.docx");
    writeFileSync(broken, makeDocx({ "word/document.xml": wrapDocument(para("Text.")).replace("</w:body>", "<!-- never closed") }));
    writeFileSync(join(dir, "text2.docx"), "not a zip");
    const cases: [string, string, string, string, string | undefined][] = [
      [join(dir, "missing.docx"), good, join(dir, "x.docx"), "missing", "original"],
      [good, dir, join(dir, "x.docx"), "missing", "revised"],
      [join(dir, "text2.docx"), good, join(dir, "x.docx"), "not_docx", "original"],
      [good, xlsx, join(dir, "x.docx"), "not_docx", "revised"],
      [good, broken, join(dir, "x.docx"), "damaged", "revised"],
      [good, good, good, "result_is_input", undefined],
      [good, good, join(dir, "no-such-folder", "x.docx"), "result_folder_missing", undefined],
    ];
    for (const [original, revised, out, code, file] of cases) {
      const error = await errorOf(original, revised, out);
      expect([error.code, error.file]).toEqual([code, file]);
    }
  });

  it("refuses to overwrite an input, and to write into a missing folder", async () => {
    await fails(good, good, good).toThrow(/must not overwrite an input/);
    await fails(good, good, join(dir, "no-such-folder", "x.docx")).toThrow(/folder does not exist/);
  });
});
