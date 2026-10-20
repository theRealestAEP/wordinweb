import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { crc32, deflateSync, inflateSync } from "node:zlib";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { Face } from "../../core/src/pdf/sfnt.js";
import { XmlElement, child, localName, parseXml, serializeXml } from "../../core/src/xml.js";
import { CompareNote, compareDocuments } from "../../core/src/edit/compare/index.js";
import { acceptAllRevisions, collectRevisions, rejectAllRevisions } from "../../core/src/edit/suggest.js";
import { makeDocx, makeDocxWithMedia, wrapDocument, W_NS } from "../../core/test/helpers.js";
import { AUTHOR, DATE, canonicalBody, para } from "../../core/test/compare-helpers.js";
import { compareFiles } from "../src/compare.js";
import { CompareError } from "../src/errors.js";
import { flatten } from "../src/printable.js";
import { MAX_HELD_BYTES, ZipWriter } from "../src/zip.js";

const dir = mkdtempSync(join(tmpdir(), "compare-files-test-"));
let hasPdftotext: boolean;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  hasPdftotext = true;
} catch {
  hasPdftotext = false;
}
afterAll(() => rmSync(dir, { recursive: true, force: true }));

let pairs = 0;

/** A .docx from a body, or as given. */
function docx(input: string | Uint8Array): Uint8Array {
  return typeof input === "string" ? makeDocx({ "word/document.xml": wrapDocument(input) }) : input;
}

/** The parts beside the body that can carry revisions (headers, footers, notes), serialized, in package order. */
function otherParts(doc: DocxDocument): string {
  return doc.revisionRoots().slice(1).map((root) => serializeXml(root)).join("\n");
}

/**
 * compareFiles on two documents: the result's body must review back to
 * either input, and every part must say what compareDocuments says. The
 * printable copy must be the result with every part flattened. Returns the
 * result and the printable copy.
 */
async function check(originalInput: string | Uint8Array, revisedInput: string | Uint8Array, onNote?: (note: CompareNote) => void): Promise<{ result: DocxDocument; printable: DocxDocument }> {
  const n = pairs++;
  const original = docx(originalInput);
  const revised = docx(revisedInput);
  const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`), join(dir, `${n}-print.docx`)];
  writeFileSync(paths[0], original);
  writeFileSync(paths[1], revised);
  const { revisions } = await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE, onNote, printable: { path: paths[3], beforeSectPr: SUMMARY } });
  const result = new Uint8Array(readFileSync(paths[2]));

  // The counts are collectRevisions' on the result, kind by kind.
  const tally: Record<string, number> = Object.fromEntries(Object.keys(revisions).map((k) => [k, 0]));
  for (const ref of collectRevisions(DocxDocument.load(result))) tally[ref.kind]++;
  expect(revisions).toEqual(tally);
  // The printable copy is the result flattened whole, the summary before the final sectPr.
  const reference = DocxDocument.load(result);
  for (const root of reference.revisionRoots()) flatten(root);
  const body = child(reference.docRoot, "body") as XmlElement;
  const sectPr = body.children.findIndex((c) => localName(c.name) === "sectPr");
  body.children.splice(sectPr === -1 ? body.children.length : sectPr, 0, ...parseXml(`<root>${SUMMARY}</root>`).children);
  const printable = DocxDocument.load(new Uint8Array(readFileSync(paths[3])));
  expect(canonicalBody(printable)).toBe(canonicalBody(reference));
  expect(otherParts(printable)).toBe(otherParts(reference));

  const accepted = DocxDocument.load(result);
  acceptAllRevisions(accepted);
  expect(reviewedBody(accepted)).toBe(reviewedBody(DocxDocument.load(revised)));
  expect(collectRevisions(accepted)).toHaveLength(0);
  const rejected = DocxDocument.load(result);
  rejectAllRevisions(rejected);
  expect(reviewedBody(rejected)).toBe(reviewedBody(DocxDocument.load(original)));

  const expected = compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), { author: AUTHOR, date: DATE });
  expect(withoutIds(canonicalBody(DocxDocument.load(result)))).toBe(withoutIds(canonicalBody(expected)));
  expect(withoutIds(otherParts(DocxDocument.load(result)))).toBe(withoutIds(otherParts(expected)));
  return { result: DocxDocument.load(result), printable };
}

const SUMMARY = `<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>Summary</w:t></w:r></w:p>`;

/**
 * The body as a review gives it: without the last section's header and
 * footer references and title-page flag, which no revision carries, so a
 * review leaves them as the result has them.
 */
function reviewedBody(doc: DocxDocument): string {
  const sectPr = child(child(doc.docRoot, "body"), "sectPr");
  if (sectPr) sectPr.children = sectPr.children.filter((c) => !["headerReference", "footerReference", "titlePg"].includes(localName(c.name)));
  return canonicalBody(doc);
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
    const { result } = await check(body, body);
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

  it("gives a deleted note reference the original's note and an inserted one the revised's", async () => {
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`;
    const notesPart = (notes: Record<number, string>): string =>
      `<w:footnotes ${W_NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
      Object.entries(notes).map(([n, t]) => `<w:footnote w:id="${n}"><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:footnote>`).join("") +
      `</w:footnotes>`;
    const ref = (n: number) => `<w:r><w:footnoteReference w:id="${n}"/></w:r>`;
    const run = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
    const withNotes = (paragraph: string, notes: Record<number, string>) =>
      makeDocx({ "word/_rels/document.xml.rels": rels, "word/footnotes.xml": notesPart(notes), "word/document.xml": wrapDocument(`<w:p>${paragraph}</w:p>`) });
    const original = withNotes(run("Alpha") + ref(1) + run(" beta") + ref(2) + run(" gamma."), { 1: "Note one, deleted", 2: "Note two, kept" });
    const revised = withNotes(run("Alpha beta") + ref(1) + run(" gamma") + ref(2) + run("."), { 1: "Note two, kept", 2: "Note three, inserted" });
    const n = pairs++;
    const paths = [join(dir, `${n}-o.docx`), join(dir, `${n}-r.docx`), join(dir, `${n}-out.docx`)];
    writeFileSync(paths[0], original);
    writeFileSync(paths[1], revised);
    const notes: CompareNote[] = [];
    const { revisions } = await compareFiles(paths[0], paths[1], paths[2], { author: AUTHOR, date: DATE, onNote: (note) => notes.push(note) });
    expect(revisions).toMatchObject({ insertion: 2, deletion: 2 });
    expect(notes.filter((note) => note.code === "untracked-block")).toEqual([]);
    const result = DocxDocument.load(new Uint8Array(readFileSync(paths[2])));
    const expected = compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), { author: AUTHOR, date: DATE });
    expect(withoutIds(canonicalBody(result))).toBe(withoutIds(canonicalBody(expected)));
    const footnotes = (doc: DocxDocument) => withoutIds(serializeXml(doc.footnotesTree() as XmlElement));
    expect(footnotes(result)).toBe(footnotes(expected));
    expect(footnotes(result)).toContain("<w:delText xml:space=\"preserve\">Note one, deleted</w:delText>");
    expect(footnotes(result)).toMatch(/<w:ins [^>]*><w:r><w:t xml:space="preserve">Note three, inserted<\/w:t>/);
  });

  describe("the parts beside the body", () => {
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    /** A one-paragraph document with the parts named here, each holding its text, and the body referencing each. */
    const withParts = (texts: Partial<Record<"header" | "footer" | "first" | "footnote" | "endnote", string>>): Uint8Array => {
      const line = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
      const notesPart = (kind: "footnote" | "endnote", text: string): string =>
        `<w:${kind}s ${W_NS}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
        `<w:${kind} w:id="1">${line(text)}</w:${kind}></w:${kind}s>`;
      const rels: string[] = [];
      const parts: Record<string, string> = {};
      const refs: string[] = [];
      let runs = "";
      if (texts.footnote !== undefined) {
        rels.push(`<Relationship Id="rId1" Type="${REL}/footnotes" Target="footnotes.xml"/>`);
        parts["word/footnotes.xml"] = notesPart("footnote", texts.footnote);
        runs += `<w:r><w:footnoteReference w:id="1"/></w:r>`;
      }
      if (texts.endnote !== undefined) {
        rels.push(`<Relationship Id="rId2" Type="${REL}/endnotes" Target="endnotes.xml"/>`);
        parts["word/endnotes.xml"] = notesPart("endnote", texts.endnote);
        runs += `<w:r><w:endnoteReference w:id="1"/></w:r>`;
      }
      for (const [name, kind, type, id] of [["header", "header", "default", 3], ["footer", "footer", "default", 4], ["first", "header", "first", 5]] as const) {
        const text = texts[name];
        if (text === undefined) continue;
        rels.push(`<Relationship Id="rId${id}" Type="${REL}/${kind}" Target="${kind}${id}.xml"/>`);
        parts[`word/${kind}${id}.xml`] = `<w:${kind === "header" ? "hdr" : "ftr"} ${W_NS}>${line(text)}${line("Line two.")}</w:${kind === "header" ? "hdr" : "ftr"}>`;
        refs.push(`<w:${kind}Reference w:type="${type}" r:id="rId${id}"/>`);
      }
      return makeDocx({
        ...parts,
        "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`,
        "word/document.xml":
          `<w:document ${W_NS} ${R_NS}><w:body><w:p><w:r><w:t xml:space="preserve">Body text.</w:t></w:r>${runs}</w:p>` +
          `<w:sectPr>${refs.join("")}${texts.first === undefined ? "" : "<w:titlePg/>"}</w:sectPr></w:body></w:document>`,
      });
    };
    /** The text of each header and footer part the body references, by reference, and of each note. */
    const texts = (doc: DocxDocument): Record<string, string> => {
      const d = DocxDocument.load(doc.save());
      const text = (xml: string): string => [...xml.matchAll(/<w:t[^>]*>([^<]*)/g)].map((m) => m[1]).join(" ");
      const out: Record<string, string> = {};
      for (const m of (d.pkg.text("word/document.xml") ?? "").matchAll(/<w:(header|footer)Reference [^>]*w:type="(\w+)"[^>]*r:id="(\w+)"/g)) {
        out[`${m[1]}:${m[2]}`] = text(d.pkg.text(d.documentRels.get(m[3])?.target ?? "") ?? "");
      }
      for (const kind of ["footnote", "endnote"] as const) {
        const note = (kind === "footnote" ? d.footnotesTree() : d.endnotesTree())?.children.find((c) => c.attrs["w:id"] === "1");
        if (note) out[kind] = text(serializeXml(note));
      }
      return out;
    };
    const reviewed = (doc: DocxDocument, review: (d: DocxDocument) => unknown): DocxDocument => {
      const copy = DocxDocument.load(doc.save());
      review(copy);
      return copy;
    };
    /** check, with the header-footer-reference notes and the deletion and insertion counts. */
    const compare = async (original: Uint8Array, revised: Uint8Array) => {
      const notes: string[] = [];
      const { result } = await check(original, revised, (note) => {
        if (note.code === "header-footer-reference") notes.push(note.detail);
      });
      const counts = collectRevisions(result).reduce((n, r) => (r.kind === "deletion" ? [n[0] + 1, n[1]] : r.kind === "insertion" ? [n[0], n[1] + 1] : n), [0, 0]);
      return { notes, counts, accepted: reviewed(result, acceptAllRevisions), rejected: reviewed(result, rejectAllRevisions) };
    };

    it("edits in a header, a footer, a footnote and an endnote show in the printable copy as plain formatting", async () => {
      const { printable } = await check(
        withParts({ header: "BEFOREHEADER here", footer: "BEFOREFOOTER here", footnote: "BEFORENOTE here", endnote: "BEFOREEND here" }),
        withParts({ header: "AFTERHEADER here", footer: "AFTERFOOTER here", footnote: "AFTERNOTE here", endnote: "AFTEREND here" }),
      );
      const printed = DocxDocument.load(printable.save());
      for (const [part, old] of [["word/header3.xml", "BEFOREHEADER"], ["word/footer4.xml", "BEFOREFOOTER"], ["word/footnotes.xml", "BEFORENOTE"], ["word/endnotes.xml", "BEFOREEND"]]) {
        const xml = printed.pkg.text(part) ?? "";
        expect(xml).not.toMatch(/<w:(?:ins|del|delText)[\s>]/);
        // The struck word is a red, struck-through run; the new word a blue, underlined one.
        expect(xml).toMatch(new RegExp(`<w:r><w:rPr><w:strike/><w:dstrike w:val="0"/><w:color w:val="FF0000"/></w:rPr><w:t[^>]*>${old}`));
        expect(xml).toMatch(/<w:r><w:rPr><w:strike w:val="0"\/><w:dstrike w:val="0"\/><w:color w:val="0000FF"\/><w:u w:val="single"\/><\/w:rPr><w:t[^>]*>AFTER/);
      }
    });

    it("a header removed", async () => {
      const original = withParts({ header: "BEFOREHEADER", footnote: "A note." });
      const { notes, counts, accepted, rejected } = await compare(original, withParts({ footnote: "A note." }));
      expect(counts).toEqual([2, 0]);
      expect(notes).toEqual(["the last section's header was removed; its reference is not a revision, so accept-all leaves the header empty"]);
      expect(texts(rejected)).toEqual(texts(DocxDocument.load(original)));
      expect(texts(accepted)).toEqual({ "header:default": "", footnote: "A note." });
    });

    it("a header added", async () => {
      const revised = withParts({ header: "AFTERHEADER" });
      const { notes, counts, accepted, rejected } = await compare(withParts({}), revised);
      expect(counts).toEqual([0, 2]);
      expect(notes).toEqual(["the last section's header was added; its reference is not a revision, so reject-all leaves the header empty"]);
      expect(texts(accepted)).toEqual(texts(DocxDocument.load(revised)));
      expect(texts(rejected)).toEqual({ "header:default": "" });
    });

    it("a first-page header added", async () => {
      const original = withParts({ header: "Shared header" });
      const revised = withParts({ header: "Shared header", first: "FIRSTPAGE" });
      const { notes, counts, accepted, rejected } = await compare(original, revised);
      expect(counts).toEqual([0, 2]);
      expect(notes).toEqual(["the last section's first-page header was added; its reference is not a revision, so reject-all leaves the header empty"]);
      expect(texts(accepted)).toEqual(texts(DocxDocument.load(revised)));
      expect(texts(rejected)).toEqual({ ...texts(DocxDocument.load(original)), "header:first": "" });
    });

    it("a footer removed", async () => {
      const original = withParts({ header: "Shared header", footer: "BEFOREFOOTER" });
      const revised = withParts({ header: "Shared header" });
      const { notes, counts, accepted, rejected } = await compare(original, revised);
      expect(counts).toEqual([2, 0]);
      expect(notes).toEqual(["the last section's footer was removed; its reference is not a revision, so accept-all leaves the footer empty"]);
      expect(texts(rejected)).toEqual(texts(DocxDocument.load(original)));
      expect(texts(accepted)).toEqual({ ...texts(DocxDocument.load(revised)), "footer:default": "" });
    });
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
    await fails(join(dir, "ole.docx"), good, join(dir, "x.docx")).toThrow(/ole\.docx: the compound file cannot be read/);
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

  it("refuses two outputs on one path, naming them", async () => {
    const options = { author: AUTHOR, date: DATE };
    const out = join(dir, "collide.docx");
    await expect(compareFiles(good, good, out, { ...options, pdf: { path: out } })).rejects.toMatchObject({ code: "output_collision", message: `${out}: the result and the PDF must be different files` });
    await expect(compareFiles(good, good, out, { ...options, printable: { path: join(dir, "p.docx") }, pdf: { path: join(dir, "p.docx") } })).rejects.toMatchObject({ code: "output_collision", message: expect.stringContaining("the printable copy and the PDF must be different files") });
    expect(existsSync(out)).toBe(false);
  });

  it("refuses to overwrite an input, and to write into a missing folder", async () => {
    await fails(good, good, good).toThrow(/must not overwrite an input/);
    await fails(good, good, join(dir, "no-such-folder", "x.docx")).toThrow(/folder does not exist/);
  });

  it("refuses parts over what it holds in memory, from the directory's sizes, before inflating anything", async () => {
    // A .docx whose directory declares each named part at `declared` bytes
    // (the entries themselves are tiny): what a zip bomb declares, checked
    // before any inflating, so the compare never pays for it.
    const declaring = (name: string, declared: [string, number][]): string => {
      const path = join(dir, name);
      const writer = new ZipWriter(path);
      for (const [part, bytes] of Object.entries(unzipSync(makeDocx({ "word/document.xml": wrapDocument(para("Text.")) })))) {
        if (!declared.some(([d]) => d === part)) writer.addBuffer(part, bytes);
      }
      for (const [part, size] of declared) {
        const data = Buffer.from("<a/>");
        writer.addEntry(part, 0, crc32(data), data.length, size, (write) => write(data));
      }
      writer.close();
      return path;
    };
    // The document's relationships are read before the total is known: a part on its own.
    const onePart = declaring("one-big-part.docx", [["word/_rels/document.xml.rels", (3 << 20) + 1]]);
    const manyParts = declaring("many-big-parts.docx", [["word/styles.xml", 1 << 20], ["word/numbering.xml", 1 << 20], ["word/header1.xml", 1 << 20]]);
    const bigBody = declaring("big-body.docx", [["word/media/image1.png", 65 << 20]]);
    await fails(good, onePart, join(dir, "x.docx")).toThrow(/one-big-part\.docx: word\/_rels\/document\.xml\.rels is 4 MB; the parts held in memory are at most 3 MB together, so that a compare fits in a 256 MB heap/);
    await fails(manyParts, good, join(dir, "x.docx")).toThrow(/many-big-parts\.docx: the parts held in memory .* come to 4 MB; at most 3 MB is supported, so that a compare fits in a 256 MB heap/);
    for (const [path, file] of [[onePart, "revised"], [manyParts, "original"]] as const) {
      const error = await errorOf(file === "original" ? path : good, file === "revised" ? path : good, join(dir, "x.docx"));
      expect([error.code, error.file]).toEqual(["too_large", file]);
    }
    // A binary part is never inflated, so no limit applies: it is copied as it is.
    await compareFiles(good, bigBody, join(dir, "x.docx"), options);
  });

  it("compares a document at the limit, in its costliest shape", async () => {
    // The parts outside the body at exactly MAX_HELD_BYTES, the bulk a
    // header of empty paragraphs: the shape that costs the most heap per
    // byte (bench/memory/big-parts.mjs measures the heap this takes).
    const atLimit = (name: string, over: number): string => {
      const base = makeDocx({
        "word/document.xml": wrapDocument(para("Text.") + `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>`),
        "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
      });
      const files = unzipSync(base);
      const held = Object.entries(files).filter(([part]) => /\.(xml|rels)$/.test(part) && part !== "word/document.xml").reduce((n, [, bytes]) => n + bytes.length, 0);
      const frame = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:hdr>`;
      const paragraphs = Math.floor((MAX_HELD_BYTES - held - frame.length + over) / "<w:p/>".length);
      files["word/header1.xml"] = strToU8(frame.replace("></w:hdr>", `>${"<w:p/>".repeat(paragraphs)}</w:hdr>`));
      const path = join(dir, name);
      writeFileSync(path, zipSync(files, { level: 1 }));
      return path;
    };
    const original = atLimit("at-limit-o.docx", 0);
    const over = atLimit("at-limit-over.docx", "<w:p/>".length);
    const { revisions } = await compareFiles(original, original, join(dir, "at-limit-out.docx"), options);
    expect(revisions.insertion).toBe(0);
    const error = await errorOf(original, over, join(dir, "x.docx"));
    expect([error.code, error.file]).toEqual(["too_large", "revised"]);
  }, 120000);
});

describe("the pdf option", () => {
  it("writes the redline as a PDF beside the result, with the printable copy a working file", async () => {
    const original = join(dir, "pdf-o.docx");
    const revised = join(dir, "pdf-r.docx");
    writeFileSync(original, docx(para("The quick brown fox.")));
    writeFileSync(revised, docx(para("The quick red fox.")));
    const pdf = join(dir, "pdf-redline.pdf");
    const { revisions } = await compareFiles(original, revised, join(dir, "pdf-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: pdf } });
    expect(revisions.insertion).toBe(1);
    const bytes = readFileSync(pdf);
    expect(bytes.subarray(0, 5).toString("latin1")).toBe("%PDF-");
    expect(bytes.toString("latin1").trimEnd().endsWith("%%EOF")).toBe(true);
  });

  it("scans pdf.extraFontDirs ahead of the default font directories", async () => {
    // A copy of Carlito renamed "Zorbaq7" in its name table, in a directory of its own: no default directory holds the family.
    const fontDir = mkdtempSync(join(tmpdir(), "compare-fonts-"));
    const carlito = fileURLToPath(new URL("../../../node_modules/@fontsource/carlito/files/carlito-latin-400-normal.woff", import.meta.url));
    const sfnt = Buffer.from(new Face(carlito).sfnt);
    for (const [from, to] of [[Buffer.from("Carlito", "latin1"), Buffer.from("Zorbaq7", "latin1")], [Buffer.from("Carlito", "utf16le").swap16(), Buffer.from("Zorbaq7", "utf16le").swap16()]]) {
      for (let at = sfnt.indexOf(from); at !== -1; at = sfnt.indexOf(from, at + from.length)) to.copy(sfnt, at);
    }
    writeFileSync(join(fontDir, "zorbaq7.ttf"), sfnt);
    const run = (text: string) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="Zorbaq7" w:hAnsi="Zorbaq7"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const original = join(dir, "pdf-fonts-o.docx");
    const revised = join(dir, "pdf-fonts-r.docx");
    writeFileSync(original, docx(run("The quick brown fox.")));
    writeFileSync(revised, docx(run("The quick red fox.")));
    const pdf = join(dir, "pdf-fonts.pdf");
    await compareFiles(original, revised, join(dir, "pdf-fonts-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: pdf, extraFontDirs: [fontDir] } });
    expect(readFileSync(pdf).toString("latin1")).toMatch(/\/BaseFont \/\w+\+Zorbaq7/);
    const without = join(dir, "pdf-fonts-without.pdf");
    await compareFiles(original, revised, join(dir, "pdf-fonts-out2.docx"), { author: AUTHOR, date: DATE, pdf: { path: without } });
    expect(readFileSync(without).toString("latin1")).not.toContain("Zorbaq7");
  });

  it("ends the printable copy with a summary built from the counts, and returns the PDF's pages", async () => {
    const original = join(dir, "pdf-sum-o.docx");
    const revised = join(dir, "pdf-sum-r.docx");
    writeFileSync(original, docx(para("The quick brown fox.")));
    writeFileSync(revised, docx(para("The quick red fox.")));
    const printable = join(dir, "pdf-sum-printable.docx");
    const summary = (revisions: Record<string, number>): string =>
      `<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>Insertions: ${revisions.insertion}</w:t></w:r></w:p>`;
    const { pdfPages, missingCharacters } = await compareFiles(original, revised, join(dir, "pdf-sum-out.docx"), {
      author: AUTHOR,
      date: DATE,
      printable: { path: printable, beforeSectPr: summary },
      pdf: { path: join(dir, "pdf-sum.pdf") },
    });
    expect(pdfPages).toBe(2);
    expect(missingCharacters).toEqual([]);
    expect(strFromU8(unzipSync(readFileSync(printable))["word/document.xml"])).toContain("Insertions: 1");
  });

  it("lists the picture parts the PDF drew as empty frames", async () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const picture =
      `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="p.bmp"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip ${R} r:embed="rId5"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const rels =
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.bmp"/></Relationships>`;
    // A BMP: a format the renderer does not read.
    const bytes = makeDocxWithMedia(
      { "word/document.xml": `<w:document ${W_NS} ${R}><w:body>${para("Before.")}${picture}${para("After.")}</w:body></w:document>`, "word/_rels/document.xml.rels": rels },
      { "word/media/image1.bmp": new Uint8Array([0x42, 0x4d, 0, 0, 0, 0]) },
    );
    const original = join(dir, "bmp-o.docx");
    const revised = join(dir, "bmp-r.docx");
    writeFileSync(original, bytes);
    writeFileSync(revised, bytes);
    const { pdfPages, unsupportedImages } = await compareFiles(original, revised, join(dir, "bmp-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: join(dir, "bmp.pdf") } });
    expect(pdfPages).toBe(1);
    expect(unsupportedImages).toEqual(["word/media/image1.bmp"]);
  });

  it("calls onResultWritten once the result and the printable copy exist, before the PDF is rendered", async () => {
    const original = join(dir, "hook-o.docx");
    const revised = join(dir, "hook-r.docx");
    writeFileSync(original, docx(para("The quick brown fox.")));
    writeFileSync(revised, docx(para("The quick red fox.")));
    const out = join(dir, "hook-out.docx");
    const pdf = join(dir, "hook.pdf");
    let seen: { result: boolean; pdf: boolean } | null = null;
    const { pdfPages } = await compareFiles(original, revised, out, {
      author: AUTHOR,
      date: DATE,
      pdf: { path: pdf },
      onResultWritten: async () => {
        await new Promise((resolve) => setTimeout(resolve, 5));
        seen = { result: existsSync(out), pdf: existsSync(pdf) };
      },
    });
    expect(seen).toEqual({ result: true, pdf: false });
    expect(pdfPages).toBe(1);
  });

  it("fails as a CompareError when the PDF path cannot be opened, and leaves no partial PDF", async () => {
    const original = join(dir, "pdf-err-o.docx");
    const revised = join(dir, "pdf-err-r.docx");
    writeFileSync(original, docx(para("The quick brown fox.")));
    writeFileSync(revised, docx(para("The quick red fox.")));
    // A folder as the PDF path: its own folder exists, and opening it to write fails.
    const pdf = dir;
    const error = await compareFiles(original, revised, join(dir, "pdf-err-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: pdf } }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(CompareError);
    expect((error as CompareError).code).toBe("failed");
    expect((error as CompareError).message).toContain(pdf);
    // A render that fails after the PDF was opened takes the partial file
    // with it (Mongolian is the script the renderer still cannot shape).
    writeFileSync(revised, docx(para("ᠮᠣᠩᠭᠣᠯ ᠬᠡᠯᠡ")));
    const scriptPdf = join(dir, "pdf-err-script.pdf");
    const refused = await compareFiles(original, revised, join(dir, "pdf-err-out2.docx"), { author: AUTHOR, date: DATE, pdf: { path: scriptPdf } }).catch((e: unknown) => e);
    expect((refused as CompareError).code).toBe("unsupported_script");
    expect(existsSync(scriptPdf)).toBe(false);
  });

  it("reports a script the PDF renderer cannot shape as unsupported_script", async () => {
    const original = join(dir, "pdf-mongolian-o.docx");
    const revised = join(dir, "pdf-mongolian-r.docx");
    writeFileSync(original, docx(para("ᠮᠣᠩᠭᠣᠯ")));
    writeFileSync(revised, docx(para("ᠮᠣᠩᠭᠣᠯ ᠬᠡᠯᠡ")));
    const error = await compareFiles(original, revised, join(dir, "pdf-mongolian-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: join(dir, "pdf-mongolian.pdf") } }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(CompareError);
    expect((error as CompareError).code).toBe("unsupported_script");
    expect((error as CompareError).message).toContain("Mongolian");
  });

  it.skipIf(!hasPdftotext)("names no author in the PDF with authors: false, the summary page included", async () => {
    const commented = (words: string): Uint8Array =>
      makeDocx({
        "word/document.xml": wrapDocument(`<w:p><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">${words}</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`),
        "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS}><w:comment w:id="0" w:author="Carol Commenter" w:initials="CC" w:date="2026-08-14T09:00:00Z"><w:p><w:r><w:t>Check the term.</w:t></w:r></w:p></w:comment></w:comments>`,
      });
    const original = join(dir, "anon-o.docx");
    const revised = join(dir, "anon-r.docx");
    writeFileSync(original, commented("The term is two years."));
    writeFileSync(revised, commented("The term is three years."));
    const pdf = (name: string) => join(dir, `anon-${name}.pdf`);
    const run = (name: string, authors?: boolean) =>
      compareFiles(original, revised, join(dir, `anon-${name}.docx`), { author: "Dan Drafter", date: DATE, summary: { originalName: "a.docx", revisedName: "b.docx" }, pdf: { path: pdf(name), comments: true, authors } });
    await run("shown");
    await run("hidden", false);
    const text = (name: string) => execFileSync("pdftotext", [pdf(name), "-"], { encoding: "utf8" }).replace(/\s+/g, " ");
    expect(text("shown")).toContain("Carol Commenter");
    const hidden = text("hidden");
    for (const name of ["Carol", "Commenter", "Dan", "Drafter"]) expect(hidden).not.toContain(name);
    expect(hidden).not.toMatch(/\bCC\b/);
    expect(hidden).toContain("8/14/2026 UTC Check the term.");
    expect(hidden).toContain("Summary report");
    const raw = readFileSync(pdf("hidden"), "latin1");
    for (const name of ["Carol", "Dan Drafter"]) expect(raw).not.toContain(name);
  });

  it("renders Arabic to the PDF now that the renderer shapes it", async () => {
    const original = join(dir, "pdf-arabic-o.docx");
    const revised = join(dir, "pdf-arabic-r.docx");
    writeFileSync(original, docx(para("مرحبا")));
    writeFileSync(revised, docx(para("مرحبا بالعالم")));
    const pdf = join(dir, "pdf-arabic.pdf");
    const { pdfPages } = await compareFiles(original, revised, join(dir, "pdf-arabic-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: pdf } });
    expect(pdfPages).toBe(1);
    expect(existsSync(pdf)).toBe(true);
  });
});

describe("the command line", () => {
  it("refuses a --pdf path that is an input or the result", () => {
    const original = join(dir, "cli-o.docx");
    const revised = join(dir, "cli-r.docx");
    writeFileSync(original, docx(para("One.")));
    writeFileSync(revised, docx(para("Two.")));
    const out = join(dir, "cli-out.docx");
    const run = (pdf: string) =>
      spawnSync(process.execPath, [fileURLToPath(new URL("../../../node_modules/.bin/vite-node", import.meta.url)), fileURLToPath(new URL("../src/cli.ts", import.meta.url)), original, revised, "-o", out, "--author", "A", "--pdf", pdf], { encoding: "utf8" });
    for (const pdf of [out, original]) {
      const result = run(pdf);
      expect(result.status).toBe(2);
      expect(result.stderr).toContain("--pdf FILE must not be an input or the result");
    }
    expect(existsSync(out)).toBe(false);
  }, 60000);

  it("takes --hide-authors only with --pdf", () => {
    const original = join(dir, "cli-hide-o.docx");
    const revised = join(dir, "cli-hide-r.docx");
    writeFileSync(original, docx(para("One.")));
    writeFileSync(revised, docx(para("Two.")));
    const out = join(dir, "cli-hide-out.docx");
    const pdf = join(dir, "cli-hide.pdf");
    const run = (...extra: string[]) =>
      spawnSync(process.execPath, [fileURLToPath(new URL("../../../node_modules/.bin/vite-node", import.meta.url)), fileURLToPath(new URL("../src/cli.ts", import.meta.url)), original, revised, "-o", out, "--author", "A", "--hide-authors", ...extra], { encoding: "utf8" });
    const refused = run();
    expect(refused.status).toBe(2);
    expect(refused.stderr).toContain("--hide-authors needs --pdf FILE");
    expect(run("--pdf", pdf).status).toBe(0);
    expect(existsSync(pdf)).toBe(true);
  }, 60000);
});

describe("the summary option", () => {
  const STYLES =
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="1F4E79"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="360" w:line="480" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="360" w:line="480" w:lineRule="auto"/><w:ind w:left="720" w:firstLine="720"/><w:jc w:val="both"/></w:pPr>` +
    `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="1F4E79"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:u w:val="single"/></w:rPr></w:style>` +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
  const CONTENT_TYPES = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
  /** The summary table's rows as their cells' text. */
  const tableRows = (documentXml: string): string[][] => {
    const table = /<w:tbl>[\s\S]*<\/w:tbl>/.exec(documentXml)?.[0] ?? "";
    return [...table.matchAll(/<w:tr>([\s\S]*?)<\/w:tr>/g)].map((row) =>
      [...row[1].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map((cell) =>
        [...cell[1].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")).join(""),
      ),
    );
  };
  const summary = { originalName: "Draft <1> & co.docx", revisedName: "Draft 2.docx" };

  it("ends the printable copy with the summary table: date, names, counts, total", async () => {
    const original = join(dir, "sum-o.docx");
    const revised = join(dir, "sum-r.docx");
    writeFileSync(original, docx(para("The quick brown fox.") + para("A second paragraph.")));
    writeFileSync(revised, docx(para("The quick red fox.") + para("A second paragraph.", { pPr: '<w:jc w:val="center"/>' })));
    const printable = join(dir, "sum-printable.docx");
    const out = join(dir, "sum-out.docx");
    const { revisions } = await compareFiles(original, revised, out, { author: AUTHOR, date: DATE, printable: { path: printable }, summary });
    expect(revisions).toMatchObject({ insertion: 1, deletion: 1, paragraphFormat: 1 });
    const xml = strFromU8(unzipSync(readFileSync(printable))["word/document.xml"]);
    expect(tableRows(xml)).toEqual([
      [`Summary report: document comparison done on ${new Date(DATE).toLocaleString("en-US", { timeZone: "UTC" })} UTC`],
      ["Original:", "Draft <1> & co.docx"],
      ["Modified:", "Draft 2.docx"],
      ["Add", "1"],
      ["Delete", "1"],
      ["Table Insert", "0"],
      ["Table Delete", "0"],
      ["Format changes", "1"],
      ["Total Changes:", "3"],
    ]);
    // A page break ahead of the table, the table ahead of the final sectPr, and nothing of it in the result.
    expect(xml).toMatch(/<w:br w:type="page"\/>[\s\S]*<w:tbl>[\s\S]*<\/w:tbl><\/w:body>/);
    expect(xml).toContain("Draft &lt;1&gt; &amp; co.docx");
    expect(strFromU8(unzipSync(readFileSync(out))["word/document.xml"])).not.toContain("Summary report");
  });

  it("heads the table with the caller's title, and the date unless it is turned off", async () => {
    const original = join(dir, "sum-title-o.docx");
    const revised = join(dir, "sum-title-r.docx");
    writeFileSync(original, docx(para("The quick brown fox.")));
    writeFileSync(revised, docx(para("The quick red fox.")));
    const printable = join(dir, "sum-title-printable.docx");
    const heading = async (extra: { title?: string; date?: boolean }): Promise<{ row: string[]; xml: string }> => {
      await compareFiles(original, revised, join(dir, "sum-title-out.docx"), { author: AUTHOR, date: DATE, printable: { path: printable }, summary: { ...summary, ...extra } });
      const xml = strFromU8(unzipSync(readFileSync(printable))["word/document.xml"]);
      return { row: tableRows(xml)[0], xml };
    };
    const when = `${new Date(DATE).toLocaleString("en-US", { timeZone: "UTC" })} UTC`;
    const titled = await heading({ title: "Cobbery <comparison> & results" });
    expect(titled.row).toEqual([`Cobbery <comparison> & results, ${when}`]);
    expect(titled.xml).toContain("Cobbery &lt;comparison&gt; &amp; results, ");
    expect((await heading({ title: "Cobbery comparison results", date: false })).row).toEqual(["Cobbery comparison results"]);
    expect((await heading({ date: false })).row).toEqual(["Summary report: document comparison"]);
    expect((await heading({})).row).toEqual([`Summary report: document comparison done on ${when}`]);
  });

  it("lists each distinct note once, after a custom beforeSectPr", async () => {
    // Two changed paragraphs holding inline math: the same note twice.
    const math = (text: string) =>
      `<w:p><w:r><w:t xml:space="preserve">${text} </w:t></w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>`;
    const original = join(dir, "sum-notes-o.docx");
    const revised = join(dir, "sum-notes-r.docx");
    writeFileSync(original, docx(math("One") + math("Two")));
    writeFileSync(revised, docx(math("Uno") + math("Dos")));
    const printable = join(dir, "sum-notes-printable.docx");
    const notes: CompareNote[] = [];
    await compareFiles(original, revised, join(dir, "sum-notes-out.docx"), {
      author: AUTHOR,
      date: DATE,
      onNote: (n) => notes.push(n),
      printable: { path: printable, beforeSectPr: `<w:p><w:r><w:t>Custom first</w:t></w:r></w:p>` },
      summary,
    });
    expect(notes.filter((n) => n.code === "coarse-paragraph")).toHaveLength(2);
    const xml = strFromU8(unzipSync(readFileSync(printable))["word/document.xml"]);
    expect(tableRows(xml).slice(9)).toEqual([["Note: a paragraph holding inline math or other content a word diff cannot rebuild changed; struck and reinserted whole"]]);
    expect(xml.indexOf("Custom first")).toBeLessThan(xml.indexOf("<w:tbl>"));
  });

  it("keeps its own formatting under a justified, 14 pt, colored, spaced Normal style", async () => {
    const parts = (body: string): Uint8Array => makeDocx({ "[Content_Types].xml": CONTENT_TYPES, "word/styles.xml": STYLES, "word/document.xml": wrapDocument(body) });
    const original = join(dir, "sum-styled-o.docx");
    const revised = join(dir, "sum-styled-r.docx");
    writeFileSync(original, parts(para("The quick brown fox jumps over the lazy dog.")));
    writeFileSync(revised, parts(para("The quick red fox jumps over the lazy dog.")));
    const printable = join(dir, "sum-styled-printable.docx");
    const pdf = join(dir, "sum-styled.pdf");
    const { pdfPages } = await compareFiles(original, revised, join(dir, "sum-styled-out.docx"), { author: AUTHOR, date: DATE, printable: { path: printable }, pdf: { path: pdf }, summary });
    expect(pdfPages).toBe(2);
    const xml = strFromU8(unzipSync(readFileSync(printable))["word/document.xml"]);
    const table = /<w:tbl>[\s\S]*<\/w:tbl>/.exec(xml)![0];
    // Nine rows: the title's one cell, then eight rows of two.
    expect(table.match(/<w:p>/g)).toHaveLength(17);
    expect(table.match(/<w:jc w:val="left"\/>/g)).toHaveLength(17);
    expect(table.match(/<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"\/>/g)).toHaveLength(17);
    expect(table.match(/<w:ind w:left="0" w:right="0" w:firstLine="0"\/>/g)).toHaveLength(17);
    expect(table.match(/<w:rPr>/g)).toHaveLength(17);
    expect(table.match(/<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"\/>/g)).toHaveLength(17);
    expect(table.match(/<w:sz w:val="20"\/>/g)).toHaveLength(17);
    expect(table.match(/<w:u w:val="none"\/>/g)).toHaveLength(17);
    expect(table).toContain('<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>');
    expect(table).toContain('<w:tblGrid><w:gridCol w:w="6926"/><w:gridCol w:w="2434"/></w:tblGrid>');
    expect(table).not.toContain("w:pStyle");
    if (hasPdftotext) {
      const page2 = execFileSync("pdftotext", ["-layout", pdf, "-"], { encoding: "utf8" }).split("\f")[1];
      expect(page2).toContain("Summary report: document comparison done on");
      expect(page2).toMatch(/Original:\s+Draft <1> & co.docx/);
      expect(page2).toMatch(/Add\s+1/);
      expect(page2).toMatch(/Total Changes:\s+2/);
    }
  });
});

describe("the printable copy's marks", () => {
  /** Every content stream of a PDF, inflated and joined. */
  const contentOps = (pdf: string): string => {
    const bytes = readFileSync(pdf);
    const text = bytes.toString("latin1");
    return [...text.matchAll(/\/Length (\d+) >>\nstream\n/g)]
      .map((m) => inflateSync(bytes.subarray(m.index! + m[0].length, m.index! + m[0].length + Number(m[1]))).toString("latin1"))
      .filter((s) => s.includes(" Tf"))
      .join("\n");
  };
  /** The text-showing groups (q … Q) whose fill is `rgb`, with whether an underline rule follows the text. */
  const textGroups = (ops: string, rgb: string): { underlined: boolean }[] =>
    [...ops.matchAll(/^q\n([\s\S]*?)\nQ$/gm)]
      .filter((m) => m[1].includes(`${rgb} rg`) && m[1].includes("BT"))
      .map((m) => ({ underlined: /ET\n[\d. ]+ re f/.test(m[1]) }));
  /** The strike-through rules (a filled rect of its own) in `rgb`. */
  const strikes = (ops: string, rgb: string): number => (ops.match(new RegExp(`^q ${rgb} rg [\\d. ]+ re f Q$`, "gm")) ?? []).length;
  const BLUE = "0 0 1";
  const RED = "1 0 0";
  /** A 4 x 4 solid PNG. */
  const solidPng = (): Uint8Array => {
    const chunk = (type: string, data: Buffer): Buffer => {
      const head = Buffer.alloc(8);
      head.writeUInt32BE(data.length, 0);
      head.write(type, 4, "latin1");
      const crc = Buffer.alloc(4);
      crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, "latin1"), data])), 0);
      return Buffer.concat([head, data, crc]);
    };
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(4, 0);
    ihdr.writeUInt32BE(4, 4);
    ihdr.set([8, 2, 0, 0, 0], 8);
    const raw = Buffer.concat(Array.from({ length: 4 }, () => Buffer.from([0, ...Array(4).fill([200, 30, 30]).flat()])));
    return new Uint8Array(Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]));
  };
  const run = (text: string, rPr: string): string => `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const compared = async (name: string, original: Uint8Array, revised: Uint8Array): Promise<{ xml: string; ops: string }> => {
    const o = join(dir, `${name}-o.docx`);
    const r = join(dir, `${name}-r.docx`);
    writeFileSync(o, original);
    writeFileSync(r, revised);
    const printable = join(dir, `${name}-printable.docx`);
    const pdf = join(dir, `${name}.pdf`);
    await compareFiles(o, r, join(dir, `${name}-out.docx`), { author: AUTHOR, date: DATE, printable: { path: printable }, pdf: { path: pdf } });
    return { xml: strFromU8(unzipSync(readFileSync(printable))["word/document.xml"]), ops: contentOps(pdf) };
  };

  it("marks runs whose own properties say black, no underline and no strike", async () => {
    const plain = '<w:color w:val="000000"/><w:sz w:val="24"/><w:u w:val="none"/>';
    const unstruck = '<w:strike w:val="0"/><w:dstrike w:val="false"/><w:color w:val="000000"/>';
    const { xml, ops } = await compared(
      "marks-direct",
      docx(`<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r>${run("brown", unstruck)}<w:r><w:t xml:space="preserve"> fox.</w:t></w:r></w:p>`),
      docx(`<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r>${run("red", plain)}<w:r><w:t xml:space="preserve"> fox.</w:t></w:r></w:p>`),
    );
    // The marks replace the run's own color, underline and strike, in schema order, each once.
    expect(xml).toContain('<w:rPr><w:strike/><w:dstrike w:val="0"/><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">brown</w:t>');
    expect(xml).toContain('<w:rPr><w:strike w:val="0"/><w:dstrike w:val="0"/><w:color w:val="0000FF"/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">red</w:t>');
    expect(textGroups(ops, BLUE)).toEqual([{ underlined: true }]);
    expect(textGroups(ops, RED)).toEqual([{ underlined: false }]);
    expect(strikes(ops, RED)).toBe(1);
    expect(strikes(ops, BLUE)).toBe(0);
  });

  it("marks runs whose style says black, no underline and struck, since direct properties win", async () => {
    const styles =
      `<?xml version="1.0"?><w:styles ${W_NS}>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:strike/><w:color w:val="000000"/><w:u w:val="none"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="Plain"><w:name w:val="Plain"/><w:rPr><w:strike/><w:color w:val="000000"/><w:u w:val="none"/></w:rPr></w:style></w:styles>`;
    const contentTypes = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
    const styled = (body: string): Uint8Array => makeDocx({ "[Content_Types].xml": contentTypes, "word/styles.xml": styles, "word/document.xml": wrapDocument(body) });
    const { xml, ops } = await compared(
      "marks-styled",
      styled(`<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r>${run("brown", '<w:rStyle w:val="Plain"/>')}<w:r><w:t xml:space="preserve"> fox.</w:t></w:r></w:p>`),
      styled(`<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r>${run("red", '<w:rStyle w:val="Plain"/>')}<w:r><w:t xml:space="preserve"> fox.</w:t></w:r></w:p>`),
    );
    expect(xml).toContain('<w:rPr><w:rStyle w:val="Plain"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">red</w:t>');
    // The plain words are struck black by the style; the new word is blue, underlined and not struck, the old one red and struck.
    expect(strikes(ops, "0 0 0")).toBeGreaterThan(0);
    expect(textGroups(ops, BLUE)).toEqual([{ underlined: true }]);
    expect(strikes(ops, BLUE)).toBe(0);
    expect(textGroups(ops, RED)).toEqual([{ underlined: false }]);
    expect(strikes(ops, RED)).toBe(1);
  });

  it("adds the change bar to a paragraph's own borders, in place", async () => {
    const bordered = (jc: string): Uint8Array =>
      docx(`<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>${jc}</w:pPr><w:r><w:t>Bordered text.</w:t></w:r></w:p>`);
    const { xml, ops } = await compared("marks-bar", bordered(""), bordered('<w:jc w:val="center"/>'));
    expect(xml).toContain('<w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:left w:val="single" w:sz="6" w:space="12" w:color="000000"/></w:pBdr><w:jc w:val="center"/></w:pPr>');
    // Two rules: the top border (horizontal) and the change bar (vertical).
    const rules = [...ops.matchAll(/ w ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S/g)].map((m) => m.slice(1).map(Number));
    expect(rules.filter(([x1, y1, x2, y2]) => y1 === y2 && x2 > x1)).toHaveLength(1);
    expect(rules.filter(([x1, y1, x2, y2]) => x1 === x2 && y1 !== y2)).toHaveLength(1);
  });

  it("marks an equation's runs through their w:rPr, and draws the marks", async () => {
    const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    const math = (t: string): string =>
      `<w:p><w:r><w:t xml:space="preserve">Formula: </w:t></w:r><m:oMath ${M}><m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>${t}</m:t></m:r></m:oMath></w:p>`;
    const { xml, ops } = await compared("marks-math", docx(math("a+b=c")), docx(math("a+b=d")));
    // The marks sit in the math run's w:rPr, after its m:rPr, which stays as it was.
    expect(xml).toContain('<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:strike/><w:dstrike w:val="0"/><w:color w:val="FF0000"/></w:rPr><m:t>a+b=c</m:t></m:r>');
    expect(xml).toContain('<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:strike w:val="0"/><w:dstrike w:val="0"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><m:t>a+b=d</m:t>');
    // Each equation piece is drawn in its mark's color: struck red, underlined blue.
    expect(textGroups(ops, RED).length).toBeGreaterThan(2);
    expect(strikes(ops, RED)).toBeGreaterThan(2);
    expect(textGroups(ops, BLUE).filter((g) => g.underlined).length).toBeGreaterThan(2);
  });

  it("bars every paragraph of a table, row or cell whose format changed, and the last paragraph for the final section's", async () => {
    const bordered = `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>`;
    const cell = (text: string, tcPr = ""): string => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>${tcPr}</w:tcPr>${para(text)}</w:tc>`;
    const tbl = (tblW: string, trHeight: string, shd: string): string =>
      `<w:tbl><w:tblPr><w:tblW w:w="${tblW}" w:type="dxa"/>${bordered}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:trHeight w:val="${trHeight}"/></w:trPr>${cell("Cell one", shd)}${cell("Cell two")}</w:tr><w:tr>${cell("Cell three")}${cell("Cell four")}</w:tr></w:tbl>`;
    const sect = (left: number): string => `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="${left}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
    const one = para("Before") + tbl("8000", "300", "") + para("After") + sect(1440);
    const two = para("Before") + tbl("6000", "900", '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>') + para("After") + sect(2880);
    const { xml, ops } = await compared("marks-formats", docx(one), docx(two));
    expect(xml).not.toMatch(/<w:(?:tblPrChange|trPrChange|tcPrChange|sectPrChange|tblGridChange)\b/);
    const bar = '<w:pPr><w:pBdr><w:left w:val="single" w:sz="6" w:space="12" w:color="000000"/></w:pBdr></w:pPr>';
    for (const text of ["Cell one", "Cell two", "Cell three", "Cell four", "After"]) expect(xml).toContain(`${bar}<w:r><w:t xml:space="preserve">${text}</w:t>`);
    expect(xml).not.toContain(`${bar}<w:r><w:t xml:space="preserve">Before</w:t>`);
    // Five change bars drawn: one vertical rule per barred paragraph.
    const rules = [...ops.matchAll(/ w ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S/g)].map((m) => m.slice(1).map(Number));
    expect(rules.filter(([x1, y1, x2, y2]) => x1 === x2 && y1 !== y2)).toHaveLength(5);
    // A column grid change alone, with no table property change, bars the table too; and the copies still flatten whole.
    await check(one, two);
    const grid = (a: number, b: number): string =>
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${bordered}</w:tblPr><w:tblGrid><w:gridCol w:w="${a}"/><w:gridCol w:w="${b}"/></w:tblGrid><w:tr>${cell("Left")}${cell("Right")}</w:tr></w:tbl>`;
    const { printable } = await check(para("Before") + grid(4000, 4000), para("Before") + grid(2000, 6000));
    const gridXml = strFromU8(unzipSync(printable.save())["word/document.xml"]);
    expect(gridXml).not.toContain("tblGridChange");
    expect(gridXml).toContain(`${bar}<w:r><w:t xml:space="preserve">Left</w:t>`);
  });

  it("unlinks a field whose result holds a revision, so the marks print", async () => {
    const field = (instr: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const { xml, ops } = await compared(
      "marks-field",
      docx(`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field("PAGE", "1")}<w:r><w:t xml:space="preserve"> of the document. Date: </w:t></w:r>${field('DATE \\@ "yyyy"', "2025")}<w:r><w:t>.</w:t></w:r></w:p>`),
      docx(`<w:p><w:r><w:t xml:space="preserve">Page of the document. Date: </w:t></w:r>${field('DATE \\@ "yyyy"', "2026")}<w:r><w:t>.</w:t></w:r></w:p>`),
    );
    // Both fields' codes are gone (the PAGE field was deleted whole, the DATE field's result changed); the results stay, marked.
    expect(xml).not.toMatch(/<w:(?:fldChar|instrText|delInstrText)\b/);
    expect(xml).toContain('<w:rPr><w:strike/><w:dstrike w:val="0"/><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">2025</w:t>');
    expect(xml).toContain('<w:rPr><w:strike w:val="0"/><w:dstrike w:val="0"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">2026</w:t>');
    // Drawn as marked text, not as a recomputed page number or today's date: the deleted
    // "1" and the space after it share one strike rule, "2025" takes the other.
    expect(strikes(ops, RED)).toBe(2);
    expect(textGroups(ops, BLUE)).toEqual([{ underlined: true }]);
  });

  it("keeps the visible text that shares a run with a field character when it unlinks the field", () => {
    // A valid field may put text in the runs that hold its characters: text before
    // begin, the separate marker with an unchanged result prefix, text after end.
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const paragraph = parseXml(
      `<w:p ${W}>` +
        `<w:r><w:t xml:space="preserve">Date: </w:t><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/><w:t xml:space="preserve">October </w:t></w:r>` +
        `<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>15</w:delText></w:r></w:del>` +
        `<w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>16</w:t></w:r></w:ins>` +
        `<w:r><w:t xml:space="preserve">, 2026</w:t><w:fldChar w:fldCharType="end"/><w:t xml:space="preserve"> (printed)</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `</w:p>`,
    );
    const xml = serializeXml(flatten(paragraph));
    expect(xml).not.toMatch(/fldCharType="(?:begin|separate|end)"[^]*DATE|DATE[^]*fldCharType="separate"/);
    expect(xml).not.toContain(" DATE ");
    // Every visible piece of the unlinked field survives, in order, with the marks on the changed day.
    const text = [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
    expect(text).toBe("Date: October 1516, 2026 (printed)3");
    expect(xml).toMatch(/<w:color w:val="FF0000"\/><\/w:rPr><w:t>15<\/w:t>/);
    // The PAGE field holds no revision: it stays a field.
    expect(xml).toContain(" PAGE ");
    expect((xml.match(/w:fldCharType/g) ?? []).length).toBe(3);
  });

  describe("unlinks a field that a revision touches anywhere, so no field code or w:delInstrText is left", () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const DEL = '<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z">';
    const INS = '<w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z">';
    const chars = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
    const code = (instr: string, deleted = false): string => `<w:r><w:${deleted ? "delInstrText" : "instrText"} xml:space="preserve"> ${instr} </w:${deleted ? "delInstrText" : "instrText"}></w:r>`;
    const texts = (xml: string): string => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("|");
    const flat = (body: string): string => serializeXml(flatten(parseXml(`<w:body ${W}>${body}</w:body>`)));

    it("a deleted field with no result: its code goes, and nothing is left of it", () => {
      const xml = flat(`<w:p>${DEL}${chars("begin")}${code("DATE", true)}${chars("separate")}${chars("end")}</w:del></w:p>`);
      expect(xml).not.toMatch(/<w:(?:fldChar|instrText|delInstrText)\b/);
      expect(xml).toBe(`<w:body ${W}><w:p/></w:body>`);
    });

    it("a field inside a hyperlink or a content control, its code replaced", () => {
      const field = `${chars("begin")}${DEL}${code("REF a", true)}</w:del>${INS}${code("REF b")}</w:ins>${chars("separate")}<w:r><w:t>result</w:t></w:r>${chars("end")}`;
      for (const [open, close] of [['<w:hyperlink r:id="rId1">', "</w:hyperlink>"], ['<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>', "</w:sdtContent></w:sdt>"]]) {
        const xml = flat(`<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${open}${field}${close}</w:p>`);
        expect(xml).not.toMatch(/<w:(?:fldChar|instrText|delInstrText)\b/);
        expect(texts(xml)).toBe("See |result");
        expect(xml).toContain(open);
      }
    });

    it("a text box in a deleted run: the fields inside it go the same way", () => {
      const box = (content: string): string => `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
      const xml = flat(`<w:p>${DEL}${box(`<w:p>${chars("begin")}${code("PAGE", true)}${chars("separate")}<w:r><w:delText>7</w:delText></w:r>${chars("end")}</w:p>`)}</w:del></w:p>`);
      expect(xml).not.toMatch(/<w:(?:fldChar|instrText|delInstrText|delText)\b/);
      expect(xml).toMatch(/<w:color w:val="FF0000"\/><\/w:rPr><w:t>7<\/w:t>/);
    });

    it("a field spanning paragraphs (a table of contents) deleted whole", () => {
      const xml = flat(
        `<w:p>${DEL}${chars("begin")}${code('TOC \\o "1-3"', true)}${chars("separate")}<w:r><w:delText>Heading 1</w:delText></w:r></w:del></w:p>` +
          `<w:p>${DEL}<w:r><w:delText>Heading 2</w:delText></w:r>${chars("end")}</w:del></w:p>`,
      );
      expect(xml).not.toMatch(/<w:(?:instrText|delInstrText|delText)\b/);
      expect(xml).not.toMatch(/fldCharType="(?:begin|separate)"/);
      expect(texts(xml)).toBe("Heading 1|Heading 2");
    });

    it("a replaced field: the struck code and the inserted code both go, the text around them stays", () => {
      // The compare strikes the old begin and code and inserts new text with a new begin and
      // code, while the separate, result and end are shared: one field on each side.
      const xml = flat(
        `<w:p>${DEL}${chars("begin")}${code("MERGEFIELD x", true)}</w:del>${INS}<w:r><w:t xml:space="preserve">new </w:t></w:r>${chars("begin")}${code("PAGE")}</w:ins>` +
          `<w:r><w:fldChar w:fldCharType="separate"/><w:t xml:space="preserve">result</w:t></w:r>${chars("end")}</w:p><w:p><w:r><w:t>After.</w:t></w:r></w:p>`,
      );
      expect(xml).not.toMatch(/<w:(?:fldChar|instrText|delInstrText)\b/);
      expect(texts(xml)).toBe("new |result|After.");
      expect(xml).toMatch(/<w:u w:val="single"\/><\/w:rPr><w:t xml:space="preserve">new <\/w:t>/);
    });
  });

  it("unbinds a content control whose content holds a revision, so the property does not refill it", () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const control = (content: string): string =>
      `<w:sdt><w:sdtPr><w:alias w:val="Author"/><w:dataBinding w:xpath="/ns1:coreProperties[1]/ns0:creator[1]" w:storeItemID="{1}"/><w:text/></w:sdtPr><w:sdtContent><w:p>${content}</w:p></w:sdtContent></w:sdt>`;
    const changed = control(`<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Old Name</w:delText></w:r></w:del><w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>New Name</w:t></w:r></w:ins>`);
    const same = control("<w:r><w:t>Same Name</w:t></w:r>");
    const xml = serializeXml(flatten(parseXml(`<w:body ${W}>${changed}${same}</w:body>`)));
    expect(xml.match(/<w:dataBinding /g)).toHaveLength(1);
    expect(xml).toMatch(/<w:sdtPr><w:alias w:val="Author"\/><w:text\/><\/w:sdtPr><w:sdtContent><w:p><w:r><w:rPr><w:strike\/>/);
    expect(xml).toContain('<w:dataBinding w:xpath="/ns1:coreProperties[1]/ns0:creator[1]" w:storeItemID="{1}"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Same Name</w:t>');
  });

  it("bars the paragraph of an inserted picture", async () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const picture =
      `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip ${R} r:embed="rId5"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const rels =
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`;
    const png = solidPng();
    const withPicture = (body: string): Uint8Array =>
      makeDocxWithMedia({ "word/document.xml": `<w:document ${W_NS} ${R}><w:body>${body}</w:body></w:document>`, "word/_rels/document.xml.rels": rels }, { "word/media/image1.png": png });
    const { xml, ops } = await compared("marks-picture", withPicture(para("Before.") + para("After.")), withPicture(para("Before.") + picture + para("After.")));
    expect(xml).toMatch(/<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="6" w:space="12" w:color="000000"\/><\/w:pBdr><\/w:pPr><w:r><w:rPr>[^<]*(?:<[^>]*>)*<\/w:rPr><w:drawing/);
    const rules = [...ops.matchAll(/ w ([\d.]+) ([\d.]+) m ([\d.]+) ([\d.]+) l S/g)].map((m) => m.slice(1).map(Number));
    expect(rules.filter(([x1, y1, x2, y2]) => x1 === x2 && y1 !== y2)).toHaveLength(1);
    expect(ops).toMatch(/\/Im\d+ Do/);
  });

  it("underlines and strikes a space alone, so a whitespace-only change shows", async () => {
    const { ops } = await compared(
      "marks-space",
      docx(para("Join these: wordone wordtwo; remove here: one two three.")),
      docx(para("Join these: wordonewordtwo; remove here: one two  three.")),
    );
    // The inserted space is drawn blue with an underline; the deleted one red with a strike rule.
    expect(textGroups(ops, BLUE).some((g) => g.underlined)).toBe(true);
    expect(strikes(ops, RED)).toBeGreaterThan(0);
  });

  it("marks a nested table's text in an inserted or a deleted row", async () => {
    const cell = (body: string): string => `<w:tc><w:tcPr/>${body}</w:tc>`;
    const inner = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${cell(para("NESTEDTEXT"))}</w:tr></w:tbl>`;
    const outer = (rows: string[]): string => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows.map((r) => `<w:tr>${r}</w:tr>`).join("")}</w:tbl>`;
    const one = para("Before.") + outer([cell(para("Plain row."))]) + para("After.");
    const two = para("Before.") + outer([cell(para("Plain row.")), cell(inner + para("OUTERTEXT"))]) + para("After.");
    for (const [name, original, revised, color, rPr] of [
      ["marks-row-ins", one, two, BLUE, '<w:strike w:val="0"/><w:dstrike w:val="0"/><w:color w:val="0000FF"/><w:u w:val="single"/>'],
      ["marks-row-del", two, one, RED, '<w:strike/><w:dstrike w:val="0"/><w:color w:val="FF0000"/>'],
    ] as const) {
      const { xml, ops } = await compared(name, docx(original), docx(revised));
      expect(xml).not.toMatch(/<w:(?:ins|del|delText)[\s>]/);
      for (const text of ["NESTEDTEXT", "OUTERTEXT"]) expect(xml).toContain(`<w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t>`);
      // Both words drawn in the mark's color: the nested table's and the outer cell's.
      expect(textGroups(ops, color)).toHaveLength(2);
      expect(strikes(ops, color)).toBe(color === RED ? 2 : 0);
    }
    // The printable copy is still the whole document flattened.
    await check(one, two);
    await check(two, one);
  });
});
