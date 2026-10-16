import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { crc32 } from "node:zlib";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
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
        expect(xml).toMatch(new RegExp(`<w:r><w:rPr><w:color w:val="FF0000"/><w:strike/></w:rPr><w:t[^>]*>${old}`));
        expect(xml).toMatch(/<w:r><w:rPr><w:color w:val="0000FF"\/><w:u w:val="single"\/><\/w:rPr><w:t[^>]*>AFTER/);
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

  it("reports a script the PDF renderer cannot shape as unsupported_script", async () => {
    const original = join(dir, "pdf-arabic-o.docx");
    const revised = join(dir, "pdf-arabic-r.docx");
    writeFileSync(original, docx(para("مرحبا")));
    writeFileSync(revised, docx(para("مرحبا بالعالم")));
    const error = await compareFiles(original, revised, join(dir, "pdf-arabic-out.docx"), { author: AUTHOR, date: DATE, pdf: { path: join(dir, "pdf-arabic.pdf") } }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(CompareError);
    expect((error as CompareError).code).toBe("unsupported_script");
    expect((error as CompareError).message).toContain("Arabic");
  });
});
