import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { compareDocuments, compareDocx } from "../src/edit/compare/index.js";
import { makeDocx, wrapDocument } from "./helpers.js";
import { AUTHOR, DATE, para } from "./compare-helpers.js";

/**
 * compareDocx keeps each body as block text and parses only what differs;
 * its result must be compareDocuments(...).save() byte for byte, on every
 * shape the two paths handle differently.
 */
const options = { author: AUTHOR, date: DATE };
const docx = (body: string, parts: Record<string, string> = {}) => makeDocx({ "word/document.xml": wrapDocument(body), ...parts });
const same = (original: Uint8Array, revised: Uint8Array) => {
  const whole = compareDocuments(DocxDocument.load(original), DocxDocument.load(revised), options).save();
  expect(Buffer.from(compareDocx(original, revised, options)).equals(Buffer.from(whole))).toBe(true);
};
const table = (cells: string[]) =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${cells.map((c) => `<w:tr><w:tc><w:tcPr/>${para(c)}</w:tc></w:tr>`).join("")}</w:tbl>`;

describe("compareDocx", () => {
  it("matches compareDocuments on edits, removals, additions and moves around kept text", () => {
    const keep = Array.from({ length: 30 }, (_, i) => para(`Unchanged paragraph number ${i}, long enough to anchor on.`));
    const original = keep.slice(0, 10).join("") + para("The quick brown fox.") + para("A paragraph that goes.") + keep.slice(10, 20).join("") + table(["a", "b"]) + keep.slice(20).join("");
    const revised = keep.slice(0, 10).join("") + para("The quick red fox.") + keep.slice(10, 20).join("") + para("A paragraph that comes.") + table(["a", "c"]) + keep.slice(20, 25).join("") + keep.slice(26).join("") + keep[25];
    same(docx(original), docx(revised));
  });

  it("matches compareDocuments when the documents are identical", () => {
    const body = para("One.") + table(["x"]) + para("Two.");
    same(docx(body), docx(body));
  });

  it("matches compareDocuments with section breaks and empty paragraphs", () => {
    const brk = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>`;
    same(docx(para("One.") + "<w:p/>" + brk + para("Two.")), docx(para("One!") + brk + "<w:p/>" + para("Two.")));
  });

  it("matches compareDocuments with footnotes, including a moved reference", () => {
    const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const withNotes = (body: string) =>
      makeDocx({
        "word/document.xml": `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${R}><w:body>${body}<w:sectPr/></w:body></w:document>`,
        "word/_rels/document.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdN" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`,
        "word/footnotes.xml":
          `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
          `<w:footnote w:id="1"><w:p><w:r><w:t>Note one.</w:t></w:r></w:p></w:footnote></w:footnotes>`,
      });
    const noted = `<w:p><w:r><w:t>Alpha has a note.</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;
    same(withNotes(noted + para("Beta.") + para("Gamma.")), withNotes(para("Beta.") + para("Gamma.") + noted));
  });

  it("goes through compareDocuments when a document holds tracked changes", () => {
    const A = 'w:author="B" w:date="2026-01-01T00:00:00Z"';
    const tracked = `<w:p><w:ins w:id="1" ${A}><w:r><w:t>Inserted earlier. </w:t></w:r></w:ins><w:r><w:t>Text.</w:t></w:r></w:p>`;
    same(docx(para("Text.")), docx(tracked + para("More.")));
  });
});
