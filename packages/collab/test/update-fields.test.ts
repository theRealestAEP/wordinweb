import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/** A complex field: begin / instrText / separate / cached result / end. */
function field(instr: string, cached: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t xml:space="preserve">${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function makeDoc(body: string): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(documentXml),
    }),
  );
}

const BODY =
  `<w:p>${field('DATE \\@ "yyyy"', "1999")}</w:p>` +
  `<w:p>${field("AUTHOR", "Nobody")}</w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>`;

function xmlOf(s: DocumentSession): string {
  return s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
}

describe("updateFields: a document-scoped registered operation", () => {
  it("applies the carried results, so two replicas end byte-identical", () => {
    const a = new DocumentSession(makeDoc(BODY));
    const b = new DocumentSession(makeDoc(BODY));
    const intent = {
      kind: "updateFields" as const,
      clientId: "a",
      clientSeq: 1,
      base: 0,
      results: ["2026", "A. Pickett"],
    };
    expect(a.submit(intent).kind).toBe("applied");
    expect(b.submit(intent).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toContain(`<w:t xml:space="preserve">2026</w:t>`);
    expect(xmlOf(a)).toContain(`<w:t xml:space="preserve">A. Pickett</w:t>`);
  });

  it("rejects when the replica's field count has moved, changing nothing", () => {
    // The field count stands in for the stable id a document-scoped operation
    // has no room for: a stale result list applies to no field at all.
    const s = new DocumentSession(makeDoc(BODY));
    const before = xmlOf(s);
    expect(
      s.submit({ kind: "updateFields", clientId: "a", clientSeq: 1, base: 0, results: ["2026"] }).kind,
    ).toBe("rejected");
    expect(xmlOf(s)).toBe(before);
  });

  it("is a clean no-op when the results match what the fields already hold", () => {
    const s = new DocumentSession(makeDoc(BODY));
    expect(
      s.submit({ kind: "updateFields", clientId: "a", clientSeq: 1, base: 0, results: ["1999", "Nobody"] }).kind,
    ).toBe("rejected");
  });

  it("does not create a result run for a field that has never held one", () => {
    // A fresh run would be an id-tracked node this intent carries no id for,
    // so the field keeps its empty result in a room.
    const body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const s = new DocumentSession(makeDoc(body));
    const before = xmlOf(s);
    expect(
      s.submit({ kind: "updateFields", clientId: "a", clientSeq: 1, base: 0, results: ["Someone"] }).kind,
    ).toBe("rejected");
    expect(xmlOf(s)).toBe(before);
  });

  it("bounds the payload", () => {
    const base = { kind: "updateFields" as const, clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ ...base, results: ["ok"] })).toBeNull();
    expect(validateIntent({ ...base, results: "no" as never })).toMatch(/not an array/);
    expect(validateIntent({ ...base, results: [1 as never] })).toMatch(/not a string/);
    expect(validateIntent({ ...base, results: ["x".repeat(4097)] })).toMatch(/too long/);
    expect(validateIntent({ ...base, results: Array(20001).fill("") })).toMatch(/too many/);
  });
});

describe("insertField instruction allowlist", () => {
  const base = { kind: "insertField" as const, clientId: "a", clientSeq: 1, base: 0, runId: 1, nodeIds: [] };

  it("admits the field types the engine evaluates", () => {
    for (const instr of ["AUTHOR", 'DATE \\@ "yyyy"', "STYLEREF Heading1", "SEQ Figure", "FILENAME", "PAGEREF bk \\h"]) {
      expect(validateIntent({ ...base, instruction: instr })).toBeNull();
    }
  });

  it("refuses the field types that reach outside the document", () => {
    for (const instr of [
      'INCLUDETEXT "\\\\\\\\host\\\\share\\\\a.docx"',
      'INCLUDEPICTURE "https://example.com/p.png"',
      'DDEAUTO Excel "C:\\\\a.xls" "R1C1"',
      'LINK Excel.Sheet.12 "C:\\\\a.xls"',
      "MACROBUTTON AutoOpen Click",
      'HYPERLINK "https://example.com"',
    ]) {
      expect(validateIntent({ ...base, instruction: instr })).toMatch(/not allowed/);
    }
  });

  it("refuses a control character that could split the instruction in two", () => {
    expect(validateIntent({ ...base, instruction: "AUTHOR \u0000 INCLUDETEXT x" })).toMatch(/not allowed/);
  });
});
