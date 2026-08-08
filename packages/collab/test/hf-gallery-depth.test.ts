import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/**
 * Tool-depth wave 3, lane C: the page-number position gallery
 * (insertPageNumberPosition / removePageNumbers) and the Header & Footer
 * preset gallery (insertHeaderFooterPreset) — registered operations —
 * through the canonical wire apply.
 */

function makeDoc(text = "body"): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

const seq = (n: number, base = 700) => Array.from({ length: n }, (_, i) => base + i);

function allRootsXml(doc: DocxDocument): string {
  return doc.editableRoots().map((r) => serializeXml(r)).join("\n----\n");
}

describe("insertPageNumberPosition / removePageNumbers on the wire", () => {
  it("creates the footer part on carried ids and writes an aligned PAGE field", () => {
    const s = new DocumentSession(makeDoc());
    const e = s.submit({
      kind: "insertPageNumberPosition", clientId: "a", clientSeq: 1, base: 0,
      position: "bottom", align: "right", nodeIds: seq(2),
    });
    expect(e.kind).toBe("applied");
    expect(s.doc.footerRoots().length).toBe(1);
    const xml = serializeXml(s.doc.footerRoots()[0]);
    expect(xml).toContain("PAGE");
    expect(xml).toContain(`w:val="right"`);
    // The gallery paragraph + the field's result run took the carried ids.
    expect(s.ids.elOf(700)).toBeTruthy();
    expect(localName(s.ids.elOf(700)!.name)).toBe("p");
  });

  it("a second pick replaces the content (still applies, not a no-op)", () => {
    const s = new DocumentSession(makeDoc());
    s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 1, base: 0, position: "top", align: "left", nodeIds: seq(2) });
    const e = s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 2, base: s.seq, position: "top", align: "center", nodeIds: seq(2, 800) });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.headerRoots()[0]);
    expect(xml).toContain(`w:val="center"`);
    expect(xml).not.toContain(`w:val="left"`);
  });

  it("removePageNumbers strips the field; honest no-op when there is none", () => {
    const s = new DocumentSession(makeDoc());
    s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 1, base: 0, position: "top", align: "center", nodeIds: seq(2) });
    const e = s.submit({ kind: "removePageNumbers", clientId: "a", clientSeq: 2, base: s.seq });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.headerRoots()[0])).not.toContain("PAGE");
    expect(s.submit({ kind: "removePageNumbers", clientId: "a", clientSeq: 3, base: s.seq }).kind).toBe("rejected");
  });

  it("rejects a malformed payload", () => {
    const s = new DocumentSession(makeDoc());
    expect(s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 1, base: 0, position: "middle" as never, align: "left", nodeIds: seq(2) }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 2, base: 0, position: "top", align: "justify" as never, nodeIds: seq(2) }).kind).toBe("rejected");
  });

  it("two sessions converge byte-identically on the whole package", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc());
      s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 1, base: 0, position: "top", align: "center", nodeIds: seq(2) });
      s.submit({ kind: "insertPageNumberPosition", clientId: "a", clientSeq: 2, base: s.seq, position: "bottom", align: "right", nodeIds: seq(2, 800) });
      return allRootsXml(s.doc);
    };
    expect(build()).toBe(build());
  });
});

describe("insertHeaderFooterPreset on the wire", () => {
  it("writes the threeColumn preset with paragraph tab stops and a live PAGE field", () => {
    const s = new DocumentSession(makeDoc());
    const e = s.submit({
      kind: "insertHeaderFooterPreset", clientId: "a", clientSeq: 1, base: 0,
      hfKind: "footer", preset: "threeColumn", nodeIds: seq(6),
    });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.footerRoots()[0]);
    expect(xml).toContain("w:tabs");
    expect(xml).toContain("PAGE");
    expect(xml).toContain("[Company Name]");
  });

  it("titleAndDate writes two paragraphs, the second a live DATE field", () => {
    const s = new DocumentSession(makeDoc());
    const e = s.submit({
      kind: "insertHeaderFooterPreset", clientId: "a", clientSeq: 1, base: 0,
      hfKind: "header", preset: "titleAndDate", nodeIds: seq(4),
    });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.headerRoots()[0]);
    expect(xml).toContain("[Document Title]");
    expect(xml).toContain("DATE");
  });

  it("rejects a malformed payload", () => {
    const s = new DocumentSession(makeDoc());
    expect(s.submit({ kind: "insertHeaderFooterPreset", clientId: "a", clientSeq: 1, base: 0, hfKind: "margin" as never, preset: "blank", nodeIds: seq(2) }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertHeaderFooterPreset", clientId: "a", clientSeq: 2, base: 0, hfKind: "header", preset: "fancy" as never, nodeIds: seq(2) }).kind).toBe("rejected");
  });

  it("two sessions converge byte-identically on the whole package", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc());
      s.submit({ kind: "insertHeaderFooterPreset", clientId: "a", clientSeq: 1, base: 0, hfKind: "header", preset: "centeredTitle", nodeIds: seq(2) });
      s.submit({ kind: "insertHeaderFooterPreset", clientId: "a", clientSeq: 2, base: s.seq, hfKind: "footer", preset: "threeColumn", nodeIds: seq(6, 800) });
      return allRootsXml(s.doc);
    };
    expect(build()).toBe(build());
  });
});
