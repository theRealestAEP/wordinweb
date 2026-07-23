import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

function makeDoc(text: string): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}
function firstRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}
function hasDrawing(el: XmlElement): boolean {
  if (localName(el.name) === "drawing") return true;
  return el.children.some(hasDrawing);
}
/** Stable id of the run element (w:r) that now carries a w:drawing. */
function drawingRunId(s: DocumentSession): number | null {
  let found: number | null = null;
  const walk = (el: XmlElement) => {
    if (localName(el.name) === "r" && hasDrawing(el)) { const id = s.ids.idOf(el); if (id != null) found = id; }
    el.children.forEach(walk);
  };
  s.doc.editableRoots().forEach(walk);
  return found;
}

describe("more intents batch 6 (drawing-property edits)", () => {
  it("setDrawingRotation rotates a run's drawing", () => {
    const s = new DocumentSession(makeDoc("body"));
    expect(s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Spin", preset: "plain", nodeIds: Array.from({ length: 12 }, (_, i) => 900 + i) }).kind).toBe("applied");
    const rid = drawingRunId(s);
    expect(rid).not.toBeNull();
    const e = s.submit({ kind: "setDrawingRotation", clientId: "a", clientSeq: 2, base: s.seq, runId: rid!, degrees: 45 });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain(String(Math.round(45 * 60000)));
  });

  it("setDrawingFill sets a solid fill on a run's drawing", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Fill", preset: "plain", nodeIds: Array.from({ length: 12 }, (_, i) => 900 + i) });
    const rid = drawingRunId(s)!;
    const e = s.submit({ kind: "setDrawingFill", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, color: "FF8800" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("FF8800");
  });

  it("rejects a bad rotation and a bad fill color", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "X", preset: "plain", nodeIds: Array.from({ length: 12 }, (_, i) => 900 + i) });
    const rid = drawingRunId(s)!;
    expect(s.submit({ kind: "setDrawingRotation", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, degrees: Infinity as never }).kind).toBe("rejected");
    expect(s.submit({ kind: "setDrawingFill", clientId: "a", clientSeq: 3, base: s.seq, runId: rid, color: "not-hex" }).kind).toBe("rejected");
  });

  it("a drawing edit on a run with no drawing is a clean no-op (rejected)", () => {
    const s = new DocumentSession(makeDoc("plain"));
    expect(s.submit({ kind: "setDrawingRotation", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), degrees: 30 }).kind).toBe("rejected");
  });
});
