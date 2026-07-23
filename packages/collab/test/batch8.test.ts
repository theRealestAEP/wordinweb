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
  return localName(el.name) === "drawing" || el.children.some(hasDrawing);
}
function drawingRunId(s: DocumentSession): number {
  let found: number | null = null;
  const walk = (el: XmlElement) => {
    if (localName(el.name) === "r" && hasDrawing(el)) { const id = s.ids.idOf(el); if (id != null) found = id; }
    el.children.forEach(walk);
  };
  s.doc.editableRoots().forEach(walk);
  if (found == null) throw new Error("no drawing run");
  return found;
}
const seq = (n: number, base = 1000) => Array.from({ length: n }, (_, i) => base + i);
function anyPartContains(s: DocumentSession, needle: string): boolean {
  return s.doc.pkg.names().some((n) => (s.doc.pkg.text(n) ?? "").includes(needle));
}
function wordArtDoc(): DocumentSession {
  const s = new DocumentSession(makeDoc("body"));
  s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Art", preset: "plain", nodeIds: seq(12) });
  return s;
}
function smartArtDoc(): DocumentSession {
  const s = new DocumentSession(makeDoc("body"));
  s.submit({ kind: "insertSmartArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), smartArt: { layout: "process", items: ["One", "Two"] }, nodeIds: seq(20) });
  return s;
}

describe("more intents batch 8 (wrap / order / smartart data+fill)", () => {
  it("setImageWrap makes an inline drawing float", () => {
    const s = wordArtDoc();
    const e = s.submit({ kind: "setImageWrap", clientId: "a", clientSeq: 2, base: s.seq, runId: drawingRunId(s), mode: "square" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("anchor");
  });

  it("setDrawingOrder brings a floating drawing to front", () => {
    const s = wordArtDoc();
    // Float it first (setDrawingOrder needs an anchor), then reorder.
    expect(s.submit({ kind: "setImageWrap", clientId: "a", clientSeq: 2, base: s.seq, runId: drawingRunId(s), mode: "square" }).kind).toBe("applied");
    const e = s.submit({ kind: "setDrawingOrder", clientId: "a", clientSeq: 3, base: s.seq, runId: drawingRunId(s), order: "front" });
    expect(e.kind).toBe("applied");
  });

  it("setSmartArtData replaces the whole diagram", () => {
    const s = smartArtDoc();
    const e = s.submit({ kind: "setSmartArtData", clientId: "a", clientSeq: 2, base: s.seq, runId: drawingRunId(s), smartArt: { layout: "cycle", items: ["Alpha", "Beta", "Gamma"] } });
    expect(e.kind).toBe("applied");
    expect(anyPartContains(s, "Gamma")).toBe(true);
  });

  it("setSmartArtFill colors a diagram node", () => {
    const s = smartArtDoc();
    const e = s.submit({ kind: "setSmartArtFill", clientId: "a", clientSeq: 2, base: s.seq, runId: drawingRunId(s), color: "AABBCC", nodeIndex: 0 });
    expect(e.kind).toBe("applied");
    expect(anyPartContains(s, "AABBCC")).toBe(true);
  });

  it("rejects malformed wrap / order / smartart payloads", () => {
    const s = smartArtDoc();
    const rid = drawingRunId(s);
    expect(s.submit({ kind: "setImageWrap", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, mode: "sideways" as never }).kind).toBe("rejected");
    expect(s.submit({ kind: "setDrawingOrder", clientId: "a", clientSeq: 3, base: s.seq, runId: rid, order: "middle" as never }).kind).toBe("rejected");
    expect(s.submit({ kind: "setSmartArtData", clientId: "a", clientSeq: 4, base: s.seq, runId: rid, smartArt: { layout: "spiral" as never, items: ["x"] } }).kind).toBe("rejected");
    expect(s.submit({ kind: "setSmartArtFill", clientId: "a", clientSeq: 5, base: s.seq, runId: rid, color: "xyz" }).kind).toBe("rejected");
  });
});
