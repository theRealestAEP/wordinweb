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
/** Search every package part (chart/diagram data live in separate parts, not
 * document.xml). */
function anyPartContains(s: DocumentSession, needle: string): boolean {
  return s.doc.pkg.names().some((n) => (s.doc.pkg.text(n) ?? "").includes(needle));
}

describe("more intents batch 7 (edit an existing element)", () => {
  it("setChartData replaces an existing chart's data", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertChart", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), chart: { type: "column", categories: ["Q1"], series: [{ name: "A", values: [1] }] }, nodeIds: seq(20) });
    const rid = drawingRunId(s);
    const e = s.submit({ kind: "setChartData", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, chart: { type: "line", categories: ["Q1", "Q2", "Q3"], series: [{ name: "Revenue", values: [10, 20, 30] }] } });
    expect(e.kind).toBe("applied");
    expect(anyPartContains(s, "Revenue")).toBe(true);
  });

  it("setSmartArtNodeText edits a diagram node", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertSmartArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), smartArt: { layout: "process", items: ["One", "Two"] }, nodeIds: seq(20) });
    const rid = drawingRunId(s);
    const e = s.submit({ kind: "setSmartArtNodeText", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, index: 0, text: "Renamed" });
    expect(e.kind).toBe("applied");
    expect(anyPartContains(s, "Renamed")).toBe(true);
  });

  it("setDrawingWordArtText replaces WordArt text", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Before", preset: "wave", nodeIds: seq(12) });
    const rid = drawingRunId(s);
    const e = s.submit({ kind: "setDrawingWordArtText", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, text: "AfterText" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("AfterText");
  });

  it("setDrawingLineStyle sets a drawing outline", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Edge", preset: "plain", nodeIds: seq(12) });
    const rid = drawingRunId(s);
    const e = s.submit({ kind: "setDrawingLineStyle", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, color: "112233", widthPx: 3, dash: "dashed" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("112233");
  });

  it("setImageAltText sets an accessibility description", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Alt", preset: "plain", nodeIds: seq(12) });
    const rid = drawingRunId(s);
    const e = s.submit({ kind: "setImageAltText", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, alt: "A decorative banner" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("A decorative banner");
  });

  it("removeLink unwraps a hyperlink created by setLink", () => {
    const s = new DocumentSession(makeDoc("clickme"));
    const rid = firstRunId(s);
    expect(s.submit({ kind: "setLink", clientId: "a", clientSeq: 1, base: 0, runId: rid, url: "https://example.com", nodeIds: seq(4) }).kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("hyperlink");
    const e = s.submit({ kind: "removeLink", clientId: "a", clientSeq: 2, base: s.seq, runId: rid });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).not.toContain("hyperlink");
  });

  it("rejects malformed edit payloads", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "X", preset: "plain", nodeIds: seq(12) });
    const rid = drawingRunId(s);
    expect(s.submit({ kind: "setChartData", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, chart: { type: "donut" as never, categories: ["x"], series: [{ name: "A", values: [1] }] } }).kind).toBe("rejected");
    expect(s.submit({ kind: "setDrawingLineStyle", clientId: "a", clientSeq: 3, base: s.seq, runId: rid, color: "zzz", widthPx: 3, dash: "dashed" }).kind).toBe("rejected");
    expect(s.submit({ kind: "setDrawingLineStyle", clientId: "a", clientSeq: 4, base: s.seq, runId: rid, color: "112233", widthPx: 999, dash: "dashed" }).kind).toBe("rejected");
    expect(s.submit({ kind: "setSmartArtNodeText", clientId: "a", clientSeq: 5, base: s.seq, runId: rid, index: -1, text: "x" }).kind).toBe("rejected");
  });
});
