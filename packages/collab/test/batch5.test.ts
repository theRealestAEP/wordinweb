import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
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
function runId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}
const ids = (n: number, base = 1000) => Array.from({ length: n }, (_, i) => base + i);

describe("more intents batch 5 (chart / smartart / line numbering / fields)", () => {
  it("insertChart inserts a column chart", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({
      kind: "insertChart", clientId: "a", clientSeq: 1, base: 0, runId: runId(s),
      chart: { type: "column", title: "Sales", categories: ["Q1", "Q2"], series: [{ name: "A", values: [3, 5] }] },
      nodeIds: ids(20),
    });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("drawing");
  });

  it("insertSmartArt inserts a process diagram", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({
      kind: "insertSmartArt", clientId: "a", clientSeq: 1, base: 0, runId: runId(s),
      smartArt: { layout: "process", items: ["Plan", "Build", "Ship"] }, nodeIds: ids(20),
    });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("drawing");
  });

  it("setLineNumbering enables margin line numbers", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({ kind: "setLineNumbering", clientId: "a", clientSeq: 1, base: 0, patch: { enabled: true, countBy: 5, restart: "newPage" } });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("lnNumType");
  });

  it("insertDateTimeField inserts a DATE field", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({ kind: "insertDateTimeField", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), dtKind: "date", picture: "MMMM d, yyyy", nodeIds: ids(6) });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("DATE");
  });

  it("insertField inserts an allowlisted PAGE field", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({ kind: "insertField", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), instruction: "PAGE", nodeIds: ids(6) });
    expect(e.kind).toBe("applied");
  });

  it("insertField REJECTS an external-content / code field", () => {
    const s = new DocumentSession(makeDoc("body"));
    const rid = runId(s);
    const bad = (instruction: string, seq: number) =>
      s.submit({ kind: "insertField", clientId: "a", clientSeq: seq, base: 0, runId: rid, instruction, nodeIds: ids(6) }).kind;
    expect(bad('INCLUDETEXT "http://evil/x.docx"', 1)).toBe("rejected");
    expect(bad("DDEAUTO Excel Sheet1 R1C1", 2)).toBe("rejected");
    expect(bad('LINK Word.Document.12 "c:\\\\x"', 3)).toBe("rejected");
    expect(bad('HYPERLINK "javascript:alert(1)"', 4)).toBe("rejected");
  });

  it("rejects malformed chart / smartart / line-numbering payloads", () => {
    const s = new DocumentSession(makeDoc("body"));
    const rid = runId(s);
    expect(s.submit({ kind: "insertChart", clientId: "a", clientSeq: 1, base: 0, runId: rid, chart: { type: "donut" as never, categories: ["x"], series: [{ name: "A", values: [1] }] }, nodeIds: ids(4) }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertChart", clientId: "a", clientSeq: 2, base: 0, runId: rid, chart: { type: "pie", categories: [], series: [{ name: "A", values: [1] }] }, nodeIds: ids(4) }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertChart", clientId: "a", clientSeq: 3, base: 0, runId: rid, chart: { type: "column", categories: ["x"], series: [{ name: "A", values: [1] }], grouping: "sideways" as never }, nodeIds: ids(4) }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertChart", clientId: "a", clientSeq: 4, base: 0, runId: rid, chart: { type: "doughnut", categories: ["x"], series: [{ name: "A", values: [1] }] }, nodeIds: ids(20, 2000) }).kind).toBe("applied");
    expect(s.submit({ kind: "insertChart", clientId: "a", clientSeq: 5, base: 0, runId: rid, chart: { type: "column", categories: ["x"], series: [{ name: "A", values: [1] }], grouping: "stacked" }, nodeIds: ids(20, 2100) }).kind).toBe("applied");
    expect(s.submit({ kind: "insertSmartArt", clientId: "a", clientSeq: 6, base: 0, runId: rid, smartArt: { layout: "spiral" as never, items: ["a"] }, nodeIds: ids(4) }).kind).toBe("rejected");
    expect(s.submit({ kind: "setLineNumbering", clientId: "a", clientSeq: 7, base: 0, patch: { enabled: true, countBy: 9999 } }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertDateTimeField", clientId: "a", clientSeq: 8, base: 0, runId: rid, dtKind: "date", picture: 'bad"{char}', nodeIds: ids(4) }).kind).toBe("rejected");
  });

  it("determinism for chart insertion", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      s.submit({ kind: "insertChart", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), chart: { type: "bar", categories: ["a", "b"], series: [{ name: "S", values: [1, 2] }] }, nodeIds: ids(20) });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
