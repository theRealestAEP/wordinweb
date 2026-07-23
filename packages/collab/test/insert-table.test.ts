import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

function makeDoc(text: string): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
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

describe("insertTable intent (toolbar table insertion)", () => {
  it("inserts a rows×cols table after the anchor paragraph", () => {
    const s = new DocumentSession(makeDoc("above"));
    const e = s.submit({ kind: "insertTable", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), rows: 2, cols: 3, nodeIds: Array.from({ length: 20 }, (_, i) => 900 + i) });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("<w:tbl");
    expect((xml.match(/<w:tr[ >]/g) ?? []).length).toBe(2);
    expect((xml.match(/<w:tc[ >]/g) ?? []).length).toBe(6);
  });
  it("rejects bad dimensions", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.submit({ kind: "insertTable", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), rows: 0, cols: 3, nodeIds: [] }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertTable", clientId: "a", clientSeq: 2, base: 0, runId: runId(s), rows: 2, cols: 99, nodeIds: [] }).kind).toBe("rejected");
  });
  it("deterministic across replicas", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      s.submit({ kind: "insertTable", clientId: "a", clientSeq: 1, base: 0, runId: runId(s), rows: 2, cols: 2, nodeIds: Array.from({ length: 16 }, (_, i) => 900 + i) });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
