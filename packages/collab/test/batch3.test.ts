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
function ids2(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { runId: s.ids.idOf(run.src!)! };
}

describe("more intents batch 3 (blank page / section break / cross-ref / cover page)", () => {
  it("insertBlankPage inserts a page break structure", () => {
    const s = new DocumentSession(makeDoc("before"));
    const e = s.submit({ kind: "insertBlankPage", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, nodeIds: [700, 701, 702] });
    expect(e.kind).toBe("applied");
  });
  it("insertSectionBreak inserts a section break (multi-paragraph doc)", () => {
    const body = `<w:p><w:r><w:t xml:space="preserve">one</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
    const s = new DocumentSession(DocxDocument.load(zipSync({
      "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
      "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(xml),
    })));
    const para0 = s.doc.sections[0].blocks[0] as Paragraph;
    const runId = s.ids.idOf((para0.children[0] as Run).src!)!;
    const e = s.submit({ kind: "insertSectionBreak", clientId: "a", clientSeq: 1, base: 0, runId, breakType: "nextPage", nodeIds: [700, 701, 702] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("sectPr");
  });
  it("insertCoverPage inserts a cover page with the title", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({ kind: "insertCoverPage", clientId: "a", clientSeq: 1, base: 0, content: { title: "My Report", subtitle: "Q3" }, nodeIds: Array.from({ length: 10 }, (_, i) => 800 + i) });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("My Report");
  });
  it("rejects a bad section break type", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.submit({ kind: "insertSectionBreak", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, breakType: "bad" as never, nodeIds: [700] }).kind).toBe("rejected");
  });
  it("determinism for section break", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      s.submit({ kind: "insertSectionBreak", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, breakType: "continuous", nodeIds: [700, 701, 702] });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
