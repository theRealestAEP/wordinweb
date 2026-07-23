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
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("more intents batch 2 (link/footnote/dropcap/divider/bookmark)", () => {
  it("setLink wraps a run in a hyperlink for a safe URL", () => {
    const s = new DocumentSession(makeDoc("click"));
    const e = s.submit({ kind: "setLink", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, url: "https://example.com", nodeIds: [700, 701] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("hyperlink");
  });
  it("setLink REJECTS a javascript: URL (gate 1)", () => {
    const s = new DocumentSession(makeDoc("x"));
    const e = s.submit({ kind: "setLink", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, url: "javascript:alert(1)", nodeIds: [700] });
    expect(e.kind).toBe("rejected");
    expect(serializeXml(s.doc.docRoot)).not.toContain("hyperlink");
  });
  it("insertFootnote adds a footnote reference", () => {
    const s = new DocumentSession(makeDoc("text"));
    const e = s.submit({ kind: "insertFootnote", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, text: "a note", nodeIds: [700, 701] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("footnoteReference");
  });
  it("setDropCap applies a drop cap", () => {
    const s = new DocumentSession(makeDoc("Chapter one"));
    const e = s.submit({ kind: "setDropCap", clientId: "a", clientSeq: 1, base: 0, blockId: ids2(s).blockId, mode: "drop", nodeIds: [700, 701, 702] });
    expect(e.kind).toBe("applied");
  });
  it("setDivider adds a paragraph divider", () => {
    const s = new DocumentSession(makeDoc("para"));
    const e = s.submit({ kind: "setDivider", clientId: "a", clientSeq: 1, base: 0, blockId: ids2(s).blockId, divider: { style: "single", color: "000000", widthPt: 1, spacePt: 1 } });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("pBdr");
  });
  it("insertBookmark adds a bookmark anchor", () => {
    const s = new DocumentSession(makeDoc("anchor"));
    const e = s.submit({ kind: "insertBookmark", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, name: "myMark" });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("bookmarkStart");
  });
  it("determinism across the batch", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      const { blockId, runId } = ids2(s);
      s.submit({ kind: "setLink", clientId: "a", clientSeq: 1, base: 0, runId, url: "https://x.com", nodeIds: [700, 701] });
      s.submit({ kind: "setDivider", clientId: "a", clientSeq: 2, base: s.seq, blockId, divider: { style: "double", color: "FF0000", widthPt: 2, spacePt: 1 } });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
