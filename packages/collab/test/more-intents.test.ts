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

describe("more intents (indent / spacing / page field)", () => {
  it("adjustIndent indents a paragraph", () => {
    const s = new DocumentSession(makeDoc("para"));
    const e = s.submit({ kind: "adjustIndent", clientId: "a", clientSeq: 1, base: 0, blockId: ids2(s).blockId, direction: 1 });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("w:ind");
  });
  it("setSpacing sets paragraph spacing", () => {
    const s = new DocumentSession(makeDoc("para"));
    const e = s.submit({ kind: "setSpacing", clientId: "a", clientSeq: 1, base: 0, blockId: ids2(s).blockId, patch: { before: 240, after: 240 } });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("w:spacing");
  });
  it("insertPageField inserts a PAGE field", () => {
    const s = new DocumentSession(makeDoc("page "));
    const e = s.submit({ kind: "insertPageField", clientId: "a", clientSeq: 1, base: 0, runId: ids2(s).runId, fieldKind: "page", nodeIds: [700, 701] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toMatch(/PAGE|fldChar|instrText/);
  });
  it("a concurrent text edit survives an indent (identity transform)", () => {
    const s = new DocumentSession(makeDoc("abc"));
    const { blockId, runId } = ids2(s);
    s.submit({ kind: "adjustIndent", clientId: "a", clientSeq: 1, base: 0, blockId, direction: 1 });
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "d" });
    expect(e.kind).toBe("applied");
  });
  it("determinism for all three", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      const { blockId, runId } = ids2(s);
      s.submit({ kind: "adjustIndent", clientId: "a", clientSeq: 1, base: 0, blockId, direction: 1 });
      s.submit({ kind: "setSpacing", clientId: "a", clientSeq: 2, base: s.seq, blockId, patch: { after: 120 } });
      s.submit({ kind: "insertPageField", clientId: "a", clientSeq: 3, base: s.seq, runId, fieldKind: "page", nodeIds: [700, 701] });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
