import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
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
function paras(doc: DocxDocument): string[] {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => { let s = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (s += collectT(c))); return s; };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT);
}
function firstBlockId(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf(para.src!)!;
}

describe("DocumentSession pasteBlocks (rich paste + gate 2)", () => {
  it("splices validated OOXML paragraphs after the target, with carried ids", () => {
    const s = new DocumentSession(makeDoc("start"));
    const afterBlockId = firstBlockId(s);
    const blocksXml =
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">one</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>`;
    const e = s.submit({ kind: "pasteBlocks", clientId: "a", clientSeq: 1, base: 0, afterBlockId, blocksXml, nodeIds: [700, 701, 702, 703] });
    expect(e.kind).toBe("applied");
    expect(paras(s.doc)).toEqual(["start", "one", "two"]);
    expect(serializeXml(s.doc.docRoot)).toContain("<w:b/>");
    // The pasted paragraphs are addressable by carried ids for a follow-up edit.
    const ins = s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { blockId: 700, runId: 701, offset: 3 }, text: "!" });
    expect(ins.kind).toBe("applied");
    expect(paras(s.doc)).toEqual(["start", "one!", "two"]);
  });

  it("REJECTS a paste containing a forbidden element (hyperlink) — gate 2", () => {
    const s = new DocumentSession(makeDoc("x"));
    const afterBlockId = firstBlockId(s);
    const blocksXml = `<w:p><w:hyperlink r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>evil</w:t></w:r></w:hyperlink></w:p>`;
    const e = s.submit({ kind: "pasteBlocks", clientId: "a", clientSeq: 1, base: 0, afterBlockId, blocksXml, nodeIds: [800, 801] });
    expect(e.kind).toBe("rejected");
    expect(paras(s.doc)).toEqual(["x"]); // nothing pasted
  });

  it("REJECTS a paste with a relationship reference (external target)", () => {
    const s = new DocumentSession(makeDoc("x"));
    const afterBlockId = firstBlockId(s);
    const blocksXml = `<w:p><w:r r:embed="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:t>x</w:t></w:r></w:p>`;
    expect(s.submit({ kind: "pasteBlocks", clientId: "a", clientSeq: 1, base: 0, afterBlockId, blocksXml, nodeIds: [900] }).kind).toBe("rejected");
  });

  it("determinism: two sessions produce identical XML from the same paste", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("s"));
      s.submit({ kind: "pasteBlocks", clientId: "a", clientSeq: 1, base: 0, afterBlockId: firstBlockId(s), blocksXml: `<w:p><w:r><w:t xml:space="preserve">p</w:t></w:r></w:p>`, nodeIds: [1000, 1001] });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
