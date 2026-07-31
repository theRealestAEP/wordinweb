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
function addr(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}
function docText(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += collectT(c))); return t; };
  return collectT(body);
}

describe("DocumentSession tracked changes (suggesting mode)", () => {
  it("records a suggested insertion as a w:ins with the author + date", () => {
    const s = new DocumentSession(makeDoc("Hello"));
    const { blockId, runId } = addr(s);
    const e = s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 5 }, text: " world", suggest: { author: "Reviewer", date: "2026-07-22T12:00:00.000Z" } });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Reviewer");
    expect(docText(s.doc)).toBe("Hello world"); // the text is present (tracked)
  });

  it("determinism: the same suggested edit produces identical XML on two replicas", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("abc"));
      const { blockId, runId } = addr(s);
      s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "def", suggest: { author: "R", date: "2026-07-22T12:00:00.000Z" } });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });

  it("a plain insert (no suggest) is NOT tracked", () => {
    const s = new DocumentSession(makeDoc("x"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 1 }, text: "y" });
    expect(serializeXml(s.doc.docRoot)).not.toContain("w:ins");
  });
});
