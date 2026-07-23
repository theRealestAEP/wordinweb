import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run } from "@wordinweb/core";
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
function rid(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("DocumentSession insertBreak", () => {
  it("inserts a page break run at the end of a run", () => {
    const s = new DocumentSession(makeDoc("before"));
    const { runId } = rid(s);
    const e = s.submit({ kind: "insertBreak", clientId: "a", clientSeq: 1, base: 0, runId, breakKind: "page", nodeIds: [400, 401] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain(`w:type="page"`);
  });

  it("a concurrent text edit in the run survives a break insertion (identity transform)", () => {
    const s = new DocumentSession(makeDoc("abc"));
    const { blockId, runId } = rid(s);
    s.submit({ kind: "insertBreak", clientId: "a", clientSeq: 1, base: 0, runId, breakKind: "column", nodeIds: [400, 401] });
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "d" });
    expect(e.kind).toBe("applied");
    const body = s.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const collectT = (el: import("@wordinweb/core").XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += collectT(c))); return t; };
    expect(collectT(body)).toContain("abcd");
  });

  it("rejects a bad break kind", () => {
    const s = new DocumentSession(makeDoc("x"));
    const { runId } = rid(s);
    expect(s.submit({ kind: "insertBreak", clientId: "a", clientSeq: 1, base: 0, runId, breakKind: "evil" as never, nodeIds: [400] }).kind).toBe("rejected");
  });

  it("determinism", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      s.submit({ kind: "insertBreak", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, breakKind: "page", nodeIds: [400, 401] });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
