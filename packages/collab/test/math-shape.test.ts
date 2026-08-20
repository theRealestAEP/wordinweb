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

describe("DocumentSession insertMath", () => {
  it("inserts a math formula (oMath) from a linear expression", () => {
    const s = new DocumentSession(makeDoc("eq: "));
    const e = s.submit({ kind: "insertMath", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, mathText: "a+b", nodeIds: [400, 401] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toMatch(/oMath|m:/);
  });
  it("rejects empty math", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.submit({ kind: "insertMath", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, mathText: "", nodeIds: [400] }).kind).toBe("rejected");
  });
  it("determinism", () => {
    const build = () => { const s = new DocumentSession(makeDoc("m")); s.submit({ kind: "insertMath", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, mathText: "x^2", nodeIds: [400, 401] }); return serializeXml(s.doc.docRoot); };
    expect(build()).toBe(build());
  });
});

describe("DocumentSession insertShape", () => {
  it("inserts a rectangle shape drawing", () => {
    const s = new DocumentSession(makeDoc("shape: "));
    const e = s.submit({ kind: "insertShape", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, preset: "rectangle", nodeIds: [500, 501] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("w:drawing");
  });
  it("rejects a bad preset", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.submit({ kind: "insertShape", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, preset: "evil" as never, nodeIds: [500] }).kind).toBe("rejected");
  });
  it("accepts gallery presets from the preset-geometry table (e18)", () => {
    const s = new DocumentSession(makeDoc("shape: "));
    const e = s.submit({ kind: "insertShape", clientId: "a", clientSeq: 1, base: 0, runId: rid(s).runId, preset: "star5", nodeIds: [500, 501] });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain('prst="star5"');
  });
  it("a concurrent text edit survives a shape insertion (identity transform)", () => {
    const s = new DocumentSession(makeDoc("abc"));
    const { blockId, runId } = rid(s);
    s.submit({ kind: "insertShape", clientId: "a", clientSeq: 1, base: 0, runId, preset: "ellipse", nodeIds: [500, 501] });
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 3 }, text: "d" });
    expect(e.kind).toBe("applied");
    const body = s.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const collectT = (el: import("@wordinweb/core").XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += collectT(c))); return t; };
    expect(collectT(body)).toContain("abcd");
  });
});
