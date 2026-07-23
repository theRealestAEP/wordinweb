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
function runId(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}
// 1x1 transparent PNG.
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC";

describe("DocumentSession insertImage", () => {
  it("inserts a drawing with a media part and reserved extents", () => {
    const s = new DocumentSession(makeDoc("photo:"));
    const { runId: r } = runId(s);
    const e = s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, imageBase64: PNG, ext: "png", widthPx: 96, heightPx: 96, nodeIds: [500] });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("w:drawing");
    expect(xml).toContain("wp:extent");
    // The image bytes are in the package as a media part.
    expect([...Object.keys((s.doc as unknown as { pkg: { raw(): Record<string, unknown> } }).pkg.raw())].some((p) => p.includes("media/image"))).toBe(true);
  });

  it("rejects a bad extension", () => {
    const s = new DocumentSession(makeDoc("x"));
    const { runId: r } = runId(s);
    const e = s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, imageBase64: PNG, ext: "png; rm -rf", widthPx: 10, heightPx: 10, nodeIds: [500] });
    expect(e.kind).toBe("rejected");
  });

  it("a concurrent text edit in the run survives an image insertion (identity transform)", () => {
    const s = new DocumentSession(makeDoc("abc"));
    const { blockId, runId: r } = runId(s);
    s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, imageBase64: PNG, ext: "png", widthPx: 20, heightPx: 20, nodeIds: [500] });
    const e = s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId: r, offset: 3 }, text: "d" });
    expect(e.kind).toBe("applied");
    const body = s.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const collectT = (el: import("@wordinweb/core").XmlElement): string => { let t = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (t += collectT(c))); return t; };
    expect(collectT(body)).toContain("abcd");
  });

  it("determinism: two sessions produce identical XML from the same image intent", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("i"));
      const { runId: r } = runId(s);
      s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, imageBase64: PNG, ext: "png", widthPx: 30, heightPx: 30, nodeIds: [600] });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
