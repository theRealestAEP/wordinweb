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
    const e = s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, blobSha: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", bytesLen: 68, ext: "png", widthPx: 96, heightPx: 96, nodeIds: [500] });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("w:drawing");
    expect(xml).toContain("wp:extent");
    // The image bytes are in the package as a media part.
    // v2 (doc 16 §2): the part is REGISTERED but PENDING — no zip entry
    // (the hole is the unambiguous not-yet state), sha recorded for the
    // verification chain, extents reserved so layout is byte-identical.
    const pending = [...s.doc.pendingMedia.entries()];
    expect(pending).toHaveLength(1);
    expect(pending[0][0]).toContain("media/image");
    expect(pending[0][1].sha).toBe("a".repeat(64));
    expect(s.doc.mediaStatus(pending[0][0])).toBe("pending");
    // installMedia completes the round trip (caller verified the sha).
    expect(s.doc.installMedia(pending[0][0], new Uint8Array([1, 2, 3]))).toBe(true);
    expect(s.doc.mediaStatus(pending[0][0])).toBe("ready");
  });

  it("rejects a bad extension", () => {
    const s = new DocumentSession(makeDoc("x"));
    const { runId: r } = runId(s);
    const e = s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, blobSha: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", bytesLen: 68, ext: "png; rm -rf", widthPx: 10, heightPx: 10, nodeIds: [500] });
    expect(e.kind).toBe("rejected");
  });

  it("a concurrent text edit in the run survives an image insertion (identity transform)", () => {
    const s = new DocumentSession(makeDoc("abc"));
    const { blockId, runId: r } = runId(s);
    s.submit({ kind: "insertImage", clientId: "a", clientSeq: 1, base: 0, runId: r, blobSha: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", bytesLen: 68, ext: "png", widthPx: 20, heightPx: 20, nodeIds: [500] });
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
