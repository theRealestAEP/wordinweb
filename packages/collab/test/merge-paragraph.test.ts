import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

function makeDoc(paras: string[]): DocxDocument {
  const body = paras.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join("");
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}
function text(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => { let s = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (s += collectT(c))); return s; };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}
function ids2(s: DocumentSession, idx: number) {
  const para = s.doc.sections[0].blocks[idx] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("DocumentSession mergeParagraph", () => {
  it("merges the second paragraph into the first", () => {
    const s = new DocumentSession(makeDoc(["Hello", "World"]));
    const { blockId } = ids2(s, 1);
    const e = s.submit({ kind: "mergeParagraph", clientId: "a", clientSeq: 1, base: 0, blockId });
    expect(e.kind).toBe("applied");
    expect(text(s.doc)).toBe("HelloWorld");
  });

  it("a concurrent edit in the merged paragraph's run survives (identity transform)", () => {
    const s = new DocumentSession(makeDoc(["AA", "BB"]));
    const p1 = ids2(s, 1); // "BB"
    // B edits inside "BB" (offset 1) and A merges BB into AA, both base 0.
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: p1.blockId, runId: p1.runId, offset: 1 }, text: "X" });
    const e = s.submit({ kind: "mergeParagraph", clientId: "a", clientSeq: 1, base: 0, blockId: p1.blockId });
    expect(e.kind).toBe("applied");
    expect(text(s.doc)).toBe("AABXB"); // merged, with the concurrent insert intact
  });

  it("round-trips split then merge back to the original text", () => {
    const s = new DocumentSession(makeDoc(["HelloWorld"]));
    const p0 = ids2(s, 0);
    s.submit({ kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0, at: { blockId: p0.blockId, runId: p0.runId, offset: 5 }, newBlockId: 900, newRunId: 901 });
    expect(text(s.doc)).toBe("Hello\nWorld");
    s.submit({ kind: "mergeParagraph", clientId: "a", clientSeq: 2, base: s.seq, blockId: 900 });
    expect(text(s.doc)).toBe("HelloWorld");
  });
});
