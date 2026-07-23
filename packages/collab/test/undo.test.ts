import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
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
function text(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => { let s = localName(el.name) === "t" ? el.text : ""; el.children.forEach((c) => (s += collectT(c))); return s; };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}
function addr(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("DocumentSession collaborative undo", () => {
  it("undoes an insert, restoring the prior text", () => {
    const s = new DocumentSession(makeDoc("ab"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 2 }, text: "CD" });
    expect(text(s.doc)).toBe("abCD");
    expect(s.undoDepth("a")).toBe(1);
    const u = s.undo("a");
    expect(u?.kind).toBe("applied");
    expect(text(s.doc)).toBe("ab"); // insert undone
    expect(s.undoDepth("a")).toBe(0);
  });

  it("undoes a delete, re-inserting the removed text", () => {
    const s = new DocumentSession(makeDoc("hello"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "deleteText", clientId: "a", clientSeq: 1, base: 0, blockId, runId, start: 1, end: 3 }); // remove "el"
    expect(text(s.doc)).toBe("hlo");
    s.undo("a");
    expect(text(s.doc)).toBe("hello"); // re-inserted
  });

  it("undo is REBASED under concurrency: another client's edit before the undone insert shifts the undo correctly", () => {
    const s = new DocumentSession(makeDoc("XY"));
    const { blockId, runId } = addr(s);
    // A inserts "AA" at offset 2 (end) → "XYAA" (seq 1).
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 2 }, text: "AA" });
    // B inserts "BB" at offset 0 (start) → "BBXYAA" (seq 2).
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId, runId, offset: 0 }, text: "BB" });
    expect(text(s.doc)).toBe("BBXYAA");
    // A undoes its insert. The naive inverse would delete [2,4) = "XY" (WRONG).
    // Rebased through B's insert, it must delete the shifted "AA" at [4,6).
    s.undo("a");
    expect(text(s.doc)).toBe("BBXY"); // only A's "AA" removed; B's "BB" intact
  });

  it("selective per-user undo: A's undo reverts A's edit, not B's", () => {
    const s = new DocumentSession(makeDoc("-"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 1 }, text: "A" }); // "-A"
    s.submit({ kind: "insertText", clientId: "b", clientSeq: 1, base: s.seq, at: { blockId, runId, offset: 2 }, text: "B" }); // "-AB"
    expect(text(s.doc)).toBe("-AB");
    s.undo("a"); // removes A's "A", rebased past B's "B"
    expect(text(s.doc)).toBe("-B"); // A's edit gone, B's kept
    expect(s.undoDepth("b")).toBe(1); // B still has an undoable edit
  });

  it("undoes a paragraph split by merging back", () => {
    const s = new DocumentSession(makeDoc("HelloWorld"));
    const { blockId, runId } = addr(s);
    s.submit({ kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0, at: { blockId, runId, offset: 5 }, newBlockId: 800, newRunId: 801 });
    expect(text(s.doc)).toBe("Hello\nWorld");
    s.undo("a");
    expect(text(s.doc)).toBe("HelloWorld"); // merged back
  });

  it("returns null when there is nothing to undo", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.undo("nobody")).toBeNull();
  });
});
