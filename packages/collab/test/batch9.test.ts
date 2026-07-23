import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

function makeDoc(text: string): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}
function firstRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}
function hasLocal(el: XmlElement, ln: string): boolean {
  return localName(el.name) === ln || el.children.some((c) => hasLocal(c, ln));
}
function runIdWith(s: DocumentSession, ln: string): number {
  let found: number | null = null;
  const walk = (el: XmlElement) => {
    if (localName(el.name) === "r" && hasLocal(el, ln)) { const id = s.ids.idOf(el); if (id != null) found = id; }
    el.children.forEach(walk);
  };
  s.doc.editableRoots().forEach(walk);
  if (found == null) throw new Error(`no run with ${ln}`);
  return found;
}
const seq = (n: number, base = 1000) => Array.from({ length: n }, (_, i) => base + i);
function anyPartContains(s: DocumentSession, needle: string): boolean {
  return s.doc.pkg.names().some((n) => (s.doc.pkg.text(n) ?? "").includes(needle));
}
function smartArtDoc(): DocumentSession {
  const s = new DocumentSession(makeDoc("body"));
  s.submit({ kind: "insertSmartArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), smartArt: { layout: "process", items: ["One", "Two"] }, nodeIds: seq(20) });
  return s;
}
function mathDoc(): DocumentSession {
  const s = new DocumentSession(makeDoc("body"));
  s.submit({ kind: "insertMath", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), mathText: "a+b", nodeIds: seq(20) });
  return s;
}
/** Stable id of the paragraph (block) that contains an m:oMath. */
function mathBlockId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf(para.src!)!;
}

describe("more intents batch 9 (smartart format / floating / math / comment / bookmark / checkbox)", () => {
  it("setSmartArtTextFormat formats a diagram node", () => {
    const s = smartArtDoc();
    const e = s.submit({ kind: "setSmartArtTextFormat", clientId: "a", clientSeq: 2, base: s.seq, runId: runIdWith(s, "drawing"), format: { fontFamily: "Arial", fontSizePt: 18, color: "223344", bold: true, italic: false, alignment: "center" }, nodeIndex: 0 });
    expect(e.kind).toBe("applied");
    expect(anyPartContains(s, "223344")).toBe(true);
  });

  it("setFloatingPagePosition positions a floating drawing", () => {
    const s = new DocumentSession(makeDoc("body"));
    s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Float", preset: "plain", nodeIds: seq(12) });
    const rid = runIdWith(s, "drawing");
    expect(s.submit({ kind: "setImageWrap", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, mode: "square" }).kind).toBe("applied");
    const e = s.submit({ kind: "setFloatingPagePosition", clientId: "a", clientSeq: 3, base: s.seq, runId: runIdWith(s, "drawing"), xPx: 100, yPx: 200 });
    expect(e.kind).toBe("applied");
  });

  it("setMathLinear replaces an equation", () => {
    const s = mathDoc();
    const e = s.submit({ kind: "setMathLinear", clientId: "a", clientSeq: 2, base: s.seq, blockId: mathBlockId(s), mathText: "x^2+y^2" });
    expect(e.kind).toBe("applied");
  });

  it("deleteMath removes a math object", () => {
    const s = mathDoc();
    const e = s.submit({ kind: "deleteMath", clientId: "a", clientSeq: 2, base: s.seq, blockId: mathBlockId(s) });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).not.toContain("oMath");
  });

  it("deleteComment removes a comment thread", () => {
    const s = new DocumentSession(makeDoc("commented"));
    expect(s.submit({ kind: "commentRun", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), text: "Note", author: "Reviewer", date: "2020-01-01T00:00:00Z", paraId: "11111111" }).kind).toBe("applied");
    expect(s.doc.comments.length).toBe(1);
    const id = s.doc.comments[0].id;
    const e = s.submit({ kind: "deleteComment", clientId: "a", clientSeq: 2, base: s.seq, commentId: id });
    expect(e.kind).toBe("applied");
    expect(s.doc.comments.length).toBe(0);
  });

  it("insertBookmarkRange wraps a sub-range in a bookmark", () => {
    const s = new DocumentSession(makeDoc("HelloWorld"));
    const e = s.submit({ kind: "insertBookmarkRange", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s), name: "Anchor1", start: 0, end: 5 });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("bookmarkStart");
    expect(xml).toContain("Anchor1");
  });

  it("toggleCheckbox on a run without a checkbox is a clean no-op", () => {
    const s = new DocumentSession(makeDoc("plain"));
    expect(s.submit({ kind: "toggleCheckbox", clientId: "a", clientSeq: 1, base: 0, runId: firstRunId(s) }).kind).toBe("rejected");
  });

  it("rejects malformed batch-9 payloads", () => {
    const s = smartArtDoc();
    const rid = runIdWith(s, "drawing");
    expect(s.submit({ kind: "setSmartArtTextFormat", clientId: "a", clientSeq: 2, base: s.seq, runId: rid, format: { fontFamily: "Arial", fontSizePt: 18, color: "zzz", bold: true, italic: false, alignment: "center" } }).kind).toBe("rejected");
    expect(s.submit({ kind: "setFloatingPagePosition", clientId: "a", clientSeq: 3, base: s.seq, runId: rid, xPx: 99999, yPx: 0 }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertBookmarkRange", clientId: "a", clientSeq: 4, base: s.seq, runId: firstRunId(s), name: "1bad name", start: 0, end: 2 }).kind).toBe("rejected");
    expect(s.submit({ kind: "deleteComment", clientId: "a", clientSeq: 5, base: s.seq, commentId: "" }).kind).toBe("rejected");
  });
});
