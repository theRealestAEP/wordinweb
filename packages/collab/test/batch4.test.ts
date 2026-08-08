import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
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
function addr(s: DocumentSession) {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)! };
}

describe("more intents batch 4 (page layout / list level / wordart)", () => {
  it("setPageLayout switches orientation to landscape", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({ kind: "setPageLayout", clientId: "a", clientSeq: 1, base: 0, patch: { orientation: "landscape" } });
    expect(e.kind).toBe("applied");
    expect(serializeXml(s.doc.docRoot)).toContain("landscape");
  });

  it("setPageLayout rejects an unknown key, bad orientation, and out-of-range margin", () => {
    const s = new DocumentSession(makeDoc("x"));
    const bad = (patch: Record<string, unknown>) =>
      s.submit({ kind: "setPageLayout", clientId: "a", clientSeq: 1, base: 0, patch }).kind;
    expect(bad({ evil: "<script>" })).toBe("rejected");
    expect(bad({ orientation: "sideways" })).toBe("rejected");
    expect(bad({ margins: { top: 999 } })).toBe("rejected");
    expect(bad({ columns: 99 })).toBe("rejected");
  });

  it("setListLevel indents a list paragraph's nesting", () => {
    const s = new DocumentSession(makeDoc("item"));
    const a = addr(s);
    // Make the paragraph a numbered list item first, then indent its level.
    expect(s.submit({ kind: "setListType", clientId: "a", clientSeq: 1, base: 0, blockId: a.blockId, listKind: "number" }).kind).toBe("applied");
    const e = s.submit({ kind: "setListLevel", clientId: "a", clientSeq: 2, base: s.seq, blockId: a.blockId, delta: 1 });
    expect(e.kind).toBe("applied");
  });

  it("rejects a bad list-level delta", () => {
    const s = new DocumentSession(makeDoc("x"));
    expect(s.submit({ kind: "setListLevel", clientId: "a", clientSeq: 1, base: 0, blockId: addr(s).blockId, delta: 2 as never }).kind).toBe("rejected");
  });

  it("insertWordArt inserts a decorative text drawing", () => {
    const s = new DocumentSession(makeDoc("body"));
    const e = s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: addr(s).runId, text: "WordyText", preset: "wave", nodeIds: Array.from({ length: 12 }, (_, i) => 900 + i) });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain("WordyText");
    expect(xml).toContain("drawing");
  });

  it("rejects empty wordart text and a bad preset", () => {
    const s = new DocumentSession(makeDoc("x"));
    const rid = addr(s).runId;
    expect(s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: rid, text: "", preset: "wave", nodeIds: [900] }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 2, base: 0, runId: rid, text: "hi", preset: "spiral" as never, nodeIds: [900] }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 3, base: 0, runId: rid, text: "hi", preset: "circle", style: { fill: "not-hex" }, nodeIds: [900] }).kind).toBe("rejected");
    expect(s.submit({ kind: "insertWordArt", clientId: "a", clientSeq: 4, base: 0, runId: rid, text: "hi", preset: "circle", style: { fill: "4472C4", outline: { color: "FFFFFF", widthPt: 900 } }, nodeIds: [900] }).kind).toBe("rejected");
  });

  it("applies a gallery style (fill / outline / shadow) on the wire", () => {
    const s = new DocumentSession(makeDoc("x"));
    const e = s.submit({
      kind: "insertWordArt", clientId: "a", clientSeq: 1, base: 0, runId: addr(s).runId,
      text: "Styled", preset: "button",
      style: { fill: "FFC000", outline: { color: "000000", widthPt: 1 }, shadow: true },
      nodeIds: Array.from({ length: 12 }, (_, i) => 900 + i),
    });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).toContain('prst="textButton"');
    expect(xml).toContain("w14:textOutline");
    expect(xml).toContain("w14:shadow");
    expect(xml).toContain('w14:val="FFC000"');
  });

  it("determinism for page layout", () => {
    const build = () => {
      const s = new DocumentSession(makeDoc("z"));
      s.submit({ kind: "setPageLayout", clientId: "a", clientSeq: 1, base: 0, patch: { orientation: "landscape", columns: 2 } });
      return serializeXml(s.doc.docRoot);
    };
    expect(build()).toBe(build());
  });
});
