import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, collectRevisions, type Paragraph, type Run } from "@wordinweb/core";
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
  return { blockId: s.ids.idOf(para.src!)!, runId: s.ids.idOf(run.src!)!, len: run.content.find((c) => c.kind === "text")!.srcT!.text.length };
}
/** A doc with one tracked (suggested) insertion at the end of the run. */
function suggestedDoc(): DocumentSession {
  const s = new DocumentSession(makeDoc("Base"));
  const a = addr(s);
  s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: a.blockId, runId: a.runId, offset: a.len }, text: "NEW", suggest: { author: "Reviewer", date: "2020-01-01T00:00:00Z" } });
  return s;
}

describe("more intents batch 10 (tracked-change review — reaches 60 intents)", () => {
  it("suggested insert shows as a tracked revision", () => {
    const s = suggestedDoc();
    expect(collectRevisions(s.doc).length).toBeGreaterThan(0);
    expect(serializeXml(s.doc.docRoot)).toContain("w:ins");
  });

  it("acceptRevision keeps a suggested insertion's text", () => {
    const s = suggestedDoc();
    const e = s.submit({ kind: "acceptRevision", clientId: "a", clientSeq: 2, base: s.seq, index: 0 });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).not.toContain("w:ins"); // no longer a pending change
    expect(xml).toContain("NEW"); // accepted -> text stays
  });

  it("rejectRevision removes a suggested insertion's text", () => {
    const s = suggestedDoc();
    const e = s.submit({ kind: "rejectRevision", clientId: "a", clientSeq: 2, base: s.seq, index: 0 });
    expect(e.kind).toBe("applied");
    const xml = serializeXml(s.doc.docRoot);
    expect(xml).not.toContain("w:ins");
    expect(xml).not.toContain("NEW"); // rejected -> text gone
  });

  it("acceptAllRevisions clears every pending change", () => {
    const s = suggestedDoc();
    // A second suggested insertion, so there are multiple to clear at once.
    const a = addr(s);
    s.submit({ kind: "insertText", clientId: "a", clientSeq: 2, base: s.seq, at: { blockId: a.blockId, runId: a.runId, offset: 0 }, text: "X", suggest: { author: "R", date: "2020-01-02T00:00:00Z" } });
    expect(collectRevisions(s.doc).length).toBeGreaterThan(0);
    const e = s.submit({ kind: "acceptAllRevisions", clientId: "a", clientSeq: 3, base: s.seq });
    expect(e.kind).toBe("applied");
    expect(collectRevisions(s.doc).length).toBe(0);
  });

  it("rejectAllRevisions clears every pending change", () => {
    const s = suggestedDoc();
    const e = s.submit({ kind: "rejectAllRevisions", clientId: "a", clientSeq: 2, base: s.seq });
    expect(e.kind).toBe("applied");
    expect(collectRevisions(s.doc)).toHaveLength(0);
    expect(serializeXml(s.doc.docRoot)).not.toContain("NEW");
  });

  it("acceptRevision with an out-of-range index is a clean no-op", () => {
    const s = suggestedDoc();
    expect(s.submit({ kind: "acceptRevision", clientId: "a", clientSeq: 2, base: s.seq, index: 99 }).kind).toBe("rejected");
  });

  it("rejects a negative revision index", () => {
    const s = suggestedDoc();
    expect(s.submit({ kind: "rejectRevision", clientId: "a", clientSeq: 2, base: s.seq, index: -1 }).kind).toBe("rejected");
  });

  it("acceptAllRevisions on a doc with no changes is a clean no-op", () => {
    const s = new DocumentSession(makeDoc("Clean"));
    expect(s.submit({ kind: "acceptAllRevisions", clientId: "a", clientSeq: 1, base: 0 }).kind).toBe("rejected");
  });
});
