import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/**
 * Table geometry intents (user-reported desyncs): column/row drag-resize and
 * table drag-move were LOCAL-ONLY editor mutations. They now ride the wire as
 * resizeTableColumn / resizeTableRow / moveTable, addressed (like tableOp) by
 * an id-tracked paragraph inside the table; all geometric inputs are carried
 * as data so every replica applies the identical mutation.
 */

function makeDoc(): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
  }));
}

/** A session holding a freshly inserted 2x2 table; returns an id-tracked
 * paragraph inside it (the addressing anchor for table intents). */
function tableSession(): { s: DocumentSession; cellParagraphId: number } {
  const s = new DocumentSession(makeDoc());
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  const runId = s.ids.idOf(run.src!)!;
  const nodeIds = Array.from({ length: 20 }, (_, i) => 100 + i);
  const e = s.submit({ kind: "insertTable", clientId: "a", clientSeq: 1, base: 0, runId, rows: 2, cols: 2, nodeIds });
  expect(e.kind).toBe("applied");
  // Find a carried id holding a PARAGRAPH inside the table.
  for (const id of nodeIds) {
    const el = s.ids.elOf(id);
    if (el && el.name.endsWith(":p")) {
      let cur = s.doc.findParentOf(el);
      while (cur && !cur.name.endsWith(":tbl")) cur = s.doc.findParentOf(cur);
      if (cur) return { s, cellParagraphId: id };
    }
  }
  throw new Error("no tracked paragraph inside the inserted table");
}

describe("table geometry intents (resize/move)", () => {
  it("resizeTableColumn changes the grid on every replica identically", () => {
    const a = tableSession();
    const b = tableSession();
    const op = {
      kind: "resizeTableColumn", clientId: "c", clientSeq: 10, base: a.s.seq,
      cellParagraphId: a.cellParagraphId, boundary: 1, deltaPx: 40,
      renderedWidths: [200, 200],
    } as const;
    const ea = a.s.submit(op as never);
    const eb = b.s.submit(op as never);
    expect(ea.kind).toBe("applied");
    expect(eb.kind).toBe("applied");
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
    // The grid actually changed.
    expect(serializeXml(a.s.doc.docRoot)).toContain("tblGrid");
  });

  it("resizeTableRow sets trHeight identically everywhere", () => {
    const a = tableSession();
    const b = tableSession();
    const op = {
      kind: "resizeTableRow", clientId: "c", clientSeq: 10, base: a.s.seq,
      cellParagraphId: a.cellParagraphId, rowIdx: 0, heightPx: 64,
    } as const;
    expect(a.s.submit(op as never).kind).toBe("applied");
    expect(b.s.submit(op as never).kind).toBe("applied");
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
    expect(serializeXml(a.s.doc.docRoot)).toContain("trHeight");
  });

  it("moveTable floats the table at page coordinates identically everywhere", () => {
    const a = tableSession();
    const b = tableSession();
    const op = {
      kind: "moveTable", clientId: "c", clientSeq: 10, base: a.s.seq,
      cellParagraphId: a.cellParagraphId, xPx: 150, yPx: 300, preservePageStart: false, pageDelta: 0,
    } as const;
    expect(a.s.submit(op as never).kind).toBe("applied");
    expect(b.s.submit(op as never).kind).toBe("applied");
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
    expect(serializeXml(a.s.doc.docRoot)).toContain("tblpPr"); // floated
  });

  it("rejects cleanly when the paragraph is not inside a table", () => {
    const s = new DocumentSession(makeDoc());
    const para = s.doc.sections[0].blocks[0] as Paragraph;
    const blockId = s.ids.idOf(para.src!)!;
    const e = s.submit({
      kind: "resizeTableColumn", clientId: "a", clientSeq: 1, base: 0,
      cellParagraphId: blockId, boundary: 1, deltaPx: 40,
    } as never);
    expect(e.kind).toBe("rejected");
  });
});
