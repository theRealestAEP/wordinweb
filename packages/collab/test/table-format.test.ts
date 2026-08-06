import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, collectRevisions, serializeXml, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/**
 * Table FORMATTING intents: per-edge borders, style application, numeric
 * widths, cell margins and the repeating header band. All eight come from the
 * core operation registry, and all eight are cell-addressed — the wire carries
 * the stable id of a paragraph inside the table.
 *
 * The address is deliberately lossy for some of them: a table-scoped operation
 * widens it to the owning w:tbl, while a cell-scoped one has to land on the
 * one cell the id names. Getting that wrong is invisible in a 1x1 table, so
 * every cell-scoped case here addresses a cell that is neither the first nor
 * the last.
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

function localPart(name: string): string {
  return name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
}

function kids(el: XmlElement, name: string): XmlElement[] {
  return el.children.filter((c) => localPart(c.name) === name);
}

/** A session holding a freshly inserted 3x3 table, plus a way to name any of
 * its cells by the stable id of the paragraph inside it. */
function tableSession(): {
  s: DocumentSession;
  tbl: XmlElement;
  cellParagraphId: (row: number, col: number) => number;
} {
  const s = new DocumentSession(makeDoc());
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const runId = s.ids.idOf((para.children[0] as Run).src!)!;
  const nodeIds = Array.from({ length: 30 }, (_, i) => 100 + i);
  expect(
    s.submit({ kind: "insertTable", clientId: "a", clientSeq: 1, base: 0, runId, rows: 3, cols: 3, nodeIds }).kind,
  ).toBe("applied");

  let tbl: XmlElement | null = null;
  const walk = (el: XmlElement): void => {
    if (tbl) return;
    if (localPart(el.name) === "tbl") tbl = el;
    else el.children.forEach(walk);
  };
  walk(s.doc.docRoot);
  if (!tbl) throw new Error("no table");

  const cellParagraphId = (row: number, col: number): number => {
    const tc = kids(kids(tbl!, "tr")[row], "tc")[col];
    const paraEl = kids(tc, "p")[0];
    const id = s.ids.idOf(paraEl);
    if (id === undefined) throw new Error(`cell ${row},${col} paragraph is not id-tracked`);
    return id;
  };
  return { s, tbl, cellParagraphId };
}

/** Submit the same body to two independent replicas and require they agree. */
function bothApply(body: Record<string, unknown>): { a: ReturnType<typeof tableSession>; xml: string } {
  const a = tableSession();
  const b = tableSession();
  const op = { clientId: "c", clientSeq: 10, base: a.s.seq, ...body };
  expect(a.s.submit(op as never).kind).toBe("applied");
  expect(b.s.submit({ ...op, base: b.s.seq } as never).kind).toBe("applied");
  const xml = serializeXml(a.s.doc.docRoot);
  expect(xml).toBe(serializeXml(b.s.doc.docRoot));
  return { a, xml };
}

describe("table formatting intents", () => {
  it("puts cell borders on the addressed cell and nowhere else", () => {
    const a = tableSession();
    const { s, cellParagraphId } = a;
    expect(
      s.submit({
        kind: "setTableBorders", clientId: "c", clientSeq: 10, base: s.seq,
        cellParagraphId: cellParagraphId(1, 1),
        scope: "cell", edges: ["top", "bottom"],
        border: { style: "double", sz: 12, color: "#FF0000" },
      } as never).kind,
    ).toBe("applied");

    const cellHasBorders = (row: number, col: number): boolean => {
      const tc = kids(kids(a.tbl, "tr")[row], "tc")[col];
      const tcPr = kids(tc, "tcPr")[0];
      return !!tcPr && kids(tcPr, "tcBorders").length > 0;
    };
    expect(cellHasBorders(1, 1)).toBe(true);
    for (const [r, c] of [[0, 0], [0, 1], [1, 0], [1, 2], [2, 2]]) {
      expect(cellHasBorders(r, c), `cell ${r},${c}`).toBe(false);
    }
  });

  it("widens a table-scoped border to the whole table from any cell", () => {
    // Addressing the middle cell must still write tblBorders, not tcBorders.
    const a = tableSession();
    expect(
      a.s.submit({
        kind: "setTableBorders", clientId: "c", clientSeq: 10, base: a.s.seq,
        cellParagraphId: a.cellParagraphId(1, 1),
        scope: "table", edges: ["insideH", "insideV"], border: { style: "dashed", sz: 6 },
      } as never).kind,
    ).toBe("applied");
    expect(kids(kids(a.tbl, "tblPr")[0], "tblBorders").length).toBe(1);
    const middle = kids(kids(a.tbl, "tr")[1], "tc")[1];
    expect(kids(middle, "tcPr").some((pr) => kids(pr, "tcBorders").length > 0)).toBe(false);
  });

  it("applies numeric widths and the layout switch identically everywhere", () => {
    const seed = tableSession();
    const addr = seed.cellParagraphId(2, 0);
    const { xml: widthXml } = bothApply({ kind: "setTableWidth", cellParagraphId: addr, unit: "pct", value: 75 });
    expect(widthXml).toContain(`w:type="pct"`);
    const { xml: colXml } = bothApply({ kind: "setTableColumnWidth", cellParagraphId: addr, colIdx: 1, widthPt: 120 });
    expect(colXml).toContain(`w:w="2400"`);
    const { xml: layoutXml } = bothApply({
      kind: "setTableLayout", cellParagraphId: addr, layout: "fixed", renderedWidths: [120, 160, 200],
    });
    expect(layoutXml).toContain(`w:tblLayout w:type="fixed"`);
  });

  it("applies margins, header rows and the look identically everywhere", () => {
    const addr = tableSession().cellParagraphId(1, 1);
    expect(bothApply({ kind: "setTableCellMargins", cellParagraphId: addr, scope: "table", margins: { left: 9, right: 9 } }).xml)
      .toContain("tblCellMar");
    expect(bothApply({ kind: "setTableHeaderRows", cellParagraphId: addr, count: 2 }).xml).toContain("tblHeader");
    expect(bothApply({ kind: "setTableLook", cellParagraphId: addr, look: { bandedRows: false } }).xml)
      .toContain(`w:noHBand="1"`);
  });

  it("rejects a cell address that names nothing on this replica", () => {
    const { s } = tableSession();
    expect(
      s.submit({
        kind: "setTableBorders", clientId: "c", clientSeq: 10, base: s.seq,
        cellParagraphId: 999999, scope: "cell", edges: ["top"], border: { style: "single" },
      } as never).kind,
    ).toBe("rejected");
  });

  it("rejects a style the document does not define, on every replica alike", () => {
    // styles.xml is part of the document, so an unknown id is a no-op every
    // replica reaches independently rather than a divergence.
    const a = tableSession();
    const b = tableSession();
    const op = {
      kind: "setTableStyle", clientId: "c", clientSeq: 10, base: a.s.seq,
      cellParagraphId: a.cellParagraphId(0, 0), styleId: "GridTable4Accent1",
    };
    expect(a.s.submit(op as never).kind).toBe("rejected");
    expect(b.s.submit({ ...op, base: b.s.seq } as never).kind).toBe("rejected");
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
  });

  it("refuses a malformed payload before it is sequenced", () => {
    const { s, cellParagraphId } = tableSession();
    const bad = (body: Record<string, unknown>) =>
      s.submit({ clientId: "c", clientSeq: 10, base: s.seq, cellParagraphId: cellParagraphId(0, 0), ...body } as never).kind;
    // insideV has no meaning on a single cell, so cell scope refuses it.
    expect(bad({ kind: "setTableBorders", scope: "cell", edges: ["insideV"], border: { style: "single" } })).toBe("rejected");
    expect(bad({ kind: "setTableBorders", scope: "table", edges: [], border: null })).toBe("rejected");
    expect(bad({ kind: "setTableWidth", unit: "pct", value: 400 })).toBe("rejected");
    expect(bad({ kind: "setTableWidth", unit: "pt" })).toBe("rejected"); // value required
    expect(bad({ kind: "setTableCellMargins", scope: "cell", margins: { left: -4 } })).toBe("rejected");
    expect(bad({ kind: "setTableLook", look: { bandedRows: "yes" } })).toBe("rejected");
    expect(bad({ kind: "setTableHeaderRows", count: 1.5 })).toBe("rejected");
  });
});

/**
 * The same eight intents plus tableOp, SUGGESTED.
 *
 * A tracked table format writes a w:tblPrChange / w:trPrChange / w:tcPrChange
 * holding what it replaced, stamped with an author and a date. Both come off
 * the wall clock, so they have to be drawn ONCE by the originating client and
 * carried in the intent — a replica that re-derived the date would fork the
 * room on a byte no renderer shows. The w:id is scan-derived from identical
 * tree state and needs no carrying, which is what makes byte equality the
 * right assertion here.
 */
describe("suggested table formatting intents", () => {
  const suggest = { author: "Rae", date: "2026-07-12T00:00:00Z" };

  it("writes the same record on every replica, byte for byte", () => {
    const addr = tableSession().cellParagraphId(1, 1);
    const { xml } = bothApply({
      kind: "setTableWidth", cellParagraphId: addr, unit: "pct", value: 75, suggest,
    });
    // The record holds the width the table HAD, and the id is scan-derived, so
    // both replicas wrote the same one.
    expect(xml).toContain(
      `<w:tblPrChange w:id="1" w:author="Rae" w:date="2026-07-12T00:00:00Z">` +
        `<w:tblPr><w:tblW w:w="0" w:type="auto"/>`,
    );
    expect(xml).toContain(`<w:tblW w:w="3750" w:type="pct"/>`);
  });

  it("records at table, row and cell scope from the right address", () => {
    const addr = tableSession().cellParagraphId(1, 1);
    expect(bothApply({ kind: "setTableHeaderRows", cellParagraphId: addr, count: 2, suggest }).xml)
      .toContain("<w:trPrChange");
    expect(
      bothApply({
        kind: "setTableBorders", cellParagraphId: addr, scope: "cell",
        edges: ["top"], border: { style: "single" }, suggest,
      }).xml,
    ).toContain("<w:tcPrChange");
    expect(
      bothApply({
        kind: "setTableCellMargins", cellParagraphId: addr, scope: "table",
        margins: { left: 9 }, suggest,
      }).xml,
    ).toContain("<w:tblPrChange");
  });

  it("records the grid alongside the table for a column width", () => {
    const addr = tableSession().cellParagraphId(1, 1);
    const { xml } = bothApply({
      kind: "setTableColumnWidth", cellParagraphId: addr, colIdx: 1, widthPt: 120, suggest,
    });
    // CT_TblGridChange carries a w:id and nothing else.
    expect(xml).toMatch(/<w:tblGridChange w:id="\d+"><w:tblGrid>/);
    expect(xml).toContain("<w:tblPrChange");
    expect(xml).toContain("<w:tcPrChange");
  });

  it("tracks the two cell-property tableOps and leaves the rest untracked", () => {
    const addr = tableSession().cellParagraphId(1, 1);
    expect(bothApply({ kind: "tableOp", cellParagraphId: addr, op: { kind: "cellShading", fill: "FF0000" }, suggest }).xml)
      .toContain("<w:tcPrChange");
    expect(bothApply({ kind: "tableOp", cellParagraphId: addr, op: { kind: "cellVAlign", v: "center" }, suggest }).xml)
      .toContain("<w:tcPrChange");
    // A structural op ignores the metadata rather than half-tracking the edit.
    // The agent CLI is what refuses to SEND one in suggesting mode.
    expect(bothApply({ kind: "tableOp", cellParagraphId: addr, op: "deleteRow", suggest }).xml)
      .not.toContain("Change");
  });

  it("converges when two replicas suggest into different cells", () => {
    const a = tableSession();
    const b = tableSession();
    const ops = [
      { kind: "tableOp", cellParagraphId: a.cellParagraphId(0, 0), op: { kind: "cellShading", fill: "FF0000" }, suggest },
      { kind: "tableOp", cellParagraphId: a.cellParagraphId(2, 2), op: { kind: "cellVAlign", v: "bottom" }, suggest },
      { kind: "setTableHeaderRows", cellParagraphId: a.cellParagraphId(1, 0), count: 1, suggest },
    ];
    ops.forEach((body, i) => {
      expect(a.s.submit({ clientId: "c", clientSeq: 10 + i, base: a.s.seq, ...body } as never).kind).toBe("applied");
    });
    // The other replica receives them in the same sequenced order.
    ops.forEach((body, i) => {
      expect(b.s.submit({ clientId: "c", clientSeq: 10 + i, base: b.s.seq, ...body } as never).kind).toBe("applied");
    });
    expect(serializeXml(a.s.doc.docRoot)).toBe(serializeXml(b.s.doc.docRoot));
    expect(collectRevisions(a.s.doc).map((r) => r.kind)).toEqual(["rowFormat", "cellFormat", "cellFormat"]);
  });

  it("refuses a malformed author or date before it is sequenced", () => {
    const { s, cellParagraphId } = tableSession();
    let seq = 10;
    const submit = (suggestValue: unknown) =>
      s.submit({
        kind: "setTableLook", clientId: "c", clientSeq: seq++, base: s.seq,
        cellParagraphId: cellParagraphId(0, 0), look: { bandedRows: false }, suggest: suggestValue,
      } as never).kind;
    expect(submit({ author: 7, date: "x" })).toBe("rejected");
    expect(submit({ author: "a", date: "d".repeat(41) })).toBe("rejected");
    expect(submit(suggest)).toBe("applied");
  });
});
