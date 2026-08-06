import { beforeEach, describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { XmlElement, localName, serializeXml } from "../src/xml.js";
import {
  applyTableOp,
  setTableBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableHeaderRows,
  setTableLayoutMode,
  setTableLook,
  setTableStyle,
  setTableWidth,
} from "../src/edit/tables.js";
import {
  RevisionMeta,
  acceptAllRevisions,
  acceptRevision,
  collectRevisions,
  rejectAllRevisions,
  rejectRevision,
  revisionForText,
} from "../src/edit/suggest.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/**
 * Tracked TABLE formatting: w:tblPrChange, w:trPrChange and w:tcPrChange, plus
 * the w:tblGridChange that rides with a column-width suggestion.
 *
 * These assert exact OOXML because Word has to read the result, and the parity
 * corpus has NO fixture carrying any of the four elements — not even a
 * synthetic one, as it does for w:rPrChange and w:pPrChange. Everything below
 * follows ECMA-376 directly:
 *
 *  - w:id / w:author / w:date on the change element, in that order, matching
 *    what the run and paragraph records already write.
 *  - The change element LAST inside the properties element it closes, which is
 *    where CT_TblPr, CT_TrPr and CT_TcPr each put it.
 *  - One child holding the properties that were replaced, typed as the BASE of
 *    the properties element: CT_TblPrBase, CT_TrPrBase, CT_TcPrInner. A row's
 *    own structural w:ins / w:del is outside CT_TrPrBase and therefore stays
 *    live rather than being recorded.
 *  - w:tblGridChange carries a w:id and NOTHING else (CT_TblGridChange extends
 *    CT_Markup), so it is not independently attributable and is reviewed as
 *    part of the table's w:tblPrChange.
 */

let idCounter = 1;
beforeEach(() => {
  idCounter = 1;
});

const meta = (author = "Alex"): RevisionMeta => ({
  author,
  date: "2026-07-12T00:00:00Z",
  nextId: () => idCounter++,
});

/** rows x cols of text cells, each tagged r{row}c{col}. */
function tableXml(rows: number, cols: number, tblPrExtra = "", cellW = 2000): string {
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${cellW}"/>`).join("");
  const body = Array.from({ length: rows }, (_, r) => {
    const cells = Array.from(
      { length: cols },
      (_, c) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${cellW}"/></w:tcPr>${p(`r${r}c${c}`)}</w:tc>`,
    ).join("");
    return `<w:tr>${cells}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr>${tblPrExtra}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

function load(bodyXml: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml) }));
}

function firstOfName(root: XmlElement, name: string): XmlElement {
  let found: XmlElement | null = null;
  const walk = (e: XmlElement): void => {
    if (found) return;
    if (localName(e.name) === name) found = e;
    else e.children.forEach(walk);
  };
  walk(root);
  if (!found) throw new Error(`no w:${name}`);
  return found;
}

function tblOf(doc: DocxDocument): XmlElement {
  return firstOfName(doc.docRoot, "tbl");
}

/** The w:t carrying `needle`, which is how a caret addresses a cell. */
function findT(doc: DocxDocument, needle: string): XmlElement {
  let found: XmlElement | null = null;
  const walk = (e: XmlElement): void => {
    if (found) return;
    if (localName(e.name) === "t" && (e.text ?? "").includes(needle)) found = e;
    else e.children.forEach(walk);
  };
  walk(doc.docRoot);
  if (!found) throw new Error(`no w:t containing ${needle}`);
  return found;
}

/** The table's w:tblPr, serialized. */
function tblPrXml(doc: DocxDocument): string {
  return serializeXml(firstOfName(tblOf(doc), "tblPr"));
}

/** Row `i`'s w:tr, serialized. */
function rowXml(doc: DocxDocument, i: number): string {
  return serializeXml(tblOf(doc).children.filter((c) => localName(c.name) === "tr")[i]);
}

/** The w:tcPr of the cell holding `needle`, serialized. */
function tcPrXml(doc: DocxDocument, needle: string): string {
  let cur: XmlElement | undefined = doc.findParentOf(findT(doc, needle));
  while (cur && localName(cur.name) !== "tc") cur = doc.findParentOf(cur);
  if (!cur) throw new Error("no w:tc");
  return serializeXml(cur.children.find((c) => localName(c.name) === "tcPr")!);
}

function tblGridXml(doc: DocxDocument): string {
  return serializeXml(firstOfName(tblOf(doc), "tblGrid"));
}

// ---------------------------------------------------------------------------
// w:tblPrChange — table properties
// ---------------------------------------------------------------------------

describe("tracked table formatting (w:tblPrChange)", () => {
  it("records the previous tblPr as the last child of the new one", () => {
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    expect(setTableWidth(doc, tblOf(doc), "pct", 50, meta())).toBe(true);
    expect(tblPrXml(doc)).toBe(
      `<w:tblPr><w:tblW w:w="2500" w:type="pct"/>` +
        `<w:tblPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">` +
        `<w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>` +
        `</w:tblPrChange></w:tblPr>`,
    );
  });

  it("records an EMPTY previous tblPr when the table declared none", () => {
    const doc = load(tableXml(2, 2));
    expect(setTableStyle(doc, tblOf(doc), null, meta())).toBe(false); // nothing to clear
    expect(setTableLook(doc, tblOf(doc), { firstRow: false }, meta())).toBe(true);
    expect(tblPrXml(doc)).toContain(`<w:tblPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:tblPr/></w:tblPrChange>`);
  });

  it("keeps the change element last when a later property is inserted", () => {
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    setTableWidth(doc, tblOf(doc), "pct", 50, meta());
    // A second suggested change refines the same pending suggestion: the
    // record stays the state before ANY of it, and tblLook lands ahead of it.
    setTableLook(doc, tblOf(doc), { bandedRows: false }, meta());
    const xml = tblPrXml(doc);
    expect(xml.indexOf("<w:tblLook")).toBeLessThan(xml.indexOf("<w:tblPrChange"));
    expect(xml).toContain(`<w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr></w:tblPrChange>`);
    expect(collectRevisions(doc).filter((r) => r.kind === "tableFormat")).toHaveLength(1);
  });

  it("accepts by dropping the record and keeping the new width", () => {
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    setTableWidth(doc, tblOf(doc), "pct", 50, meta());
    const [ref] = collectRevisions(doc);
    expect(ref.kind).toBe("tableFormat");
    expect(acceptRevision(doc, ref)).toBe(true);
    expect(tblPrXml(doc)).toBe(`<w:tblPr><w:tblW w:w="2500" w:type="pct"/></w:tblPr>`);
    expect(collectRevisions(doc)).toHaveLength(0);
  });

  it("rejects by putting the previous properties back", () => {
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    setTableWidth(doc, tblOf(doc), "pct", 50, meta());
    expect(rejectRevision(doc, collectRevisions(doc)[0])).toBe(true);
    expect(tblPrXml(doc)).toBe(`<w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>`);
  });

  it("rejects a table that had no properties by removing the tblPr", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false }, meta());
    rejectRevision(doc, collectRevisions(doc)[0]);
    expect(tblOf(doc).children.some((c) => localName(c.name) === "tblPr")).toBe(false);
  });

  it("records table-scoped borders and cell margins on the table", () => {
    const doc = load(tableXml(2, 2));
    setTableBorders(doc, findT(doc, "r0c0"), "table", ["top"], { style: "single" }, meta());
    setTableCellMargins(doc, findT(doc, "r0c0"), "table", { left: 6 }, meta());
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["tableFormat"]);
    rejectAllRevisions(doc);
    expect(tblOf(doc).children.some((c) => localName(c.name) === "tblPr")).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// w:trPrChange — row properties
// ---------------------------------------------------------------------------

describe("tracked row formatting (w:trPrChange)", () => {
  it("records only the rows whose header membership moves", () => {
    const doc = load(tableXml(3, 2));
    expect(setTableHeaderRows(doc, tblOf(doc), 1, meta())).toBe(true);
    expect(rowXml(doc, 0)).toContain(
      `<w:trPr><w:tblHeader/><w:trPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:trPr/></w:trPrChange></w:trPr>`,
    );
    expect(rowXml(doc, 1)).not.toContain("trPr");
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["rowFormat"]);
  });

  it("rejects a header row back to no trPr at all", () => {
    const doc = load(tableXml(3, 2));
    setTableHeaderRows(doc, tblOf(doc), 2, meta());
    expect(collectRevisions(doc)).toHaveLength(2);
    expect(rejectAllRevisions(doc)).toBe(2);
    expect(rowXml(doc, 0)).not.toContain("trPr");
    expect(rowXml(doc, 1)).not.toContain("trPr");
  });

  it("accepts a header row by keeping the tblHeader and dropping the record", () => {
    const doc = load(tableXml(3, 2));
    setTableHeaderRows(doc, tblOf(doc), 1, meta());
    expect(acceptAllRevisions(doc)).toBe(1);
    expect(rowXml(doc, 0)).toContain("<w:trPr><w:tblHeader/></w:trPr>");
  });

  it("leaves a row's structural w:ins outside the record and untouched", () => {
    // CT_TrPrBase carries no w:ins, so a row a WORD user already suggested as
    // inserted must keep that revision live rather than have it recorded as
    // formatting — and rejecting the formatting must not take it away.
    const inserted = `<w:trPr><w:ins w:id="90" w:author="Word" w:date="2026-01-01T00:00:00Z"/></w:trPr>`;
    const doc = load(
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr>${inserted}<w:tc><w:tcPr/>${p("r0c0")}</w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr/>${p("r1c0")}</w:tc></w:tr></w:tbl>`,
    );
    setTableHeaderRows(doc, tblOf(doc), 1, meta());
    expect(rowXml(doc, 0)).toContain(
      `<w:trPr><w:tblHeader/><w:ins w:id="90" w:author="Word" w:date="2026-01-01T00:00:00Z"/>` +
        `<w:trPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z"><w:trPr/></w:trPrChange></w:trPr>`,
    );
    rejectAllRevisions(doc);
    expect(rowXml(doc, 0)).toContain(`<w:trPr><w:ins w:id="90" w:author="Word" w:date="2026-01-01T00:00:00Z"/></w:trPr>`);
  });

  it("never collects a row's structural w:ins as a run-level revision", () => {
    const doc = load(
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:ins w:id="90" w:author="Word" w:date="2026-01-01T00:00:00Z"/></w:trPr>` +
        `<w:tc><w:tcPr/>${p("r0c0")}</w:tc></w:tr></w:tbl>`,
    );
    // Unwrapping it as an "insertion" would splice its children into the trPr.
    expect(collectRevisions(doc)).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// w:tcPrChange — cell properties
// ---------------------------------------------------------------------------

describe("tracked cell formatting (w:tcPrChange)", () => {
  it("records the previous tcPr for a cell shading suggestion", () => {
    const doc = load(tableXml(2, 2));
    expect(applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellShading", fill: "FF0000" }, meta())).toBe(true);
    expect(tcPrXml(doc, "r0c0")).toBe(
      `<w:tcPr><w:tcW w:type="dxa" w:w="2000"/><w:shd w:val="clear" w:fill="FF0000"/>` +
        `<w:tcPrChange w:id="1" w:author="Alex" w:date="2026-07-12T00:00:00Z">` +
        `<w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>` +
        `</w:tcPrChange></w:tcPr>`,
    );
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["cellFormat"]);
  });

  it("keeps the change element last when vAlign is added afterwards", () => {
    const doc = load(tableXml(2, 2));
    applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellShading", fill: "FF0000" }, meta());
    applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellVAlign", v: "center" }, meta());
    const xml = tcPrXml(doc, "r0c0");
    expect(xml.indexOf("<w:vAlign")).toBeLessThan(xml.indexOf("<w:tcPrChange"));
    expect(collectRevisions(doc)).toHaveLength(1);
  });

  it("rejects cell shading back to the properties the cell had", () => {
    const doc = load(tableXml(2, 2));
    applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellShading", fill: "FF0000" }, meta());
    rejectRevision(doc, collectRevisions(doc)[0]);
    expect(tcPrXml(doc, "r0c0")).toBe(`<w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr>`);
  });

  it("records cell-scoped borders on the one cell, not the table", () => {
    const doc = load(tableXml(2, 2));
    setTableBorders(doc, findT(doc, "r1c1"), "cell", ["top"], { style: "double" }, meta());
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["cellFormat"]);
    expect(tcPrXml(doc, "r1c1")).toContain("<w:tcPrChange");
    expect(tcPrXml(doc, "r0c0")).not.toContain("tcPrChange");
  });
});

// ---------------------------------------------------------------------------
// Column widths: the tblGrid record that rides with the table's
// ---------------------------------------------------------------------------

describe("tracked column widths (w:tblGridChange)", () => {
  it("records the grid, the table total and every cell", () => {
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    expect(setTableColumnWidth(doc, tblOf(doc), 0, 150, meta())).toBe(true);
    // CT_TblGridChange extends CT_Markup: a w:id and nothing else.
    expect(tblGridXml(doc)).toBe(
      `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="2000"/>` +
        `<w:tblGridChange w:id="2"><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
        `</w:tblGridChange></w:tblGrid>`,
    );
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual([
      "tableFormat",
      "cellFormat",
      "cellFormat",
      "cellFormat",
      "cellFormat",
    ]);
  });

  it("rejects a column width back to the grid, the total and the cells it had", () => {
    const before = tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`);
    const doc = load(before);
    setTableColumnWidth(doc, tblOf(doc), 0, 150, meta());
    expect(rejectAllRevisions(doc)).toBe(5);
    expect(serializeXml(tblOf(doc))).toBe(before);
  });

  it("accepts a column width by retiring the grid record with the table's", () => {
    const doc = load(tableXml(2, 2, `<w:tblW w:w="4000" w:type="dxa"/>`));
    setTableColumnWidth(doc, tblOf(doc), 0, 150, meta());
    expect(acceptAllRevisions(doc)).toBe(5);
    expect(tblGridXml(doc)).toBe(`<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="2000"/></w:tblGrid>`);
    expect(serializeXml(tblOf(doc))).not.toContain("Change");
  });

  it("rejects a layout switch back to autofit's absent cell widths", () => {
    const before = tableXml(2, 2, `<w:tblW w:w="0" w:type="auto"/>`);
    const doc = load(before);
    expect(setTableLayoutMode(doc, tblOf(doc), "autofit", undefined, meta())).toBe(true);
    expect(rejectAllRevisions(doc)).toBe(5);
    expect(serializeXml(tblOf(doc))).toBe(before);
  });
});

// ---------------------------------------------------------------------------
// Review order and caret resolution
// ---------------------------------------------------------------------------

describe("reviewing tracked table formatting", () => {
  it("collects table, row and cell records in document order", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false }, meta());
    setTableHeaderRows(doc, tblOf(doc), 1, meta());
    applyTableOp(doc, findT(doc, "r1c0"), { kind: "cellShading", fill: "00FF00" }, meta());
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["tableFormat", "rowFormat", "cellFormat"]);
  });

  it("accepts everything in one pass, reverse order and all", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false }, meta());
    setTableHeaderRows(doc, tblOf(doc), 1, meta());
    applyTableOp(doc, findT(doc, "r1c0"), { kind: "cellShading", fill: "00FF00" }, meta());
    expect(acceptAllRevisions(doc)).toBe(3);
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(serializeXml(tblOf(doc))).not.toContain("Change");
  });

  it("resolves the innermost record from a caret in a cell", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false }, meta());
    setTableHeaderRows(doc, tblOf(doc), 1, meta());
    applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellShading", fill: "00FF00" }, meta());
    // Cell, row and table records all enclose this caret; the cell is nearest.
    expect(revisionForText(doc, findT(doc, "r0c0"))?.kind).toBe("cellFormat");
    // A caret in a cell with no record of its own falls out to the row's.
    expect(revisionForText(doc, findT(doc, "r0c1"))?.kind).toBe("rowFormat");
    // Row 1 carries no record, so the table's is the nearest enclosing one.
    expect(revisionForText(doc, findT(doc, "r1c0"))?.kind).toBe("tableFormat");
  });

  it("accepts just the cell record the caret resolved", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false }, meta());
    applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellShading", fill: "00FF00" }, meta());
    const ref = revisionForText(doc, findT(doc, "r0c0"))!;
    expect(acceptRevision(doc, ref)).toBe(true);
    expect(collectRevisions(doc).map((r) => r.kind)).toEqual(["tableFormat"]);
    expect(tcPrXml(doc, "r0c0")).toContain(`<w:shd w:val="clear" w:fill="00FF00"/>`);
  });

  it("carries the author onto every record so the popover can name it", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false }, meta("Rae"));
    setTableHeaderRows(doc, tblOf(doc), 1, meta("Rae"));
    applyTableOp(doc, findT(doc, "r1c0"), { kind: "cellShading", fill: "00FF00" }, meta("Rae"));
    expect(collectRevisions(doc).map((r) => r.author)).toEqual(["Rae", "Rae", "Rae"]);
  });

  it("leaves the document untouched without meta (a direct edit)", () => {
    const doc = load(tableXml(2, 2));
    setTableLook(doc, tblOf(doc), { firstRow: false });
    applyTableOp(doc, findT(doc, "r0c0"), { kind: "cellShading", fill: "00FF00" });
    expect(collectRevisions(doc)).toHaveLength(0);
    expect(serializeXml(tblOf(doc))).not.toContain("Change");
  });
});
