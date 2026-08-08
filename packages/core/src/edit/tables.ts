import { DocxDocument } from "../docx.js";
import { XmlElement, attr, child, cloneXml, localName, onOff } from "../xml.js";
import { pxToTwips } from "../units.js";
import {
  RevisionMeta,
  recordCellFormatChange,
  recordRowFormatChange,
  recordTableFormatChange,
  recordTableGridChange,
} from "./suggest.js";

/**
 * Table manipulation: add/remove rows and columns relative to the cell
 * containing the caret. Mutates source XML; callers checkpoint + relayout.
 *
 * v1 scope: column operations refuse tables using gridSpan (merged cells)
 * rather than corrupt them.
 *
 * SUGGESTING MODE. Every formatting operation below takes an optional
 * RevisionMeta. With one, it records the properties it is about to replace as
 * a tracked change — w:tblPrChange on the table, w:trPrChange on a row,
 * w:tcPrChange on a cell — before mutating them, so a reviewer can accept or
 * reject the change (edit/suggest.ts). Without one it mutates outright, which
 * is what a direct edit does. STRUCTURAL operations (row and column insert and
 * delete, merge, split) take no meta: they are not tracked, and the surfaces
 * that enforce suggesting mode refuse them rather than apply them untracked.
 */

export type TableOp =
  | "rowAbove"
  | "rowBelow"
  | "deleteRow"
  | "colLeft"
  | "colRight"
  | "deleteCol"
  | "deleteTable"
  | "mergeRight"
  | "mergeDown"
  | "splitCell"
  | { kind: "cellShading"; fill: string | null }
  | { kind: "cellVAlign"; v: "top" | "center" | "bottom" }
  | { kind: "textWrapping"; wrapping: "none" | "around"; xPx: number; yPx: number };

interface CellContext {
  tbl: XmlElement;
  tblParent: XmlElement;
  tr: XmlElement;
  tc: XmlElement;
  rowIdx: number;
  cellIdx: number;
  w: string;
}

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

function rowsOf(tbl: XmlElement): XmlElement[] {
  return tbl.children.filter((c) => localName(c.name) === "tr");
}
function cellsOf(tr: XmlElement): XmlElement[] {
  return tr.children.filter((c) => localName(c.name) === "tc");
}

/** Locate the table cell containing a (caret) element. */
export function cellContextOf(doc: DocxDocument, target: XmlElement): CellContext | null {
  let tc: XmlElement | undefined;
  let cur: XmlElement | undefined = doc.findParentOf(target);
  while (cur) {
    const ln = localName(cur.name);
    if (ln === "tc") tc = cur;
    if (ln === "tr" && tc) {
      const tr = cur;
      const tbl = doc.findParentOf(tr);
      if (!tbl || localName(tbl.name) !== "tbl") return null;
      const tblParent = doc.findParentOf(tbl);
      if (!tblParent) return null;
      return {
        tbl,
        tblParent,
        tr,
        tc,
        rowIdx: rowsOf(tbl).indexOf(tr),
        cellIdx: cellsOf(tr).indexOf(tc),
        w: prefixOf(tc),
      };
    }
    cur = doc.findParentOf(cur);
  }
  return null;
}

/**
 * Word's cell-granular selection: when two (caret) elements sit in DIFFERENT
 * cells of the SAME table, the selection covers the whole rectangular grid
 * range between the two cells. Returns those cells in document order, or null
 * when the elements are not in distinct cells of one table (same cell, only
 * one inside a table, or different/nested tables) — the caller then keeps
 * character-granular behavior.
 */
export function cellRangeBetween(
  doc: DocxDocument,
  targetA: XmlElement,
  targetB: XmlElement,
): XmlElement[] | null {
  const a = cellContextOf(doc, targetA);
  const b = cellContextOf(doc, targetB);
  if (!a || !b || a.tbl !== b.tbl || a.tc === b.tc) return null;
  const rows = rowsOf(a.tbl);
  const colA = gridColOf(a.tr, a.tc);
  const colB = gridColOf(b.tr, b.tc);
  const rowLo = Math.min(a.rowIdx, b.rowIdx);
  const rowHi = Math.max(a.rowIdx, b.rowIdx);
  const colLo = Math.min(colA, colB);
  const colHi = Math.max(colA, colB);
  const out: XmlElement[] = [];
  for (let ri = rowLo; ri <= rowHi; ri++) {
    for (const tc of cellsOf(rows[ri])) {
      const col = gridColOf(rows[ri], tc);
      if (col <= colHi && col + gridSpanOf(tc) - 1 >= colLo) out.push(tc);
    }
  }
  return out;
}

/** Current cell fill, undefined outside a table and null for no direct fill. */
export function cellShadingAt(doc: DocxDocument, target: XmlElement): string | null | undefined {
  const ctx = cellContextOf(doc, target);
  if (!ctx) return undefined;
  const shd = child(child(ctx.tc, "tcPr"), "shd");
  const fill = shd ? attr(shd, "fill") : undefined;
  return fill && fill.toLowerCase() !== "auto" ? `#${fill.replace(/^#/, "").toUpperCase()}` : null;
}

function emptyCellLike(tc: XmlElement, w: string): XmlElement {
  const tcPr = tc.children.find((c) => localName(c.name) === "tcPr");
  return {
    name: `${w}tc`,
    attrs: {},
    text: "",
    children: [
      ...(tcPr ? [cloneXml(tcPr)] : []),
      {
        name: `${w}p`,
        attrs: {},
        text: "",
        children: [
          {
            name: `${w}r`,
            attrs: {},
            text: "",
            children: [{ name: `${w}t`, attrs: { "xml:space": "preserve" }, text: "", children: [] }],
          },
        ],
      },
    ],
  };
}

/** OOXML tcPr child order (the slice we edit). w:tcPrChange closes the
 * element, so a property inserted later still lands ahead of it. */
const TCPR_ORDER = [
  "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "tcMar", "vAlign", "tcPrChange",
];

function ensureTcPr(tc: XmlElement, w: string): XmlElement {
  let tcPr = tc.children.find((c) => localName(c.name) === "tcPr");
  if (!tcPr) {
    tcPr = { name: `${w}tcPr`, attrs: {}, children: [], text: "" };
    tc.children.unshift(tcPr);
  }
  return tcPr;
}

/** Set (or with attrs=null remove) a tcPr child, keeping schema order. */
function setTcPrChild(tc: XmlElement, w: string, name: string, attrs: Record<string, string> | null): void {
  const tcPr = ensureTcPr(tc, w);
  const idx = tcPr.children.findIndex((c) => localName(c.name) === name);
  if (attrs === null) {
    if (idx !== -1) tcPr.children.splice(idx, 1);
    return;
  }
  const el: XmlElement = { name: `${w}${name}`, attrs, children: [], text: "" };
  if (idx !== -1) {
    tcPr.children[idx] = el;
    return;
  }
  const rank = TCPR_ORDER.indexOf(name);
  let insertAt = tcPr.children.length;
  for (let i = 0; i < tcPr.children.length; i++) {
    const r = TCPR_ORDER.indexOf(localName(tcPr.children[i].name));
    if (r !== -1 && r > rank) {
      insertAt = i;
      break;
    }
  }
  tcPr.children.splice(insertAt, 0, el);
}

function gridSpanOf(tc: XmlElement): number {
  const tcPr = child(tc, "tcPr");
  const gs = tcPr ? child(tcPr, "gridSpan") : undefined;
  return gs ? parseInt(attr(gs, "val") ?? "1", 10) || 1 : 1;
}

/** Grid column index of a cell (accounts for gridSpans to its left). */
function gridColOf(tr: XmlElement, tc: XmlElement): number {
  let col = 0;
  for (const c of cellsOf(tr)) {
    if (c === tc) return col;
    col += gridSpanOf(c);
  }
  return col;
}

/** The cell of `tr` occupying grid column `col`, if any. */
function cellAtGridCol(tr: XmlElement, col: number): XmlElement | null {
  let cur = 0;
  for (const c of cellsOf(tr)) {
    if (cur === col) return c;
    cur += gridSpanOf(c);
    if (cur > col) return null; // covered mid-span
  }
  return null;
}

function contentOf(tc: XmlElement): XmlElement[] {
  return tc.children.filter((c) => localName(c.name) !== "tcPr");
}

function hasText(tc: XmlElement): boolean {
  const walk = (e: XmlElement): boolean => {
    if (localName(e.name) === "t" && e.text.trim()) return true;
    return e.children.some(walk);
  };
  return contentOf(tc).some(walk);
}

function hasSpans(tbl: XmlElement): boolean {
  for (const tr of rowsOf(tbl)) {
    for (const tc of cellsOf(tr)) {
      const tcPr = child(tc, "tcPr");
      if (tcPr && child(tcPr, "gridSpan")) return true;
    }
  }
  return false;
}

/**
 * `meta` tracks the three FORMATTING ops (cell shading, cell vertical
 * alignment, table text wrapping). The structural ops ignore it: a row or
 * column insert or delete, a merge and a split have no tracked form here, so
 * every surface that enforces suggesting mode refuses them instead.
 */
export function applyTableOp(
  doc: DocxDocument,
  target: XmlElement,
  op: TableOp,
  meta?: RevisionMeta,
): boolean {
  if (typeof op === "object" || op === "mergeRight" || op === "mergeDown" || op === "splitCell") {
    const ctx = cellContextOf(doc, target);
    if (!ctx) return false;
    const { tbl, tr, tc, w } = ctx;
    if (typeof op === "object") {
      if (op.kind === "textWrapping") {
        return setTableTextWrapping(doc, tbl, op.wrapping, op.xPx, op.yPx, meta);
      }
      if (meta) recordCellFormatChange(tc, meta);
      if (op.kind === "cellShading") {
        setTcPrChild(tc, w, "shd", op.fill === null ? null : { [`${w}val`]: "clear", [`${w}fill`]: op.fill.replace(/^#/, "").toUpperCase() });
      } else {
        setTcPrChild(tc, w, "vAlign", { [`${w}val`]: op.v });
      }
      doc.refresh();
      return true;
    }
    if (op === "mergeRight") {
      const cells = cellsOf(tr);
      const next = cells[cells.indexOf(tc) + 1];
      if (!next) return false;
      setTcPrChild(tc, w, "gridSpan", { [`${w}val`]: String(gridSpanOf(tc) + gridSpanOf(next)) });
      // Word keeps both cells' content, concatenated.
      if (hasText(next)) tc.children.push(...contentOf(next));
      tr.children.splice(tr.children.indexOf(next), 1);
      doc.refresh();
      return true;
    }
    if (op === "mergeDown") {
      const rows = rowsOf(tbl);
      const below = rows[rows.indexOf(tr) + 1];
      if (!below) return false;
      const col = gridColOf(tr, tc);
      const target2 = cellAtGridCol(below, col);
      if (!target2 || gridSpanOf(target2) !== gridSpanOf(tc)) return false;
      setTcPrChild(tc, w, "vMerge", { [`${w}val`]: "restart" });
      if (hasText(target2)) tc.children.push(...contentOf(target2));
      // The continued cell keeps an empty paragraph (schema requires one).
      const empty = emptyCellLike(target2, w);
      target2.children = empty.children;
      setTcPrChild(target2, w, "vMerge", {});
      doc.refresh();
      return true;
    }
    // splitCell: undo a horizontal span and/or a vertical merge start.
    let touched = false;
    const span = gridSpanOf(tc);
    if (span > 1) {
      setTcPrChild(tc, w, "gridSpan", null);
      const at = tr.children.indexOf(tc);
      const extras = Array.from({ length: span - 1 }, () => {
        const cell = emptyCellLike(tc, w);
        const tcPr = cell.children.find((c) => localName(c.name) === "tcPr");
        if (tcPr) tcPr.children = tcPr.children.filter((c) => !["gridSpan", "vMerge"].includes(localName(c.name)));
        return cell;
      });
      tr.children.splice(at + 1, 0, ...extras);
      touched = true;
    }
    const tcPrNow = child(tc, "tcPr");
    const vm = tcPrNow ? child(tcPrNow, "vMerge") : undefined;
    if (vm && (attr(vm, "val") ?? "restart") === "restart") {
      setTcPrChild(tc, w, "vMerge", null);
      const rows = rowsOf(tbl);
      const col = gridColOf(tr, tc);
      for (let ri = rows.indexOf(tr) + 1; ri < rows.length; ri++) {
        const cell = cellAtGridCol(rows[ri], col);
        if (!cell) break;
        const pr = child(cell, "tcPr");
        const cvm = pr ? child(pr, "vMerge") : undefined;
        if (!cvm || attr(cvm, "val") === "restart") break;
        setTcPrChild(cell, w, "vMerge", null);
      }
      touched = true;
    }
    if (touched) doc.refresh();
    return touched;
  }
  const ctx = cellContextOf(doc, target);
  if (!ctx) return false;
  const { tbl, tblParent, tr, rowIdx, cellIdx, w } = ctx;

  switch (op) {
    case "rowAbove":
    case "rowBelow": {
      const newRow: XmlElement = {
        name: `${w}tr`,
        attrs: {},
        text: "",
        children: cellsOf(tr).map((tc) => emptyCellLike(tc, w)),
      };
      const trPr = tr.children.find((c) => localName(c.name) === "trPr");
      if (trPr) newRow.children.unshift(cloneXml(trPr));
      const idx = tbl.children.indexOf(tr);
      tbl.children.splice(op === "rowAbove" ? idx : idx + 1, 0, newRow);
      break;
    }
    case "deleteRow": {
      if (rowsOf(tbl).length <= 1) return applyTableOp(doc, target, "deleteTable");
      tbl.children.splice(tbl.children.indexOf(tr), 1);
      break;
    }
    case "colLeft":
    case "colRight": {
      if (hasSpans(tbl)) return false;
      const insertAt = op === "colLeft" ? cellIdx : cellIdx + 1;
      for (const row of rowsOf(tbl)) {
        const cells = cellsOf(row);
        const ref = cells[Math.min(cellIdx, cells.length - 1)];
        if (!ref) continue;
        const domIdx =
          insertAt >= cells.length
            ? row.children.indexOf(cells[cells.length - 1]) + 1
            : row.children.indexOf(cells[insertAt]);
        row.children.splice(domIdx, 0, emptyCellLike(ref, w));
      }
      const grid = child(tbl, "tblGrid");
      if (grid) {
        const cols = grid.children.filter((c) => localName(c.name) === "gridCol");
        const ref = cols[Math.min(cellIdx, cols.length - 1)];
        if (ref) {
          const domIdx =
            insertAt >= cols.length
              ? grid.children.indexOf(cols[cols.length - 1]) + 1
              : grid.children.indexOf(cols[insertAt]);
          grid.children.splice(domIdx, 0, cloneXml(ref));
        }
      }
      break;
    }
    case "deleteCol": {
      if (hasSpans(tbl)) return false;
      if (cellsOf(tr).length <= 1) return applyTableOp(doc, target, "deleteTable");
      for (const row of rowsOf(tbl)) {
        const cells = cellsOf(row);
        const victim = cells[cellIdx];
        if (victim) row.children.splice(row.children.indexOf(victim), 1);
      }
      const grid = child(tbl, "tblGrid");
      if (grid) {
        const cols = grid.children.filter((c) => localName(c.name) === "gridCol");
        const victim = cols[cellIdx];
        if (victim) grid.children.splice(grid.children.indexOf(victim), 1);
      }
      break;
    }
    case "deleteTable": {
      tblParent.children.splice(tblParent.children.indexOf(tbl), 1);
      break;
    }
  }
  doc.refresh();
  return true;
}

/** First w:t inside a cell (document order), creating an empty run in the
 * cell's first paragraph when the cell has no typed text yet. */
function firstTextOfCell(tc: XmlElement, w: string): XmlElement | null {
  const walk = (el: XmlElement): XmlElement | null => {
    for (const c of el.children) {
      if (localName(c.name) === "t") return c;
      const found = walk(c);
      if (found) return found;
    }
    return null;
  };
  const existing = walk(tc);
  if (existing) return existing;
  const para = tc.children.find((c) => localName(c.name) === "p");
  if (!para) return null;
  const t: XmlElement = { name: `${w}t`, attrs: { "xml:space": "preserve" }, text: "", children: [] };
  para.children.push({ name: `${w}r`, attrs: {}, text: "", children: [t] });
  return t;
}

/**
 * Tab-to-next-cell navigation. From the cell containing `target`, advance to
 * the next cell (dir=1) or previous cell (dir=-1), wrapping across rows. Tab in
 * the LAST cell of the table appends a new row and lands in its first cell
 * (Word-familiar, and keeps the user from ever being stuck at the table end).
 * Shift+Tab in the very first cell has nowhere to go and returns null.
 *
 * Returns the w:t the caret should move to (offset 0), or null when the target
 * is not inside a table or navigation is a no-op — the caller then leaves the
 * key to its normal handling.
 */
export function advanceCell(doc: DocxDocument, target: XmlElement, dir: 1 | -1): { t: XmlElement } | null {
  const ctx = cellContextOf(doc, target);
  if (!ctx) return null;
  const { tbl, tr, w } = ctx;
  const rows = rowsOf(tbl);
  const cells = cellsOf(tr);
  let destCell: XmlElement | null = null;

  if (dir === 1) {
    if (ctx.cellIdx < cells.length - 1) {
      destCell = cells[ctx.cellIdx + 1];
    } else if (ctx.rowIdx < rows.length - 1) {
      destCell = cellsOf(rows[ctx.rowIdx + 1])[0] ?? null;
    } else {
      // Last cell of the last row: append a fresh row like Word does.
      const newRow: XmlElement = {
        name: `${w}tr`,
        attrs: {},
        text: "",
        children: cells.map((c) => emptyCellLike(c, w)),
      };
      const trPr = tr.children.find((c) => localName(c.name) === "trPr");
      if (trPr) newRow.children.unshift(cloneXml(trPr));
      tbl.children.splice(tbl.children.indexOf(tr) + 1, 0, newRow);
      destCell = cellsOf(newRow)[0] ?? null;
    }
  } else {
    if (ctx.cellIdx > 0) {
      destCell = cells[ctx.cellIdx - 1];
    } else if (ctx.rowIdx > 0) {
      const prevCells = cellsOf(rows[ctx.rowIdx - 1]);
      destCell = prevCells[prevCells.length - 1] ?? null;
    } else {
      return null; // first cell of the table — nowhere earlier to go
    }
  }
  if (!destCell) return null;
  const t = firstTextOfCell(destCell, w);
  if (!t) return null;
  doc.refresh();
  return { t };
}

const MIN_COL_TWIPS = 144; // 0.1"
const MIN_COL_PX = 12;

/**
 * Drag-resize a table column boundary by deltaPx, working in the RENDERED
 * pixel space (grids are often stale/percent-scaled). Like Word, dragging
 * normalizes the table to explicit fixed widths: the grid is rewritten to
 * the rendered widths with the delta applied, and tblW becomes dxa.
 */
export function resizeTableColumn(
  doc: DocxDocument,
  tblEl: XmlElement,
  boundary: number,
  deltaPx: number,
  renderedWidths?: number[],
): boolean {
  const grid = child(tblEl, "tblGrid");
  if (!grid) return false;
  const cols = grid.children.filter((c) => localName(c.name) === "gridCol");
  if (boundary < 1 || boundary > cols.length) return false;

  const widthOf = (c: XmlElement): number => {
    const key = Object.keys(c.attrs).find((k) => localName(k) === "w");
    return key ? parseInt(c.attrs[key], 10) || 0 : 0;
  };
  const setWidth = (c: XmlElement, tw: number) => {
    const key = Object.keys(c.attrs).find((k) => localName(k) === "w") ?? prefixOf(c) + "w";
    c.attrs[key] = String(Math.round(tw));
  };

  if (renderedWidths && renderedWidths.length === cols.length) {
    const px = [...renderedWidths];
    if (boundary === cols.length) {
      px[boundary - 1] = Math.max(px[boundary - 1] + deltaPx, MIN_COL_PX);
    } else {
      const d = Math.max(
        Math.min(deltaPx, px[boundary] - MIN_COL_PX),
        MIN_COL_PX - px[boundary - 1],
      );
      px[boundary - 1] += d;
      px[boundary] -= d;
    }
    cols.forEach((c, i) => setWidth(c, pxToTwips(px[i])));
    // Word converts dragged tables to explicit fixed widths.
    const tblPr = child(tblEl, "tblPr");
    const tblW = tblPr ? child(tblPr, "tblW") : undefined;
    if (tblW) {
      const total = Math.round(pxToTwips(px.reduce((a, b) => a + b, 0)));
      const wKey = Object.keys(tblW.attrs).find((k) => localName(k) === "w") ?? prefixOf(tblW) + "w";
      const tKey = Object.keys(tblW.attrs).find((k) => localName(k) === "type") ?? prefixOf(tblW) + "type";
      tblW.attrs[wKey] = String(total);
      tblW.attrs[tKey] = "dxa";
    }
  } else {
    let d = Math.round(pxToTwips(deltaPx));
    const left = cols[boundary - 1];
    const lw = widthOf(left);
    if (boundary === cols.length) {
      setWidth(left, Math.max(lw + d, MIN_COL_TWIPS));
    } else {
      const right = cols[boundary];
      const rw = widthOf(right);
      d = Math.max(Math.min(d, rw - MIN_COL_TWIPS), MIN_COL_TWIPS - lw);
      setWidth(left, lw + d);
      setWidth(right, rw - d);
    }
  }

  stampCellWidthsFromGrid(tblEl);
  doc.refresh();
  return true;
}

/** Twips in a w:gridCol / other w-attributed element, 0 when absent. */
function widthAttrOf(c: XmlElement): number {
  const key = Object.keys(c.attrs).find((k) => localName(k) === "w");
  return key ? parseInt(c.attrs[key], 10) || 0 : 0;
}

function gridColsOf(tblEl: XmlElement): XmlElement[] {
  const grid = child(tblEl, "tblGrid");
  return grid ? grid.children.filter((c) => localName(c.name) === "gridCol") : [];
}

/**
 * Stamp explicit tcW (dxa) on EVERY cell from the table's grid — gridSpan
 * cells get the sum of their covered grid columns, and pct/auto cells convert
 * to fixed. Word does this whenever it commits a table to explicit widths (a
 * boundary drag, a typed column width, "Fixed Column Width").
 *
 * It must be unconditional: the layout only trusts a grid whose cells declare
 * widths (a grid with no tcW anywhere autofits, Word semantics), so skipping
 * spanned tables here made their drags silent no-ops — the rewritten grid
 * stayed untrusted and autofit snapped the columns back.
 */
function stampCellWidthsFromGrid(tblEl: XmlElement, meta?: RevisionMeta): void {
  const cols = gridColsOf(tblEl);
  for (const tr of rowsOf(tblEl)) {
    let col = 0;
    for (const tc of cellsOf(tr)) {
      const w = prefixOf(tc);
      const span = gridSpanOf(tc);
      const total = cols.slice(col, col + span).reduce((a, c) => a + widthAttrOf(c), 0);
      col += span;
      if (!total) continue;
      // Every restamped cell is its own w:tcPrChange, which is what Word
      // writes: the width lives per cell, so the record of the old one does
      // too, and a reviewer rejecting the change gets every cell back.
      if (meta) recordCellFormatChange(tc, meta);
      setTcPrChild(tc, w, "tcW", { [`${w}type`]: "dxa", [`${w}w`]: String(total) });
    }
  }
}

/** Drag-resize a table row: set trHeight (atLeast) on the given row. */
export function resizeTableRow(
  doc: DocxDocument,
  tblEl: XmlElement,
  rowIdx: number,
  newHeightPx: number,
): boolean {
  const rows = rowsOf(tblEl);
  const tr = rows[rowIdx];
  if (!tr) return false;
  const w = prefixOf(tr);
  let trPr = tr.children.find((c) => localName(c.name) === "trPr");
  if (!trPr) {
    trPr = { name: `${w}trPr`, attrs: {}, children: [], text: "" };
    tr.children.unshift(trPr);
  }
  const tw = String(Math.max(Math.round(pxToTwips(newHeightPx)), 120));
  let h = trPr.children.find((c) => localName(c.name) === "trHeight");
  if (!h) {
    h = { name: `${w}trHeight`, attrs: {}, children: [], text: "" };
    trPr.children.push(h);
  }
  const valKey = Object.keys(h.attrs).find((k) => localName(k) === "val") ?? `${w}val`;
  const ruleKey = Object.keys(h.attrs).find((k) => localName(k) === "hRule") ?? `${w}hRule`;
  h.attrs[valKey] = tw;
  if (h.attrs[ruleKey] !== "exact") h.attrs[ruleKey] = "atLeast";
  doc.refresh();
  return true;
}

/** Position a table at page coordinates, converting an inline table to a float. */
export function moveTableTo(
  doc: DocxDocument,
  tblEl: XmlElement,
  xPx: number,
  yPx: number,
  preservePageStart = false,
  pageDelta = 0,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const w = prefixOf(tblEl);
  let tblPr = child(tblEl, "tblPr");
  if (!tblPr) {
    tblPr = { name: `${w}tblPr`, attrs: {}, children: [], text: "" };
    tblEl.children.unshift(tblPr);
  }
  let position = child(tblPr, "tblpPr");
  const wasInline = !position;
  if (!position) {
    position = { name: `${w}tblpPr`, attrs: {}, children: [], text: "" };
    const styleIdx = tblPr.children.findIndex((c) => localName(c.name) === "tblStyle");
    tblPr.children.splice(styleIdx + 1, 0, position);
  }
  let pageBreakAdjustment = pageDelta + (wasInline && preservePageStart ? 1 : 0);
  const parent = doc.findParentOf(tblEl);
  if (parent && pageBreakAdjustment > 0) {
    const tableIndex = parent.children.indexOf(tblEl);
    let previous = tableIndex > 0 ? parent.children[tableIndex - 1] : undefined;
    if (!previous || localName(previous.name) !== "p") {
      previous = { name: `${w}p`, attrs: {}, children: [], text: "" };
      parent.children.splice(tableIndex, 0, previous);
    }
    while (pageBreakAdjustment-- > 0) {
      previous.children.push({
        name: `${w}r`,
        attrs: {},
        children: [{ name: `${w}br`, attrs: { [`${w}type`]: "page" }, children: [], text: "" }],
        text: "",
      });
    }
  } else if (parent && pageBreakAdjustment < 0) {
    while (pageBreakAdjustment++ < 0) {
      const tableIndex = parent.children.indexOf(tblEl);
      const previous = tableIndex > 0 ? parent.children[tableIndex - 1] : undefined;
      if (!previous || localName(previous.name) !== "p") break;
      let removed = false;
      for (let runIndex = previous.children.length - 1; runIndex >= 0 && !removed; runIndex--) {
        const run = previous.children[runIndex];
        if (localName(run.name) !== "r") continue;
        for (let childIndex = run.children.length - 1; childIndex >= 0; childIndex--) {
          const runChild = run.children[childIndex];
          if (localName(runChild.name) !== "br" || attr(runChild, "type") !== "page") continue;
          run.children.splice(childIndex, 1);
          if (!run.children.some((childEl) => localName(childEl.name) !== "rPr")) {
            previous.children.splice(runIndex, 1);
          }
          removed = true;
          break;
        }
      }
      if (!removed) break;
      if (!previous.children.some((childEl) => localName(childEl.name) !== "pPr")) {
        parent.children.splice(parent.children.indexOf(previous), 1);
      }
    }
  }
  const keyOf = (name: string): string | undefined =>
    Object.keys(position!.attrs).find((key) => localName(key) === name);
  const set = (name: string, value: string): void => {
    position!.attrs[keyOf(name) ?? `${w}${name}`] = value;
  };
  const remove = (name: string): void => {
    const key = keyOf(name);
    if (key) delete position!.attrs[key];
  };
  for (const name of ["leftFromText", "rightFromText", "topFromText", "bottomFromText"]) {
    if (!keyOf(name)) set(name, "0");
  }
  set("horzAnchor", "page");
  set("vertAnchor", "page");
  set("tblpX", String(Math.round(pxToTwips(xPx))));
  set("tblpY", String(Math.round(pxToTwips(yPx))));
  remove("tblpXSpec");
  remove("tblpYSpec");
  doc.refresh();
  return true;
}

/** Set Word's table text wrapping: None is in document flow; Around floats. */
export function setTableTextWrapping(
  doc: DocxDocument,
  tblEl: XmlElement,
  wrapping: "none" | "around",
  xPx: number,
  yPx: number,
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  if (wrapping === "around") {
    // w:tblpPr is a table property, so a tracked float is a w:tblPrChange like
    // any other. Recorded here rather than inside moveTableTo, which the
    // editor's drag handler also drives, untracked.
    if (meta) recordTableFormatChange(tblEl, meta);
    return moveTableTo(doc, tblEl, xPx, yPx);
  }
  const tblPr = child(tblEl, "tblPr");
  const position = tblPr ? child(tblPr, "tblpPr") : undefined;
  if (!tblPr || !position) return false;
  if (meta) recordTableFormatChange(tblEl, meta);
  tblPr.children.splice(tblPr.children.indexOf(position), 1);
  doc.refresh();
  return true;
}

/** Set the extents of an inline drawing (wp:extent + a:ext), px. */
export function resizeDrawing(
  doc: DocxDocument,
  drawingEl: XmlElement,
  widthPx: number,
  heightPx: number,
): boolean {
  if (localName(drawingEl.name) === "line") {
    const lengthPx = (raw: string): number => {
      const match = /^(-?[\d.]+)\s*(pt|in|px)?$/.exec(raw.trim());
      if (!match) return 0;
      const value = parseFloat(match[1]);
      return match[2] === "pt" ? value * 4 / 3 : match[2] === "in" ? value * 96 : value;
    };
    const point = (raw: string | undefined): [number, number] => {
      const parts = (raw ?? "0,0").split(",");
      return [lengthPx(parts[0] ?? "0"), lengthPx(parts[1] ?? "0")];
    };
    const pt = (value: number): string => `${Math.round(value * 75) / 100}pt`;
    const [x1, y1] = point(drawingEl.attrs.from);
    const [x2, y2] = point(drawingEl.attrs.to);
    const minX = Math.min(x1, x2);
    const minY = Math.min(y1, y2);
    const dx = x2 - x1;
    const dy = y2 - y1;
    const next = {
      x1: dx === 0 ? x1 : dx > 0 ? minX : minX + widthPx,
      y1: dy === 0 ? y1 : dy > 0 ? minY : minY + heightPx,
      x2: dx === 0 ? x2 : dx > 0 ? minX + widthPx : minX,
      y2: dy === 0 ? y2 : dy > 0 ? minY + heightPx : minY,
    };
    drawingEl.attrs.from = `${pt(next.x1)},${pt(next.y1)}`;
    drawingEl.attrs.to = `${pt(next.x2)},${pt(next.y2)}`;
    doc.refresh();
    return true;
  }
  const EMU = 9525;
  const cx = String(Math.max(Math.round(widthPx * EMU), EMU));
  const cy = String(Math.max(Math.round(heightPx * EMU), EMU));
  let lineExtent: XmlElement | undefined;
  const findLineExtent = (node: XmlElement): void => {
    if (lineExtent) return;
    if (localName(node.name) === "spPr") {
      const preset = attr(child(node, "prstGeom"), "prst") ?? "";
      if (preset === "line" || preset.startsWith("straightConnector")) {
        lineExtent = child(child(node, "xfrm"), "ext");
        return;
      }
    }
    for (const childNode of node.children) findLineExtent(childNode);
  };
  findLineExtent(drawingEl);
  const keepVertical = lineExtent
    ? parseInt(lineExtent.attrs[Object.keys(lineExtent.attrs).find((key) => localName(key) === "cx") ?? "cx"] ?? "0", 10) === 0
    : false;
  const keepHorizontal = lineExtent
    ? parseInt(lineExtent.attrs[Object.keys(lineExtent.attrs).find((key) => localName(key) === "cy") ?? "cy"] ?? "0", 10) === 0
    : false;
  let touched = false;
  const setExt = (el: XmlElement) => {
    const cxKey = Object.keys(el.attrs).find((k) => localName(k) === "cx") ?? "cx";
    const cyKey = Object.keys(el.attrs).find((k) => localName(k) === "cy") ?? "cy";
    el.attrs[cxKey] = el === lineExtent && keepVertical ? "0" : cx;
    el.attrs[cyKey] = el === lineExtent && keepHorizontal ? "0" : cy;
    touched = true;
  };
  // Vector groups scale through their group transform: update the outer
  // extents and the group's own a:ext, but leave child shapes' a:ext alone
  // (they are in child-coordinate space; overwriting them with the absolute
  // size corrupts the geometry).
  const walk = (el: XmlElement, insideGroup: boolean) => {
    const ln = localName(el.name);
    if (ln === "extent" || (ln === "ext" && !insideGroup)) setExt(el);
    if (ln === "wgp" || ln === "grpSp") {
      const xfrm = el.children.find((c) => localName(c.name) === "grpSpPr")?.children.find((c) => localName(c.name) === "xfrm");
      const ext = xfrm?.children.find((c) => localName(c.name) === "ext");
      if (ext) setExt(ext);
      return; // children keep their own coordinate space
    }
    for (const c of el.children) walk(c, insideGroup);
  };
  walk(drawingEl, false);
  if (!touched) {
    const find = (el: XmlElement): XmlElement | undefined => {
      if (localName(el.name) === "shape" || localName(el.name) === "rect") return el;
      for (const child of el.children) {
        const match = find(child);
        if (match) return match;
      }
      return undefined;
    };
    const shape = find(drawingEl);
    if (shape) {
      const style = shape.attrs.style ?? "";
      const withoutSize = style
        .split(";")
        .filter((entry) => !/^\s*(?:width|height)\s*:/i.test(entry))
        .filter(Boolean);
      withoutSize.push(`width:${widthPx * 0.75}pt`, `height:${heightPx * 0.75}pt`);
      shape.attrs.style = withoutSize.join(";");
      const setObjectSize = (name: "dxaOrig" | "dyaOrig", value: number): void => {
        const key = Object.keys(drawingEl.attrs).find((candidate) => localName(candidate) === name);
        if (key) drawingEl.attrs[key] = String(Math.round(value * 15));
      };
      setObjectSize("dxaOrig", widthPx);
      setObjectSize("dyaOrig", heightPx);
      touched = true;
    }
  }
  if (touched) doc.refresh();
  return touched;
}

/** Move the run containing a drawing to sit after the run containing targetT. */
export function moveDrawingTo(
  doc: DocxDocument,
  drawingEl: XmlElement,
  targetT: XmlElement,
): boolean {
  // Climb from the drawing to its w:r
  let imgRun: XmlElement | undefined = doc.findParentOf(drawingEl);
  while (imgRun && localName(imgRun.name) !== "r") imgRun = doc.findParentOf(imgRun);
  if (!imgRun) return false;
  const imgParent = doc.findParentOf(imgRun);
  let destRun: XmlElement | undefined = doc.findParentOf(targetT);
  while (destRun && localName(destRun.name) !== "r") destRun = doc.findParentOf(destRun);
  if (!imgParent || !destRun) return false;
  const destParent = doc.findParentOf(destRun);
  if (!destParent) return false;
  if (destRun === imgRun) return false;
  // A shape's editable textbox is part of the drawing run itself. Using that
  // text as the destination would remove imgRun and then insert it into one
  // of its own descendants, detaching the drawing and creating a cycle.
  for (let cur: XmlElement | undefined = targetT; cur; cur = doc.findParentOf(cur)) {
    if (cur === imgRun) return false;
  }
  // Never move a run to a DIFFERENT part (body <-> header/footer): the
  // drawing's r:embed relationship is part-scoped, so the image would render
  // nowhere in the destination part — a drop on the footer silently
  // destroyed the image. Cross-part drops must be rejected so the caller can
  // re-target (or snap back).
  const rootOf = (n: XmlElement): XmlElement => {
    let cur: XmlElement = n;
    for (let p = doc.findParentOf(cur); p; p = doc.findParentOf(p)) cur = p;
    return cur;
  };
  if (rootOf(imgRun) !== rootOf(destRun)) return false;
  imgParent.children.splice(imgParent.children.indexOf(imgRun), 1);
  destParent.children.splice(destParent.children.indexOf(destRun) + 1, 0, imgRun);
  doc.refresh();
  return true;
}

// ---------------------------------------------------------------------------
// Table formatting: borders, style, widths, margins, header rows
// ---------------------------------------------------------------------------

/** OOXML tblPr child order (CT_TblPrBase, the slice we edit). */
const TBLPR_ORDER = [
  "tblStyle",
  "tblpPr",
  "tblOverlap",
  "bidiVisual",
  "tblStyleRowBandSize",
  "tblStyleColBandSize",
  "tblW",
  "jc",
  "tblCellSpacing",
  "tblInd",
  "tblBorders",
  "shd",
  "tblLayout",
  "tblCellMar",
  "tblLook",
  "tblPrChange",
];

/** OOXML trPr child order (CT_TrPrBase, the slice we edit) followed by the
 * three CT_TrPr children that close it: the row's own structural revisions
 * and the formatting record. */
const TRPR_ORDER = [
  "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
  "ins", "del", "trPrChange",
];

function ensureTblPr(tblEl: XmlElement, w: string): XmlElement {
  let tblPr = child(tblEl, "tblPr");
  if (!tblPr) {
    tblPr = { name: `${w}tblPr`, attrs: {}, children: [], text: "" };
    tblEl.children.unshift(tblPr);
  }
  return tblPr;
}

/** The properties element a scoped operation reads, without creating one. */
function scopedProps(ctx: CellContext, scope: "cell" | "table"): XmlElement | undefined {
  return scope === "cell" ? child(ctx.tc, "tcPr") : child(ctx.tbl, "tblPr");
}

/** Record the previous properties of whichever element a scoped operation is
 * about to mutate: the one cell, or the whole table. */
function recordScopedFormatChange(
  ctx: CellContext,
  scope: "cell" | "table",
  meta: RevisionMeta,
): void {
  if (scope === "cell") recordCellFormatChange(ctx.tc, meta);
  else recordTableFormatChange(ctx.tbl, meta);
}

/** Insert `el` into `parent` at the position `order` gives its local name. */
function insertOrdered(parent: XmlElement, el: XmlElement, order: readonly string[]): void {
  const name = localName(el.name);
  const idx = parent.children.findIndex((c) => localName(c.name) === name);
  if (idx !== -1) {
    parent.children[idx] = el;
    return;
  }
  const rank = order.indexOf(name);
  let insertAt = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const r = order.indexOf(localName(parent.children[i].name));
    if (r !== -1 && r > rank) {
      insertAt = i;
      break;
    }
  }
  parent.children.splice(insertAt, 0, el);
}

/** Set (or with attrs=null remove) a tblPr child, keeping schema order. */
function setTblPrChild(
  tblEl: XmlElement,
  w: string,
  name: string,
  attrs: Record<string, string> | null,
): void {
  const tblPr = ensureTblPr(tblEl, w);
  if (attrs === null) {
    const idx = tblPr.children.findIndex((c) => localName(c.name) === name);
    if (idx !== -1) tblPr.children.splice(idx, 1);
    return;
  }
  insertOrdered(tblPr, { name: `${w}${name}`, attrs, children: [], text: "" }, TBLPR_ORDER);
}

// ---------- borders ----------

/** The w:val values w:tblBorders / w:tcBorders edges accept. "none" writes
 * w:val="nil", Word's spelling for an explicitly suppressed rule. */
export type TableBorderStyle =
  | "single"
  | "thick"
  | "double"
  | "dotted"
  | "dashed"
  | "dotDash"
  | "dotDotDash"
  | "thinThickSmallGap"
  | "triple"
  | "wave"
  | "none";

export const TABLE_BORDER_STYLES: readonly TableBorderStyle[] = [
  "single",
  "thick",
  "double",
  "dotted",
  "dashed",
  "dotDash",
  "dotDotDash",
  "thinThickSmallGap",
  "triple",
  "wave",
  "none",
];

/** One border edge. Table scope carries the four sides plus the two inside
 * rules; cell scope carries the four sides plus the two diagonals. Neither
 * carries the other's — the parser reads only these, so writing the rest
 * would emit XML that renders nothing. */
export type TableBorderEdge =
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "insideH"
  | "insideV"
  | "tl2br"
  | "tr2bl";

export const TABLE_SCOPE_EDGES: readonly TableBorderEdge[] = [
  "top",
  "bottom",
  "left",
  "right",
  "insideH",
  "insideV",
];
export const CELL_SCOPE_EDGES: readonly TableBorderEdge[] = [
  "top",
  "bottom",
  "left",
  "right",
  "tl2br",
  "tr2bl",
];

export interface TableBorderSpec {
  style: TableBorderStyle;
  /** Stroke width in eighths of a point (w:sz). Word's UI offers 2..48. */
  sz?: number;
  /** "#RRGGBB", "RRGGBB", or "auto". */
  color?: string;
  /** Gap between rule and content, in points (w:space). */
  space?: number;
}

/** Child order inside CT_TblBorders / CT_TcBorders. */
const BORDER_EDGE_ORDER = [
  "top",
  "start",
  "left",
  "bottom",
  "end",
  "right",
  "insideH",
  "insideV",
  "tl2br",
  "tr2bl",
];

/** Word 2013+ writes w:start/w:end for the side rules of some templates. The
 * parser reads w:left/w:right only, so an edge we set must retire its
 * logical twin or the file would carry two contradictory rules. */
const EDGE_TWIN: Partial<Record<TableBorderEdge, string>> = { left: "start", right: "end" };

function borderElement(w: string, edge: TableBorderEdge, spec: TableBorderSpec): XmlElement {
  if (spec.style === "none") {
    return { name: `${w}${edge}`, attrs: { [`${w}val`]: "nil" }, children: [], text: "" };
  }
  const color = (spec.color ?? "auto").replace(/^#/, "");
  return {
    name: `${w}${edge}`,
    attrs: {
      [`${w}val`]: spec.style,
      [`${w}sz`]: String(Math.round(spec.sz ?? 4)),
      [`${w}space`]: String(Math.round(spec.space ?? 0)),
      [`${w}color`]: color.toLowerCase() === "auto" ? "auto" : color.toUpperCase(),
    },
    children: [],
    text: "",
  };
}

/**
 * Build a w:tblBorders holding the given edges, in schema order.
 *
 * For a caller writing a whole w:tblPr from nothing — a table STYLE
 * definition, which has no document table to read the existing container off
 * — rather than patching edges into a table that already exists.
 */
export function tableBordersElement(
  w: string,
  borders: Partial<Record<TableBorderEdge, TableBorderSpec>>,
): XmlElement {
  const container: XmlElement = { name: `${w}tblBorders`, attrs: {}, children: [], text: "" };
  for (const edge of TABLE_SCOPE_EDGES) {
    const spec = borders[edge];
    if (spec) insertOrdered(container, borderElement(w, edge, spec), BORDER_EDGE_ORDER);
  }
  return container;
}

/**
 * Set or clear border edges on one cell or on the whole table.
 *
 * A `spec` of null REMOVES the edge element, so the edge falls back to what
 * the table style or the table-level rule provides. A spec of
 * `{ style: "none" }` instead writes w:val="nil", which terminates that
 * inheritance — Word's "No Border" button. The renderer already distinguishes
 * the two (an absent edge inherits; a "none" edge paints nothing), so the
 * editing surface has to as well.
 *
 * `target` is any element inside the table; cell scope uses the cell that
 * contains it.
 */
export function setTableBorders(
  doc: DocxDocument,
  target: XmlElement,
  scope: "cell" | "table",
  edges: readonly TableBorderEdge[],
  spec: TableBorderSpec | null,
  meta?: RevisionMeta,
): boolean {
  const ctx = cellContextOf(doc, target);
  if (!ctx) return false;
  const allowed = scope === "cell" ? CELL_SCOPE_EDGES : TABLE_SCOPE_EDGES;
  const wanted = edges.filter((e) => allowed.includes(e));
  if (wanted.length === 0) return false;
  const { w } = ctx;
  const containerName = scope === "cell" ? "tcBorders" : "tblBorders";
  const existing = child(scopedProps(ctx, scope), containerName);
  if (!existing && spec === null) return false; // nothing to clear
  // Recorded before the first mutation, and only once the operation is known
  // to make one — an empty record is a suggestion the reviewer cannot explain.
  if (meta) recordScopedFormatChange(ctx, scope, meta);
  const holder = scope === "cell" ? ensureTcPr(ctx.tc, w) : ensureTblPr(ctx.tbl, w);
  let container = child(holder, containerName);
  if (!container) {
    container = { name: `${w}${containerName}`, attrs: {}, children: [], text: "" };
    if (scope === "cell") setTcPrChildElement(holder, container);
    else insertOrdered(holder, container, TBLPR_ORDER);
  }
  for (const edge of wanted) {
    const twin = EDGE_TWIN[edge];
    for (const name of twin ? [edge, twin] : [edge]) {
      const idx = container.children.findIndex((c) => localName(c.name) === name);
      if (idx !== -1) container.children.splice(idx, 1);
    }
    if (spec !== null) insertOrdered(container, borderElement(w, edge, spec), BORDER_EDGE_ORDER);
  }
  // An emptied container is noise; drop it so the cell inherits cleanly.
  if (container.children.length === 0) {
    holder.children.splice(holder.children.indexOf(container), 1);
  }
  doc.refresh();
  return true;
}

/** Place an already-built element into a tcPr at its schema position. */
function setTcPrChildElement(tcPr: XmlElement, el: XmlElement): void {
  insertOrdered(tcPr, el, TCPR_ORDER);
}

// ---------- table style + tblLook ----------

/** Word's six "Table Style Options" checkboxes. Banding is expressed
 * positively here and inverted into w:noHBand / w:noVBand on write — the
 * negative spelling is a wire detail, and every caller that had to remember
 * the inversion got it wrong at least once. */
export interface TableLookToggles {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  bandedRows: boolean;
  bandedCols: boolean;
}

/** w:tblLook legacy bitmask, kept in sync with the attribute form. */
const TBL_LOOK_BITS = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHBand: 0x0200,
  noVBand: 0x0400,
} as const;

/** The look Word assumes when w:tblLook is absent entirely. */
const DEFAULT_LOOK_TOGGLES: TableLookToggles = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  bandedRows: true,
  bandedCols: true,
};

/** Read the effective toggles of a table, mirroring the parser: explicit
 * attributes win, the legacy hex w:val is the fallback, and an ABSENT
 * w:tblLook means Word's default. */
export function tableLookOf(tblEl: XmlElement): TableLookToggles {
  const tblPr = child(tblEl, "tblPr");
  const look = tblPr ? child(tblPr, "tblLook") : undefined;
  if (!look) return { ...DEFAULT_LOOK_TOGGLES };
  const val = parseInt(attr(look, "val") ?? "0", 16) || 0;
  const flag = (name: keyof typeof TBL_LOOK_BITS): boolean => {
    const a = attr(look, name);
    if (a !== undefined) return a === "1" || a === "true";
    return (val & TBL_LOOK_BITS[name]) !== 0;
  };
  return {
    firstRow: flag("firstRow"),
    lastRow: flag("lastRow"),
    firstColumn: flag("firstColumn"),
    lastColumn: flag("lastColumn"),
    bandedRows: !flag("noHBand"),
    bandedCols: !flag("noVBand"),
  };
}

/**
 * Set some of the six table-style options. Every attribute is rewritten, not
 * just the ones in `patch`: the parser treats a PRESENT w:tblLook with a
 * missing attribute as false, so a partial element would silently clear the
 * toggles it left out. The legacy hex w:val is written alongside for readers
 * that predate the attribute form.
 */
export function setTableLook(
  doc: DocxDocument,
  tblEl: XmlElement,
  patch: Partial<TableLookToggles>,
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const w = prefixOf(tblEl);
  if (meta) recordTableFormatChange(tblEl, meta);
  const next: TableLookToggles = { ...tableLookOf(tblEl), ...patch };
  const bits = {
    firstRow: next.firstRow,
    lastRow: next.lastRow,
    firstColumn: next.firstColumn,
    lastColumn: next.lastColumn,
    noHBand: !next.bandedRows,
    noVBand: !next.bandedCols,
  };
  let val = 0;
  const attrs: Record<string, string> = {};
  for (const [name, on] of Object.entries(bits) as [keyof typeof TBL_LOOK_BITS, boolean][]) {
    attrs[`${w}${name}`] = on ? "1" : "0";
    if (on) val |= TBL_LOOK_BITS[name];
  }
  setTblPrChild(tblEl, w, "tblLook", {
    [`${w}val`]: val.toString(16).toUpperCase().padStart(4, "0"),
    ...attrs,
  });
  doc.refresh();
  return true;
}

/** The table styles defined in the document's styles.xml, id + display name,
 * sorted by name. This is what a style picker offers: applying an id the file
 * does not define renders nothing, so the list is the whole valid domain. */
export function listTableStyles(doc: DocxDocument): { id: string; name: string }[] {
  const out: { id: string; name: string }[] = [];
  for (const st of doc.styles.byId.values()) {
    if (st.type !== "table") continue;
    out.push({ id: st.id, name: st.name ?? st.id });
  }
  out.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  return out;
}

/**
 * Apply a named table style, or with styleId null remove the reference.
 *
 * An id the document's styles.xml does not define is a clean no-op rather
 * than a write: the conditional formatting would resolve to nothing, and
 * every collaborating replica reads the same styles.xml, so they all reject
 * it alike.
 */
export function setTableStyle(
  doc: DocxDocument,
  tblEl: XmlElement,
  styleId: string | null,
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const w = prefixOf(tblEl);
  if (styleId === null) {
    const tblPr = child(tblEl, "tblPr");
    if (!tblPr || !child(tblPr, "tblStyle")) return false;
    if (meta) recordTableFormatChange(tblEl, meta);
    setTblPrChild(tblEl, w, "tblStyle", null);
    doc.refresh();
    return true;
  }
  const style = doc.styles.byId.get(styleId);
  if (!style || style.type !== "table") return false;
  if (meta) recordTableFormatChange(tblEl, meta);
  setTblPrChild(tblEl, w, "tblStyle", { [`${w}val`]: styleId });
  doc.refresh();
  return true;
}

// ---------- numeric widths and autofit ----------

const PT_PER_TWIP = 1 / 20;

/**
 * Set the table's preferred width.
 *
 * - `pt` writes w:tblW type="dxa" (twips), the width Word's "Preferred
 *   width" box holds.
 * - `pct` writes fiftieths of a percent, the form Word emits.
 * - `auto` writes type="auto", which is what makes an autofit table size to
 *   its content instead of to a declared target.
 */
export function setTableWidth(
  doc: DocxDocument,
  tblEl: XmlElement,
  unit: "pt" | "pct" | "auto",
  value = 0,
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const w = prefixOf(tblEl);
  if (meta) recordTableFormatChange(tblEl, meta);
  const attrs =
    unit === "auto"
      ? { [`${w}w`]: "0", [`${w}type`]: "auto" }
      : unit === "pct"
        ? { [`${w}w`]: String(Math.round(value * 50)), [`${w}type`]: "pct" }
        : { [`${w}w`]: String(Math.round(value / PT_PER_TWIP)), [`${w}type`]: "dxa" };
  setTblPrChild(tblEl, w, "tblW", attrs);
  doc.refresh();
  return true;
}

/**
 * Set one grid column to an exact width in points, leaving the others alone.
 * The table's total width grows or shrinks by the delta — this is Word's
 * "Table Properties > Column > Preferred width", not a boundary drag, which
 * trades width with the neighbouring column (see resizeTableColumn).
 *
 * Every cell is re-stamped with a matching tcW, and a dxa tblW is re-totalled,
 * so the layout trusts the grid it is being handed.
 */
export function setTableColumnWidth(
  doc: DocxDocument,
  tblEl: XmlElement,
  colIdx: number,
  widthPt: number,
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const cols = gridColsOf(tblEl);
  const col = cols[colIdx];
  if (!col) return false;
  if (meta) recordColumnWidthChange(tblEl, meta);
  const twips = Math.max(Math.round(widthPt / PT_PER_TWIP), MIN_COL_TWIPS);
  const key = Object.keys(col.attrs).find((k) => localName(k) === "w") ?? prefixOf(col) + "w";
  col.attrs[key] = String(twips);
  stampCellWidthsFromGrid(tblEl, meta);
  retotalFixedTableWidth(tblEl, cols);
  doc.refresh();
  return true;
}

/**
 * Record everything a column-width change touches. The width is spread across
 * three places — the w:tblGrid, the w:tblW total, and a w:tcW on every cell —
 * so a suggestion that recorded only one of them would restore a table whose
 * three answers disagree. The cells are recorded as they are restamped (see
 * stampCellWidthsFromGrid); the grid and the table properties are recorded
 * here, together, because w:tblGridChange carries no author of its own and is
 * reviewed as part of the table's w:tblPrChange.
 */
function recordColumnWidthChange(tblEl: XmlElement, meta: RevisionMeta): void {
  recordTableFormatChange(tblEl, meta);
  recordTableGridChange(tblEl, meta);
}

/** Re-total a dxa w:tblW from the grid. A pct or auto width is a different
 * intent and is left alone. */
function retotalFixedTableWidth(tblEl: XmlElement, cols: XmlElement[]): void {
  const tblPr = child(tblEl, "tblPr");
  const tblW = tblPr ? child(tblPr, "tblW") : undefined;
  if (!tblW) return;
  const typeKey = Object.keys(tblW.attrs).find((k) => localName(k) === "type");
  if (typeKey && tblW.attrs[typeKey] !== "dxa") return;
  const wKey = Object.keys(tblW.attrs).find((k) => localName(k) === "w") ?? prefixOf(tblW) + "w";
  tblW.attrs[wKey] = String(cols.reduce((a, c) => a + widthAttrOf(c), 0));
  if (typeKey) tblW.attrs[typeKey] = "dxa";
}

/**
 * Switch a table between Word's "Fixed Column Width" and "AutoFit to
 * Contents".
 *
 * The two directions are NOT symmetric, and the asymmetry is the whole point:
 *
 * - **autofit -> fixed** must FREEZE what is on screen. An autofit table's
 *   rendered columns come from measuring content, and its grid may say
 *   something else entirely, so writing w:tblLayout="fixed" alone would snap
 *   the columns to a grid the user never saw. `renderedWidths` (the caller's
 *   measured columns, in px) is written into the grid, totalled into a dxa
 *   tblW, and stamped onto every cell as tcW. Without it the existing grid is
 *   the best available answer and is committed as-is.
 * - **fixed -> autofit** must RECOMPUTE. Layout trusts a grid only when the
 *   cells declare widths, so the tcW stamps have to go; the tblW target has
 *   to become auto, or the recomputed columns would still be scaled to the
 *   frozen total. The grid itself stays as a harmless hint.
 *
 * `renderedWidths` rides as data rather than being measured inside, so every
 * collaborating replica applies the identical mutation (the same contract
 * resizeTableColumn uses).
 */
export function setTableLayoutMode(
  doc: DocxDocument,
  tblEl: XmlElement,
  layout: "fixed" | "autofit",
  renderedWidths?: number[],
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const w = prefixOf(tblEl);
  const cols = gridColsOf(tblEl);
  if (meta) recordColumnWidthChange(tblEl, meta);
  if (layout === "fixed") {
    if (renderedWidths && renderedWidths.length === cols.length && cols.length > 0) {
      cols.forEach((c, i) => {
        const key = Object.keys(c.attrs).find((k) => localName(k) === "w") ?? prefixOf(c) + "w";
        c.attrs[key] = String(Math.max(Math.round(pxToTwips(renderedWidths[i])), MIN_COL_TWIPS));
      });
    }
    const total = cols.reduce((a, c) => a + widthAttrOf(c), 0);
    if (total > 0) {
      setTblPrChild(tblEl, w, "tblW", { [`${w}w`]: String(total), [`${w}type`]: "dxa" });
    }
    stampCellWidthsFromGrid(tblEl, meta);
    setTblPrChild(tblEl, w, "tblLayout", { [`${w}type`]: "fixed" });
  } else {
    for (const tr of rowsOf(tblEl)) {
      for (const tc of cellsOf(tr)) {
        if (meta) recordCellFormatChange(tc, meta);
        setTcPrChild(tc, prefixOf(tc), "tcW", null);
      }
    }
    setTblPrChild(tblEl, w, "tblW", { [`${w}w`]: "0", [`${w}type`]: "auto" });
    setTblPrChild(tblEl, w, "tblLayout", { [`${w}type`]: "autofit" });
  }
  doc.refresh();
  return true;
}

// ---------- cell margins ----------

/** Cell padding in points. An omitted side is left as it was; pass null to
 * setTableCellMargins to drop the whole override. */
export interface CellMarginsPt {
  top?: number;
  left?: number;
  bottom?: number;
  right?: number;
}

/** Child order inside CT_TblCellMar / CT_TcMar. */
const CELL_MARGIN_ORDER = ["top", "start", "left", "bottom", "end", "right"];

/** w:start / w:end are the modern spelling of the side insets; the parser
 * reads both, so a side we write must retire its twin. */
const MARGIN_TWIN: Record<string, string> = { left: "start", right: "end" };

/**
 * Set the table's default cell margins (scope "table", w:tblCellMar) or one
 * cell's override (scope "cell", w:tcMar). Sides are in points; a null patch
 * removes the element so the cell falls back to the table default, and the
 * table default falls back to the style's.
 */
export function setTableCellMargins(
  doc: DocxDocument,
  target: XmlElement,
  scope: "cell" | "table",
  margins: CellMarginsPt | null,
  meta?: RevisionMeta,
): boolean {
  const ctx = cellContextOf(doc, target);
  if (!ctx) return false;
  const { w } = ctx;
  const containerName = scope === "cell" ? "tcMar" : "tblCellMar";
  const ensureHolder = () => (scope === "cell" ? ensureTcPr(ctx.tc, w) : ensureTblPr(ctx.tbl, w));
  if (margins === null) {
    const existing = child(scopedProps(ctx, scope), containerName);
    if (!existing) return false;
    if (meta) recordScopedFormatChange(ctx, scope, meta);
    const holder = ensureHolder();
    holder.children.splice(holder.children.indexOf(existing), 1);
    doc.refresh();
    return true;
  }
  const sides = (["top", "left", "bottom", "right"] as const).filter(
    (side) => margins[side] !== undefined,
  );
  if (sides.length === 0) return false;
  if (meta) recordScopedFormatChange(ctx, scope, meta);
  const holder = ensureHolder();
  let container = child(holder, containerName);
  if (!container) {
    container = { name: `${w}${containerName}`, attrs: {}, children: [], text: "" };
    if (scope === "cell") setTcPrChildElement(holder, container);
    else insertOrdered(holder, container, TBLPR_ORDER);
  }
  for (const side of sides) {
    const twin = MARGIN_TWIN[side];
    if (twin) {
      const idx = container.children.findIndex((c) => localName(c.name) === twin);
      if (idx !== -1) container.children.splice(idx, 1);
    }
    const twips = String(Math.max(0, Math.round(margins[side]! / PT_PER_TWIP)));
    insertOrdered(
      container,
      {
        name: `${w}${side}`,
        attrs: { [`${w}w`]: twips, [`${w}type`]: "dxa" },
        children: [],
        text: "",
      },
      CELL_MARGIN_ORDER,
    );
  }
  doc.refresh();
  return true;
}

// ---------- repeat header rows ----------

/**
 * Mark the first `count` rows as the table's header band, so they repeat at
 * the top of every page the table continues onto. Word's header is a BAND,
 * not a single row: rows 1..N all carry w:tblHeader, and the band stops at
 * the first row that does not.
 *
 * The whole leading run is rewritten each call — rows past `count` have their
 * w:tblHeader removed — so the operation is idempotent and cannot leave a
 * stranded header row below a non-header one, which Word ignores.
 */
export function setTableHeaderRows(
  doc: DocxDocument,
  tblEl: XmlElement,
  count: number,
  meta?: RevisionMeta,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  const rows = rowsOf(tblEl);
  if (rows.length === 0) return false;
  const wanted = Math.max(0, Math.min(Math.floor(count), rows.length));
  let changed = false;
  rows.forEach((tr, i) => {
    const w = prefixOf(tr);
    const present = onOff(child(child(tr, "trPr"), "tblHeader")) === true;
    if (present === i < wanted) return;
    // Only the rows whose band membership actually moves are recorded; the
    // rest are untouched and have nothing for a reviewer to restore.
    if (meta) recordRowFormatChange(tr, meta);
    let trPr = tr.children.find((c) => localName(c.name) === "trPr");
    if (i < wanted) {
      if (!trPr) {
        trPr = { name: `${w}trPr`, attrs: {}, children: [], text: "" };
        tr.children.unshift(trPr);
      }
      insertOrdered(trPr, { name: `${w}tblHeader`, attrs: {}, children: [], text: "" }, TRPR_ORDER);
    } else if (trPr) {
      const idx = trPr.children.findIndex((c) => localName(c.name) === "tblHeader");
      if (idx !== -1) trPr.children.splice(idx, 1);
      if (trPr.children.length === 0) tr.children.splice(tr.children.indexOf(trPr), 1);
    }
    changed = true;
  });
  if (changed) doc.refresh();
  return changed;
}

// ---------- sort rows ----------

/** All text under a cell, in document order. */
function cellText(tc: XmlElement): string {
  let out = "";
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "t") out += e.text;
    e.children.forEach(walk);
  };
  contentOf(tc).forEach(walk);
  return out;
}

/** True when any cell of the table spans columns or rows. Sorting a table
 * with merged cells would tear the merged regions apart, so — like Word —
 * the sort refuses such tables. */
function hasMergedCells(tblEl: XmlElement): boolean {
  for (const tr of rowsOf(tblEl)) {
    for (const tc of cellsOf(tr)) {
      const tcPr = child(tc, "tcPr");
      if (tcPr && (child(tcPr, "gridSpan") || child(tcPr, "vMerge"))) return true;
    }
  }
  return false;
}

/**
 * Sort the table's body rows by the text of one grid column, ascending or
 * descending, comparing as text or as numbers. Rows keep their element
 * identity (they only reorder), so ids, bookmarks, and comments inside them
 * survive.
 *
 * The repeating header band (w:tblHeader rows) always stays in place;
 * `hasHeader` additionally pins the first row, Word's "my list has a header
 * row". Equal keys keep document order (Array.prototype.sort is stable).
 *
 * The comparisons are deliberately locale-free (code-unit order over
 * lower-cased text; parseFloat for numbers): a collab replica must sort
 * identically on every host, so no Intl machinery may be consulted at apply
 * (plan doc 05).
 */
export function sortTableRows(
  doc: DocxDocument,
  tblEl: XmlElement,
  colIdx: number,
  order: "asc" | "desc",
  compare: "text" | "number",
  hasHeader = false,
): boolean {
  if (localName(tblEl.name) !== "tbl") return false;
  if (hasMergedCells(tblEl)) return false;
  const rows = rowsOf(tblEl);
  let pinned = 0;
  while (pinned < rows.length && onOff(child(child(rows[pinned], "trPr"), "tblHeader")) === true) pinned++;
  if (hasHeader && pinned === 0) pinned = 1;
  const body = rows.slice(pinned);
  if (body.length < 2) return false;

  const keyOf = (tr: XmlElement): string => {
    const tc = cellAtGridCol(tr, colIdx);
    return tc ? cellText(tc).trim() : "";
  };
  const numeric = (s: string): number => parseFloat(s.replace(/[^0-9.eE+-]/g, ""));
  const compareRows = (a: XmlElement, b: XmlElement): number => {
    const ka = keyOf(a);
    const kb = keyOf(b);
    let d: number;
    if (compare === "number") {
      const na = numeric(ka);
      const nb = numeric(kb);
      // Non-numbers sort after every number, in both directions.
      d = Number.isNaN(na) && Number.isNaN(nb) ? 0 : Number.isNaN(na) ? 1 : Number.isNaN(nb) ? -1 : na - nb;
      if (Number.isNaN(na) !== Number.isNaN(nb)) return d;
    } else {
      const la = ka.toLowerCase();
      const lb = kb.toLowerCase();
      d = la < lb ? -1 : la > lb ? 1 : 0;
    }
    return order === "asc" ? d : -d;
  };

  const sorted = [...body].sort(compareRows);
  if (sorted.every((tr, i) => tr === body[i])) return false; // already in order
  // Reorder in place: the sortable rows take each other's slots in
  // tbl.children; everything else (tblPr, tblGrid, header rows) stays put.
  const slots = body.map((tr) => tblEl.children.indexOf(tr));
  sorted.forEach((tr, i) => {
    tblEl.children[slots[i]] = tr;
  });
  doc.refresh();
  return true;
}

// ---------- reading a table's current properties ----------

/**
 * What a Table Properties dialog needs to show the table around the caret,
 * in points throughout so no caller handles twips.
 *
 * Read-only, and separate from the setters because Word's dialog PREFILLS:
 * a width box that opens empty asks the user to retype a number the document
 * already carries, and one that opens on a guess is worse than empty.
 */
export interface TablePropertiesPt {
  /** Columns in the table's grid. */
  columnCount: number;
  /** Grid column the caret sits in, counting gridSpans to its left. */
  columnIdx: number;
  /** w:tblW. An absent element IS "auto" — that is what the schema's default
   * says — so the unit is always answered. */
  width: { unit: "pt" | "pct" | "auto"; value: number };
  /** Each grid column's width. */
  columnWidthsPt: number[];
  /** The table's default cell margins (w:tblCellMar). A side the table does
   * not declare is ABSENT rather than zero, so a caller can tell an inherited
   * margin from a suppressed one. */
  cellMargins: CellMarginsPt;
  /** Rows in the repeating header band (see setTableHeaderRows). */
  headerRows: number;
}

/** Read the table around `target`, or undefined when it is outside one — the
 * same "no table here" answer cellShadingAt gives. */
export function readTableProperties(
  doc: DocxDocument,
  target: XmlElement,
): TablePropertiesPt | undefined {
  const ctx = cellContextOf(doc, target);
  if (!ctx) return undefined;
  const cols = gridColsOf(ctx.tbl);
  const tblPr = child(ctx.tbl, "tblPr");
  const tblW = tblPr ? child(tblPr, "tblW") : undefined;
  const type = tblW ? attr(tblW, "type") ?? "dxa" : "auto";
  const raw = tblW ? widthAttrOf(tblW) : 0;
  const width =
    type === "pct"
      ? { unit: "pct" as const, value: raw / 50 }
      : type === "dxa"
        ? { unit: "pt" as const, value: raw * PT_PER_TWIP }
        : { unit: "auto" as const, value: 0 };

  const marEl = tblPr ? child(tblPr, "tblCellMar") : undefined;
  const cellMargins: CellMarginsPt = {};
  if (marEl) {
    for (const side of ["top", "left", "bottom", "right"] as const) {
      const twin = MARGIN_TWIN[side];
      const el = child(marEl, side) ?? (twin ? child(marEl, twin) : undefined);
      if (el) cellMargins[side] = widthAttrOf(el) * PT_PER_TWIP;
    }
  }

  let headerRows = 0;
  for (const tr of rowsOf(ctx.tbl)) {
    if (onOff(child(child(tr, "trPr"), "tblHeader")) !== true) break;
    headerRows++;
  }

  return {
    columnCount: cols.length,
    columnIdx: gridColOf(ctx.tr, ctx.tc),
    width,
    columnWidthsPt: cols.map((c) => widthAttrOf(c) * PT_PER_TWIP),
    cellMargins,
    headerRows,
  };
}
