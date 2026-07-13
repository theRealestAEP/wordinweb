import { DocxDocument } from "../docx.js";
import { XmlElement, attr, child, cloneXml, localName } from "../xml.js";
import { pxToTwips } from "../units.js";

/**
 * Table manipulation: add/remove rows and columns relative to the cell
 * containing the caret. Mutates source XML; callers checkpoint + relayout.
 *
 * v1 scope: column operations refuse tables using gridSpan (merged cells)
 * rather than corrupt them.
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
  | { kind: "cellVAlign"; v: "top" | "center" | "bottom" };

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

/** OOXML tcPr child order (the slice we edit). */
const TCPR_ORDER = ["tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "tcMar", "vAlign"];

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

export function applyTableOp(doc: DocxDocument, target: XmlElement, op: TableOp): boolean {
  if (typeof op === "object" || op === "mergeRight" || op === "mergeDown" || op === "splitCell") {
    const ctx = cellContextOf(doc, target);
    if (!ctx) return false;
    const { tbl, tr, tc, w } = ctx;
    if (typeof op === "object") {
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

  // Keep per-cell tcW (dxa) in sync when the table has no merged cells,
  // creating tcPr/tcW when absent: Word stamps explicit cell widths on drag,
  // and a grid without tcW is otherwise treated as autofit (Word semantics).
  if (!hasSpans(tblEl)) {
    for (const tr of rowsOf(tblEl)) {
      const cells = cellsOf(tr);
      cells.forEach((tc, i) => {
        if (!cols[i]) return;
        const w = prefixOf(tc);
        let tcPr = child(tc, "tcPr");
        if (!tcPr) {
          tcPr = { name: `${w}tcPr`, attrs: {}, children: [], text: "" };
          tc.children.unshift(tcPr);
        }
        let tcW = child(tcPr, "tcW");
        if (!tcW) {
          tcW = { name: `${w}tcW`, attrs: { [`${w}type`]: "dxa" }, children: [], text: "" };
          tcPr.children.unshift(tcW);
        }
        const typeKey = Object.keys(tcW.attrs).find((k) => localName(k) === "type");
        if (!typeKey || tcW.attrs[typeKey] === "dxa") {
          const wKey = Object.keys(tcW.attrs).find((k) => localName(k) === "w") ?? prefixOf(tcW) + "w";
          tcW.attrs[wKey] = String(widthOf(cols[i]));
        }
      });
    }
  }
  doc.refresh();
  return true;
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

/** Set the extents of an inline drawing (wp:extent + a:ext), px. */
export function resizeDrawing(
  doc: DocxDocument,
  drawingEl: XmlElement,
  widthPx: number,
  heightPx: number,
): boolean {
  const EMU = 9525;
  const cx = String(Math.max(Math.round(widthPx * EMU), EMU));
  const cy = String(Math.max(Math.round(heightPx * EMU), EMU));
  let touched = false;
  const setExt = (el: XmlElement) => {
    const cxKey = Object.keys(el.attrs).find((k) => localName(k) === "cx") ?? "cx";
    const cyKey = Object.keys(el.attrs).find((k) => localName(k) === "cy") ?? "cy";
    el.attrs[cxKey] = cx;
    el.attrs[cyKey] = cy;
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
