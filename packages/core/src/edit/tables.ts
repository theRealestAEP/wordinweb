import { DocxDocument } from "../docx.js";
import { XmlElement, cloneXml, localName, child } from "../xml.js";
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
  | "deleteTable";

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

const MIN_COL_TWIPS = 144; // 0.1"

/**
 * Drag-resize a table column boundary by deltaPx. Interior boundaries trade
 * width between the two adjacent columns (table width constant, like Word);
 * the rightmost boundary grows/shrinks the table.
 */
export function resizeTableColumn(
  doc: DocxDocument,
  tblEl: XmlElement,
  boundary: number,
  deltaPx: number,
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

  // Keep per-cell tcW (dxa) in sync when the table has no merged cells.
  if (!hasSpans(tblEl)) {
    for (const tr of rowsOf(tblEl)) {
      const cells = cellsOf(tr);
      cells.forEach((tc, i) => {
        const tcPr = child(tc, "tcPr");
        const tcW = tcPr ? child(tcPr, "tcW") : undefined;
        if (tcW && cols[i]) {
          const typeKey = Object.keys(tcW.attrs).find((k) => localName(k) === "type");
          if (!typeKey || tcW.attrs[typeKey] === "dxa") {
            const wKey = Object.keys(tcW.attrs).find((k) => localName(k) === "w") ?? prefixOf(tcW) + "w";
            tcW.attrs[wKey] = String(widthOf(cols[i]));
          }
        }
      });
    }
  }
  doc.refresh();
  return true;
}
