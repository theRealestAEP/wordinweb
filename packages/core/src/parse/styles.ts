import { XmlElement, attr, child, children, childVal, intAttr } from "../xml.js";
import { ParaProps, RunProps, Style, Styles, TableLook, TableProps, TableCondFormat, TableCondType } from "../model.js";
import { twipsToPx } from "../units.js";
import {
  ParseContext,
  mergeParaProps,
  mergeRunProps,
  parseBorder,
  parseParaProps,
  parseRunProps,
  parseShading,
} from "./properties.js";

export function parseStyles(root: XmlElement | undefined, ctx: ParseContext): Styles {
  const styles: Styles = {
    byId: new Map(),
    defaultRPr: {},
    defaultPPr: {},
  };
  if (!root) return styles;

  const docDefaults = child(root, "docDefaults");
  if (docDefaults) {
    const rPrDefault = child(child(docDefaults, "rPrDefault"), "rPr");
    styles.defaultRPr = parseRunProps(rPrDefault, ctx);
    const pPrDefault = child(child(docDefaults, "pPrDefault"), "pPr");
    styles.defaultPPr = parseParaProps(pPrDefault, ctx);
  }
  // Word's implicit defaults when docDefaults omit them (10pt = 13.33px)
  if (styles.defaultRPr.size === undefined) styles.defaultRPr.size = (10 * 4) / 3;
  if (!styles.defaultRPr.font) styles.defaultRPr.font = ctx.theme?.minorFont ?? "Calibri";

  for (const s of children(root, "style")) {
    const id = attr(s, "styleId");
    const type = attr(s, "type") as Style["type"] | undefined;
    if (!id || !type) continue;
    const style: Style = {
      id,
      type,
      name: childVal(s, "name"),
      basedOn: childVal(s, "basedOn"),
      isDefault: attr(s, "default") === "1" || attr(s, "default") === "true",
    };
    const pPr = child(s, "pPr");
    if (pPr) style.pPr = parseParaProps(pPr, ctx);
    const rPr = child(s, "rPr");
    if (rPr) style.rPr = parseRunProps(rPr, ctx);
    const tblPr = child(s, "tblPr");
    if (tblPr && type === "table") {
      const tp: TableProps = {};
      const cellMar = child(tblPr, "tblCellMar");
      if (cellMar) {
        const margins: { top?: number; right?: number; bottom?: number; left?: number } = {};
        for (const side of ["top", "right", "bottom", "left"] as const) {
          const m =
            child(cellMar, side) ??
            (side === "right" ? child(cellMar, "end") : side === "left" ? child(cellMar, "start") : undefined);
          if (m && attr(m, "type") !== "pct") {
            const w = intAttr(m, "w");
            if (w !== undefined) margins[side] = twipsToPx(w);
          }
        }
        tp.cellMargins = margins;
      }
      // Table styles carry the grid borders (e.g. the built-in "Table Grid"
      // style, referenced by tblStyle with no direct tblBorders) — resolve
      // them so styled tables render their cell grid.
      const borders = child(tblPr, "tblBorders");
      if (borders) {
        tp.borders = {
          top: parseBorder(child(borders, "top"), ctx),
          bottom: parseBorder(child(borders, "bottom"), ctx),
          left: parseBorder(child(borders, "left"), ctx),
          right: parseBorder(child(borders, "right"), ctx),
          insideH: parseBorder(child(borders, "insideH"), ctx),
          insideV: parseBorder(child(borders, "insideV"), ctx),
        };
      }
      style.tblPr = tp;
      const rbs = intAttr(child(tblPr, "tblStyleRowBandSize"), "val");
      if (rbs !== undefined) style.rowBandSize = rbs;
      const cbs = intAttr(child(tblPr, "tblStyleColBandSize"), "val");
      if (cbs !== undefined) style.colBandSize = cbs;
      // Conditional formatting (w:tblStylePr): banding fills, header/last-row
      // borders and bold — resolved per cell against tblLook at paint time.
      const cond = new Map<TableCondType, TableCondFormat>();
      for (const cp of children(s, "tblStylePr")) {
        const type = attr(cp, "type") as TableCondType | undefined;
        if (!type) continue;
        const cf: TableCondFormat = {};
        const tcPr = child(cp, "tcPr");
        if (tcPr) {
          const shd = parseShading(child(tcPr, "shd"), ctx);
          if (shd) cf.shd = shd;
          const b = child(tcPr, "tcBorders");
          if (b) {
            cf.borders = {
              top: parseBorder(child(b, "top"), ctx),
              bottom: parseBorder(child(b, "bottom"), ctx),
              left: parseBorder(child(b, "left"), ctx),
              right: parseBorder(child(b, "right"), ctx),
              insideH: parseBorder(child(b, "insideH"), ctx),
              insideV: parseBorder(child(b, "insideV"), ctx),
            };
          }
        }
        const rp = child(cp, "rPr");
        if (rp) {
          cf.rPr = parseRunProps(rp, ctx);
          if (cf.rPr.bold) cf.bold = true;
        }
        cond.set(type, cf);
      }
      if (cond.size > 0) style.condFormats = cond;
    }
    styles.byId.set(id, style);
    if (style.isDefault && type === "paragraph") styles.defaultParagraphStyle = id;
  }
  return styles;
}

const MAX_CHAIN = 20;

/**
 * Effective paragraph props from the style chain (no direct formatting).
 * When `includeDefaults` is false, only the style chain's own contribution is
 * returned (docDefaults omitted) so a caller can interpose another layer
 * (e.g. a table style's pPr) beneath the paragraph style.
 */
export function resolveParagraphStyleChain(
  styles: Styles,
  styleId: string | undefined,
  includeDefaults = true,
): {
  pPr: ParaProps;
  rPr: RunProps;
} {
  const pChain: Style[] = [];
  let cur = styleId ?? styles.defaultParagraphStyle;
  let guard = 0;
  while (cur && guard++ < MAX_CHAIN) {
    const s = styles.byId.get(cur);
    if (!s) break;
    pChain.unshift(s);
    cur = s.basedOn;
  }
  let pPr: ParaProps = includeDefaults ? { ...styles.defaultPPr } : {};
  let rPr: RunProps = includeDefaults ? { ...styles.defaultRPr } : {};
  for (const s of pChain) {
    if (s.pPr) pPr = mergeParaProps(pPr, s.pPr);
    if (s.rPr) rPr = mergeRunProps(rPr, s.rPr);
  }
  return { pPr, rPr };
}

/** Merged pPr/rPr contributed by a table style's own basedOn chain (no docDefaults). */
export function resolveTableStyleProps(styles: Styles, styleId: string): { pPr?: ParaProps; rPr?: RunProps } {
  const chain: Style[] = [];
  let cur: string | undefined = styleId;
  let guard = 0;
  while (cur && guard++ < MAX_CHAIN) {
    const s = styles.byId.get(cur);
    if (!s) break;
    chain.unshift(s);
    cur = s.basedOn;
  }
  let pPr: ParaProps | undefined;
  let rPr: RunProps | undefined;
  for (const s of chain) {
    if (s.pPr) pPr = pPr ? mergeParaProps(pPr, s.pPr) : { ...s.pPr };
    if (s.rPr) rPr = rPr ? mergeRunProps(rPr, s.rPr) : { ...s.rPr };
  }
  return { pPr, rPr };
}

/**
 * Conditional table formats (w:tblStylePr) and band sizes resolved through a
 * table style's basedOn chain (derived style wins over its base). Returns a map
 * keyed by conditional type; each entry is used per cell against the table's
 * tblLook + row/column position at paint time.
 */
export function resolveTableConditional(
  styles: Styles,
  styleId: string | undefined,
): { formats: Map<TableCondType, TableCondFormat>; rowBandSize: number; colBandSize: number } {
  const chain: Style[] = [];
  let cur = styleId;
  let guard = 0;
  while (cur && guard++ < MAX_CHAIN) {
    const st = styles.byId.get(cur);
    if (!st) break;
    chain.unshift(st);
    cur = st.basedOn;
  }
  const formats = new Map<TableCondType, TableCondFormat>();
  // Word only applies banded conditional formats when the style chain declares
  // an explicit w:tblStyleRowBandSize / ColBandSize (every built-in banded
  // style writes w:val="1"). A chain that never declares one gets NO banding,
  // despite ECMA-376's "default 1" — verified against Word output for a custom
  // style with band1Horz but no band size. 0 means "banding disabled" here.
  let rowBandSize = 0;
  let colBandSize = 0;
  for (const st of chain) {
    if (st.rowBandSize !== undefined) rowBandSize = st.rowBandSize;
    if (st.colBandSize !== undefined) colBandSize = st.colBandSize;
    if (st.condFormats) {
      for (const [type, cf] of st.condFormats) {
        const prev = formats.get(type);
        const merged = prev ? { ...prev, ...cf } : { ...cf };
        if (prev?.rPr && cf.rPr) merged.rPr = mergeRunProps(prev.rPr, cf.rPr);
        formats.set(type, merged);
      }
    }
  }
  return { formats, rowBandSize, colBandSize };
}

/** tblLook Word assumes when a table carries no w:tblLook element. */
export const DEFAULT_TBL_LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  noHBand: false,
  noVBand: false,
};

/**
 * The w:tblStylePr conditional types that apply to a cell, in layering order
 * (low→high precedence: wholeTable < banding < first/last col < first/last row
 * < corners), gated by the table's tblLook. Band sizes of 0 mean the style
 * chain never declared w:tblStyleRow/ColBandSize — Word skips banding then.
 */
export function tableCondOrder(
  look: TableLook,
  rowIdx: number,
  nRows: number,
  colStart: number,
  colSpan: number,
  nCols: number,
  rowBandSize: number,
  colBandSize: number,
): TableCondType[] {
  const isFirstRow = look.firstRow && rowIdx === 0;
  const isLastRow = look.lastRow && rowIdx === nRows - 1;
  const isFirstCol = look.firstColumn && colStart === 0;
  const isLastCol = look.lastColumn && colStart + colSpan === nCols;

  const order: TableCondType[] = ["wholeTable"];
  if (!look.noVBand && !isFirstCol && !isLastCol && colBandSize > 0) {
    const bandCol = colStart - (look.firstColumn ? 1 : 0);
    if (bandCol >= 0) {
      order.push(Math.floor(bandCol / colBandSize) % 2 === 0 ? "band1Vert" : "band2Vert");
    }
  }
  if (!look.noHBand && !isFirstRow && !isLastRow && rowBandSize > 0) {
    const bandRow = rowIdx - (look.firstRow ? 1 : 0);
    if (bandRow >= 0) {
      order.push(Math.floor(bandRow / rowBandSize) % 2 === 0 ? "band1Horz" : "band2Horz");
    }
  }
  if (isFirstCol) order.push("firstCol");
  if (isLastCol) order.push("lastCol");
  if (isFirstRow) order.push("firstRow");
  if (isLastRow) order.push("lastRow");
  if (isFirstRow && isFirstCol) order.push("nwCell");
  if (isFirstRow && isLastCol) order.push("neCell");
  if (isLastRow && isFirstCol) order.push("swCell");
  if (isLastRow && isLastCol) order.push("seCell");
  return order;
}

/** Effective run props contributed by a character style chain. */
export function resolveCharacterStyleChain(styles: Styles, styleId: string | undefined): RunProps {
  const chain: Style[] = [];
  let cur = styleId;
  let guard = 0;
  while (cur && guard++ < MAX_CHAIN) {
    const s = styles.byId.get(cur);
    if (!s) break;
    chain.unshift(s);
    cur = s.basedOn;
  }
  let rPr: RunProps = {};
  for (const s of chain) {
    if (s.rPr) rPr = mergeRunProps(rPr, s.rPr);
  }
  return rPr;
}
