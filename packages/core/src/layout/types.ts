import { Border, Run, RunProps } from "../model.js";
import type { XmlElement } from "../xml.js";

/** Maps a rendered text item back to its source XML for editing. */
export interface TextSource {
  run: Run;
  /** Source w:t; null when the text is synthetic (fields, symbols). */
  t: XmlElement | null;
  /** Char offset of this item's first character within t's text. */
  offset: number;
}

/**
 * Layout output: pages of absolutely positioned primitives (px, page-relative).
 * The renderer maps these 1:1 to DOM/canvas/SVG — no further layout happens
 * downstream, which is what guarantees pagination fidelity.
 */

export interface FontSpec {
  family: string;
  size: number;
  bold: boolean;
  italic: boolean;
}

export interface TextItem {
  kind: "text";
  x: number;
  /** Baseline y. */
  baseline: number;
  width: number;
  text: string;
  props: RunProps;
  font: FontSpec;
  /** Line box for selection/highlight backgrounds. */
  lineTop: number;
  lineHeight: number;
  /** Exact glyph box for baseline-shifted runs (superscript/subscript):
   * the renderer anchors these instead of bottoming on the line box. */
  glyphTop?: number;
  glyphBoxH?: number;
  /** Source m:oMath element when this text is a piece of an equation. */
  mathSrc?: XmlElement;
  href?: string;
  /** Present for editable text (absent on numbering labels etc.). */
  src?: TextSource;
}

export interface RectItem {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
}

export interface LineEdgeItem {
  kind: "edge";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  border: Border;
}

export interface ImageItem {
  kind: "image";
  x: number;
  y: number;
  width: number;
  height: number;
  /** Package part path; renderer resolves bytes via DocxDocument.media(). */
  part: string;
  /** a:srcRect crop (fractions) and a:xfrm rotation (degrees). */
  crop?: { l: number; t: number; r: number; b: number };
  rotation?: number;
  /** behindDoc: paint under the text layer. */
  behind?: boolean;
  /** Source w:drawing element (for interactive resize/move). */
  src?: XmlElement;
}

/** Interactive resize zone over a table boundary (column or row). */
export interface GripItem {
  kind: "grip";
  /** "col": vertical zone at x spanning y1..y2. "row": horizontal zone at y1
   * spanning x..x2. */
  axis: "col" | "row";
  x: number;
  y1: number;
  y2: number;
  /** Right edge for row grips. */
  x2?: number;
  /** Source w:tbl element. */
  tbl: XmlElement;
  /** col: boundary 1..n (n = right edge). row: row index above the boundary. */
  boundary: number;
  /** Laid-out height of the row above (row grips), px. */
  rowHeightPx?: number;
  /** Rendered column widths px (col grips) — resize works in this space. */
  renderedWidths?: number[];
}

export type PageItem = TextItem | RectItem | LineEdgeItem | ImageItem | GripItem;

export interface LaidOutPage {
  width: number;
  height: number;
  /** 1-based physical index. */
  index: number;
  /** Display page number after pgNumType.start is applied. */
  number: number;
  items: PageItem[];
  /** Body box (for header/footer editing chrome). */
  bodyTop: number;
  bodyBottom: number;
  /** Items from this index on belong to the header/footer parts. */
  hfStart: number;
}

export interface LayoutResult {
  pages: LaidOutPage[];
  totalPages: number;
}
