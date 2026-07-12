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
  /** Optional PAINT-only family tried before `family` in the CSS stack, without
   * affecting metrics or width lookups (keyed by `family`). Used for CJK: the
   * line-pitch profile stays keyed by the macOS substitute face while the real
   * Windows glyphs (MS Mincho, Microsoft JhengHei, …) paint when available. */
  paintFamily?: string;
}

export interface PathItem {
  kind: "path";
  x: number;
  y: number;
  width: number;
  height: number;
  /** SVG path data in a `viewW x viewH` coordinate space. */
  d: string;
  viewW: number;
  viewH: number;
  fill?: string;
  stroke?: { color: string; width: number };
  /** Paint under / over the body text (anchored-shape layering). */
  behind?: boolean;
  front?: boolean;
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
  /** Footnote/endnote id referenced by this run (registration happens when
   * the item lands on a real page - split table rows carry it across). */
  noteId?: number;
  /** Vertical stretch for tall delimiter glyphs (Word's glyph variants). */
  mathScaleY?: number;
  /** Horizontal stretch for wide brace glyphs (Word's over/under brace
   * variants), applied about the glyph box's horizontal center. */
  mathScaleX?: number;
  /** Stretch anchor above the baseline, px. */
  mathScaleAnchor?: number;
  /** Line box for selection/highlight backgrounds. */
  lineTop: number;
  lineHeight: number;
  /** Exact glyph box for baseline-shifted runs (superscript/subscript):
   * the renderer anchors these instead of bottoming on the line box. */
  glyphTop?: number;
  glyphBoxH?: number;
  /** Small-caps reduced segment: the base run font that must supply the
   * strut so the painted baseline matches neighboring full-size spans (the
   * renderer sizes the outer span with this font and shrinks the text via
   * a baseline-aligned inner span). */
  strutFont?: FontSpec;
  /** PAGEREF bookmark name: the final pass rewrites this item's text with
   * the bookmark's page number (Word recomputes PAGEREF on open; the docx
   * cached result is stale in real TOCs). */
  pageRef?: string;
  /** Source m:oMath element when this text is a piece of an equation. */
  mathSrc?: XmlElement;
  href?: string;
  /** Present for editable text (absent on numbering labels etc.). */
  src?: TextSource;
  /** Rotate about a point (px, relative to this item's top-left). */
  rotate?: { deg: number; ox: number; oy: number };
  /** Paint under the body text (behindDoc textbox content). */
  behind?: boolean;
  /** Paint over the body text (non-behindDoc anchored shape content: Word
   * layers in-front shapes above the text layer). */
  front?: boolean;
  /** Right-to-left run: renderer sets direction:rtl so the browser shapes and
   * orders the (Arabic/Hebrew) glyphs within the span box. */
  rtl?: boolean;
  /** Ordinal of the source paragraph within its cell/frame. Set only for
   * widow-controlled paragraphs; row splitting uses it to scope Word's
   * widow/orphan rules to the paragraph straddling the cut. */
  paraSeq?: number;
}

export interface RectItem {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
  /** Semantic paint provenance used by render-layer comparison. */
  role?: "table-fill";
  /** Rotate about a point (px, relative to this item's top-left). */
  rotate?: { deg: number; ox: number; oy: number };
  behind?: boolean;
  /** Paint over the body text (non-behindDoc anchored shape fill). */
  front?: boolean;
}

export interface LineEdgeItem {
  kind: "edge";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  border: Border;
  /** Semantic paint provenance used by render-layer comparison. */
  role?: "table-rule";
  /** Rotate about a point (px, relative to this item's top-left). */
  rotate?: { deg: number; ox: number; oy: number };
  /** Paint over the body text (non-behindDoc anchored shape border). */
  front?: boolean;
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
  /** Picture-watermark "washout" (VML v:imagedata gain/blacklevel, 0..1
   * fractions). Per-channel linear recolor, measured against Word's PDF of
   * probe2-picture-watermark (gain 0.3, blacklevel 0.35: source 32 -> 215,
   * 74 -> 227, 135 -> 246, 210 -> clamped 255):
   *   out = in * gain + 255 * (blacklevel * (1 + gain) + (1 - gain) / 2)  */
  washout?: { gain: number; blacklevel: number };
  /** a:ln picture outline, drawn just outside the image (Word hairline). */
  border?: { color: string; width: number };
  /** behindDoc: paint under the text layer. */
  behind?: boolean;
  /** wrapNone + !behindDoc ("in front of text"): paint ABOVE the text layer
   * (Word's z-order; without this, later-emitted text spans cover the image
   * and it neither shows in front nor receives clicks/drags). */
  front?: boolean;
  /** Source w:drawing element (for interactive resize/move). */
  src?: XmlElement;
}

/** Interactive resize zone over a table boundary (column or row). */
export interface DrawingHitItem {
  kind: "drawingHit";
  x: number;
  y: number;
  width: number;
  height: number;
  /** Source w:drawing element (select/move the whole drawing). */
  src: XmlElement;
  /** Anchored drawings drag by offset; inline ones re-anchor into text. */
  anchored: boolean;
}

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

/** WordArt/watermark text scaled to fill a box, rotated about its center. */
export interface WordArtItem {
  kind: "wordart";
  /** Box top-left, px. */
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  fontFamily: string;
  bold?: boolean;
  italic?: boolean;
  fill: string;
  opacity: number;
  /** Clockwise degrees. */
  rotation: number;
  behind?: boolean;
  /** v:textpath font-size (px), used only when noFit is set. */
  fontSize?: number;
  /** Malformed shapetype guide path: render at nominal fontSize, unstretched
   * (Word can't fitshape, so it draws a near-invisible mark). */
  noFit?: boolean;
  /** Source VML shape element (v:shape / v:rect), for interactive select/edit. */
  src?: XmlElement;
}

export type PageItem =
  | TextItem
  | PathItem
  | RectItem
  | LineEdgeItem
  | ImageItem
  | DrawingHitItem
  | GripItem
  | WordArtItem;

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
