/**
 * Typed document model produced by the parser and consumed by the layout
 * engine. All lengths are CSS px (see units.ts) unless a field name says
 * otherwise. Property bags are Partial-style: `undefined` means "not set at
 * this level" so the style-inheritance merge can distinguish absence from an
 * explicit value.
 *
 * Model nodes keep `src` references to the XML elements they were parsed
 * from: the XML tree is the source of truth for editing — commands mutate it
 * and the model is re-derived, which preserves round-trip fidelity for
 * everything untouched.
 */

import type { XmlElement } from "./xml.js";

// ---------- shared property primitives ----------

export type BorderStyle =
  | "none"
  | "single"
  | "double"
  | "dotted"
  | "dashed"
  | "thick"
  | "wave"
  | "dotDash"
  | "dotDotDash"
  | "triple";

export interface Border {
  style: BorderStyle;
  /** Stroke width in px. */
  width: number;
  /** CSS color. */
  color: string;
  /** Gap between border and content in px (w:space, points in OOXML). */
  space: number;
}

export interface ParagraphBorders {
  top?: Border;
  bottom?: Border;
  left?: Border;
  right?: Border;
  /** Drawn between consecutive paragraphs that both specify it. */
  between?: Border;
}

export type LineSpacingRule = "auto" | "atLeast" | "exact";

export interface LineSpacing {
  rule: LineSpacingRule;
  /**
   * For "auto": multiple of single line height (240ths in OOXML, stored here
   * as the multiplier, e.g. 1.15). For atLeast/exact: px.
   */
  value: number;
}

export type Alignment = "left" | "center" | "right" | "justify";

export interface TabStop {
  /** px from the left text edge */
  pos: number;
  align: "left" | "center" | "right" | "decimal" | "bar";
  leader: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
}

// ---------- run properties ----------

export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: string; // w:u val ("single", "none", ...)
  strike?: boolean;
  doubleStrike?: boolean;
  /** Resolved primary font family name (ascii/hAnsi, theme-resolved). */
  font?: string;
  /** Font size in px. */
  size?: number;
  /** CSS color; "auto" resolved to inherit/black by renderer. */
  color?: string;
  highlight?: string;
  /** Character shading fill (w:shd). */
  shading?: string;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  caps?: boolean;
  smallCaps?: boolean;
  /** Letter spacing px (w:spacing, twips in OOXML). */
  letterSpacing?: number;
  vanish?: boolean;
  /** Character style id (w:rStyle). */
  styleId?: string;
  lang?: string;
}

// ---------- paragraph properties ----------

export interface NumberingRef {
  numId: number;
  ilvl: number;
}

export interface ParaProps {
  styleId?: string;
  alignment?: Alignment;
  /** px */
  indentLeft?: number;
  indentRight?: number;
  /** Positive px; mutually exclusive with hanging in effect. */
  indentFirstLine?: number;
  /** Positive px hanging indent. */
  indentHanging?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  /** Word's default when absent is auto/1.0 via docDefaults. */
  lineSpacing?: LineSpacing;
  contextualSpacing?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  borders?: ParagraphBorders;
  /** Paragraph shading fill as CSS color. */
  shading?: string;
  numbering?: NumberingRef | null; // null = explicitly removed (numId 0)
  tabs?: TabStop[];
  outlineLevel?: number;
  /** Run props declared on pPr/rPr — apply to the paragraph mark & numbering label. */
  markRunProps?: RunProps;
}

// ---------- run content ----------

export interface TextContent {
  kind: "text";
  text: string;
  /** Source w:t element inside the run, when this text came verbatim from one. */
  srcT?: XmlElement;
}
export interface BreakContent {
  kind: "break";
  breakType: "line" | "page" | "column";
}
export interface TabContent {
  kind: "tab";
}
export interface ImageContent {
  kind: "image";
  /** Part path of the image inside the package. */
  part: string;
  width: number;
  height: number;
  /** True when the drawing is floating (wp:anchor); v1 lays it out inline. */
  anchored?: boolean;
  /** a:srcRect crop, fractions of the source bitmap. */
  crop?: { l: number; t: number; r: number; b: number };
  /** a:xfrm rotation, degrees clockwise. */
  rotation?: number;
  /** Source w:drawing (or pict) element, for resize/move editing. */
  srcDrawing?: XmlElement;
}
export interface FieldContent {
  kind: "field";
  /** Raw field instruction, e.g. ` PAGE \\* MERGEFORMAT `. */
  instruction: string;
  /** Last cached result text from the file, used for unsupported fields. */
  cachedResult: string;
}

/** What an anchored shape's coordinates are measured from. */
export type AnchorRel = "page" | "margin" | "text" | "column";

export interface ShapeLine {
  type: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  /** Stroke weight px. */
  weight: number;
  hRel: AnchorRel;
  vRel: AnchorRel;
}

export interface ShapeTextbox {
  type: "textbox";
  x: number;
  y: number;
  width: number;
  height: number;
  hRel: AnchorRel;
  vRel: AnchorRel;
  blocks: Block[];
}

/** How text interacts with a floating image. */
export type WrapMode = "square" | "topAndBottom" | "none";

export interface ShapeImage {
  type: "image";
  part: string;
  /** Offset from the anchor origin, px. */
  x: number;
  y: number;
  width: number;
  height: number;
  hRel: AnchorRel;
  vRel: AnchorRel;
  /** Horizontal alignment when the file uses wp:align instead of an offset. */
  hAlign?: "left" | "center" | "right";
  wrap: WrapMode;
  /** behindDoc anchors render under the text and never displace it. */
  behind?: boolean;
  crop?: { l: number; t: number; r: number; b: number };
  rotation?: number;
  /** Source w:drawing element (editing). */
  srcDrawing?: XmlElement;
}

export type Shape = ShapeLine | ShapeTextbox | ShapeImage;

/**
 * Floating/anchored object: does not occupy inline space; positioned against
 * page/margin/paragraph during layout. (How classic pleading paper draws its
 * margin line numbers and vertical rules.)
 */
export interface AnchorContent {
  kind: "anchor";
  shape: Shape;
}

/** A stroked segment inside a composite drawing, px relative to its box. */
export interface DrawingLine {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  weight: number;
}

export interface DrawingImage {
  part: string;
  x: number;
  y: number;
  width: number;
  height: number;
  crop?: { l: number; t: number; r: number; b: number };
  rotation?: number;
}

/**
 * Composite inline drawing (DrawingML groups): vector lines + placed images
 * inside a width×height box that flows like an image.
 */
export interface DrawingContent {
  kind: "drawing";
  width: number;
  height: number;
  lines: DrawingLine[];
  images: DrawingImage[];
}

/**
 * Footnote/endnote reference mark. In body text (w:footnoteReference /
 * w:endnoteReference) `id` points into the notes part. Inside a note body,
 * w:footnoteRef / w:endnoteRef render the note's own mark: `self` is true and
 * `id` is meaningless.
 */
export interface NoteRefContent {
  kind: "noteRef";
  noteType: "footnote" | "endnote";
  id: number;
  self?: boolean;
}

export type RunContent =
  | TextContent
  | BreakContent
  | TabContent
  | ImageContent
  | FieldContent
  | AnchorContent
  | DrawingContent
  | NoteRefContent;

export interface Run {
  type: "run";
  props: RunProps;
  content: RunContent[];
  /** Source w:r element. */
  src?: XmlElement;
  /** Element whose children array contains src (w:p, w:hyperlink, …). */
  srcParent?: XmlElement;
}

export interface Hyperlink {
  type: "hyperlink";
  href?: string;
  anchor?: string;
  runs: Run[];
}

export type ParaChild = Run | Hyperlink;

// ---------- blocks ----------

export interface Paragraph {
  type: "paragraph";
  props: ParaProps;
  children: ParaChild[];
  /** Section break attached to this paragraph's pPr (ends a section). */
  sectionBreak?: SectionProps;
  /** Source w:p element. */
  src?: XmlElement;
}

export interface TableCellProps {
  /** Preferred width px (from tcW when dxa). */
  width?: number;
  gridSpan: number;
  vMerge?: "restart" | "continue";
  borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border };
  shading?: string;
  /** Cell margins px. */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  verticalAlign?: "top" | "center" | "bottom";
}

export interface TableCell {
  props: TableCellProps;
  blocks: Block[];
}

export interface TableRowProps {
  /** px; undefined = auto */
  height?: number;
  heightRule?: "auto" | "atLeast" | "exact";
  cantSplit?: boolean;
  /** Repeat as header row on each page. */
  tblHeader?: boolean;
}

export interface TableRow {
  props: TableRowProps;
  cells: TableCell[];
}

export interface TableProps {
  styleId?: string;
  /** px table indent from left margin. */
  indent?: number;
  alignment?: Alignment;
  borders?: {
    top?: Border;
    bottom?: Border;
    left?: Border;
    right?: Border;
    insideH?: Border;
    insideV?: Border;
  };
  /** Default cell margins px. */
  cellMargins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Preferred total width px (tblW dxa) or undefined for auto. */
  width?: number;
  /** Preferred total width as a fraction of available width (tblW pct). */
  widthPct?: number;
  layout?: "fixed" | "autofit";
}

export interface Table {
  type: "table";
  props: TableProps;
  /** Column grid widths px (tblGrid). */
  grid: number[];
  rows: TableRow[];
  /** Source w:tbl element. */
  src?: XmlElement;
}

export type Block = Paragraph | Table;

// ---------- sections ----------

export interface HeaderFooterRefs {
  default?: string; // relationship id
  first?: string;
  even?: string;
}

export interface ColumnSpec {
  count: number;
  /** Space between columns px. */
  space: number;
  /** Explicit widths px when equalWidth=false. */
  widths?: number[];
}

export interface SectionProps {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  /** Distance of header top from page top, px (w:headerReference distance). */
  headerDistance: number;
  footerDistance: number;
  gutter: number;
  headerRefs: HeaderFooterRefs;
  footerRefs: HeaderFooterRefs;
  /** Different first-page header/footer enabled. */
  titlePage: boolean;
  /** Page numbering restart value, if set. */
  pageNumberStart?: number;
  pageNumberFormat?: string;
  columns: ColumnSpec;
  type?: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
  /** Vertical alignment of page content. */
  vAlign?: "top" | "center" | "both" | "bottom";
  /** w:pgBorders. Offsets (border.space, px) measure from text or page edge. */
  pageBorders?: { top?: Border; bottom?: Border; left?: Border; right?: Border; offsetFrom: "text" | "page" };
  /** Footnote/endnote mark numbering (w:footnotePr / w:endnotePr). */
  footnoteNumFmt?: string;
  footnoteNumStart?: number;
  endnoteNumFmt?: string;
  endnoteNumStart?: number;
}

export interface Section {
  props: SectionProps;
  blocks: Block[];
}

// ---------- comments ----------

/** A review comment from word/comments.xml. */
export interface DocComment {
  id: string;
  author: string;
  initials?: string;
  /** ISO timestamp from w:date, verbatim. */
  date?: string;
  /** Plain text of the comment body (paragraphs joined with newlines). */
  text: string;
  /** w14:paraId of the comment's last body paragraph (threading key). */
  paraId?: string;
  /** Parent comment id when this comment is a reply (commentsExtended). */
  parentId?: string;
}

// ---------- headers / footers ----------

export interface HeaderFooter {
  blocks: Block[];
}

// ---------- styles ----------

export interface Style {
  id: string;
  type: "paragraph" | "character" | "table" | "numbering";
  name?: string;
  basedOn?: string;
  isDefault?: boolean;
  pPr?: ParaProps;
  rPr?: RunProps;
  tblPr?: TableProps;
}

export interface Styles {
  byId: Map<string, Style>;
  defaultParagraphStyle?: string;
  defaultRPr: RunProps;
  defaultPPr: ParaProps;
}

// ---------- numbering ----------

export interface NumberingLevel {
  ilvl: number;
  start: number;
  format: string; // decimal, bullet, lowerRoman, upperRoman, lowerLetter, upperLetter, none...
  /** Template like "%1." */
  text: string;
  alignment: Alignment;
  /** Paragraph props contributed by the level (indents). */
  pPr?: ParaProps;
  rPr?: RunProps;
  suffix: "tab" | "space" | "nothing";
  restartAfter?: number;
}

export interface AbstractNum {
  id: number;
  levels: Map<number, NumberingLevel>;
  numStyleLink?: string;
}

export interface NumInstance {
  numId: number;
  abstractNumId: number;
  overrides: Map<number, { startOverride?: number; level?: NumberingLevel }>;
}

export interface Numbering {
  abstract: Map<number, AbstractNum>;
  instances: Map<number, NumInstance>;
}

// ---------- theme ----------

export interface Theme {
  majorFont: string;
  minorFont: string;
  colors: Map<string, string>;
}
