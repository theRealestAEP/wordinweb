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
  /** w:w horizontal character scaling as a fraction (1.5 = 150%). */
  textScale?: number;
  /** w:position baseline shift in px, positive = raised. */
  raise?: number;
  /** w:outline — hollow stroked glyphs. */
  outline?: boolean;
  /** w:emboss / w:imprint — Word triple-draws offset copies. */
  emboss?: boolean;
  imprint?: boolean;
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
  /** w:beforeAutospacing/afterAutospacing: HTML-style automatic paragraph
   * spacing — Word ignores the literal before/after and inserts one blank
   * line's worth of space. */
  beforeAutospacing?: boolean;
  afterAutospacing?: boolean;
  /** Word's default when absent is auto/1.0 via docDefaults. */
  lineSpacing?: LineSpacing;
  contextualSpacing?: boolean;
  keepNext?: boolean;
  /** w:framePr w:dropCap: the paragraph is a drop-cap letter frame. */
  dropCap?: { mode: "drop" | "margin"; lines: number; hSpace: number };
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  borders?: ParagraphBorders;
  /** Paragraph shading fill as CSS color. */
  shading?: string;
  numbering?: NumberingRef | null; // null = explicitly removed (numId 0)
  /** A numPr that carries ilvl but no numId (Heading3 basedOn Heading2)
   * overrides only the list LEVEL, keeping the inherited numId; the style
   * chain applies it to `numbering` on merge. */
  numberingLevelOverride?: number;
  tabs?: TabStop[];
  outlineLevel?: number;
  /** Run props declared on pPr/rPr — apply to the paragraph mark & numbering label. */
  markRunProps?: RunProps;
  /**
   * styleId of the enclosing table (set on cell paragraphs at parse time).
   * A table style's own pPr layers between docDefaults and the paragraph
   * style, e.g. TableGrid's `spacing after=0 line=240` overrides docDefaults'
   * `after=200 line=276` so list cells lay out compactly.
   */
  tableStyleId?: string;
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
export interface PTabContent {
  kind: "ptab";
  /** Absolute-position tab: jump to left/center/right of the base. */
  alignment: "left" | "center" | "right";
  relativeTo: "margin" | "indent";
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
/** OMML equation node (subset: runs, scripts, fractions, radicals). */
export type MathNode =
  | { t: "run"; text: string }
  | { t: "sup" | "sub"; base: MathNode[]; script: MathNode[] }
  | { t: "frac"; num: MathNode[]; den: MathNode[]; bar?: boolean }
  | { t: "rad"; e: MathNode[] }
  /** n-ary operator (sum/integral); chr defaults to the integral sign. */
  | { t: "nary"; chr: string; sub: MathNode[]; sup: MathNode[]; e: MathNode[] }
  /** Delimiters grown to the content height; beg/end default to parens. */
  | { t: "dlm"; beg: string; end: string; e: MathNode[][] }
  /** Matrix: rows x cells. */
  | { t: "mat"; rows: MathNode[][][] };

export interface MathContent {
  kind: "math";
  nodes: MathNode[];
  /** Source m:oMath element (math editing). */
  src?: XmlElement;
  /** Display equation (wrapped in m:oMathPara): centered on its own line with
   * display-style layout - larger n-ary operators with limits stacked
   * above/below, and full-size fraction numerators/denominators. */
  display?: boolean;
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
  /** Background fill (CSS color) painted behind the text. */
  fill?: string;
  /** Outline. */
  stroke?: { color: string; weight: number };
  /** Alignment-based positioning (mso-position-horizontal/vertical). */
  hAlign?: "left" | "center" | "right";
  vAlign?: "top" | "center" | "bottom";
  /** Percent-of-base geometry, 0..1 (mso-*-percent / wp14 pct offsets).
   * Bases: pctWidthRel/pctHeightRel say page vs margin. */
  pctX?: number;
  pctY?: number;
  pctWidth?: number;
  pctHeight?: number;
  pctWidthRel?: "page" | "margin";
  pctHeightRel?: "page" | "margin";
  /** Vertical anchoring of the text INSIDE the box (v-text-anchor). */
  textAnchor?: "top" | "middle" | "bottom";
  /** How body text flows around a DrawingML textbox (wp:wrap*). */
  wrap?: WrapMode;
  /** behindDoc: paint under the body text, never displace it. */
  behind?: boolean;
  /** Wrap distances px (wp:anchor distT/B/L/R). */
  dist?: { t: number; b: number; l: number; r: number };
  /** a:xfrm rotation, degrees clockwise (rotates the whole box). */
  rotation?: number;
  /** Text insets px (bodyPr lIns/tIns/rIns/bIns); default 9.6/4.8. */
  insets?: { l: number; t: number; r: number; b: number };
}

/** WordArt (VML v:textpath, e.g. a "CONFIDENTIAL" watermark): text scaled to
 * fill a box, optionally rotated, painted semi-transparent behind the body. */
export interface ShapeWordArt {
  type: "wordart";
  text: string;
  fontFamily: string;
  bold?: boolean;
  italic?: boolean;
  /** CSS fill color. */
  fill: string;
  /** 0..1 alpha. */
  opacity: number;
  x: number;
  y: number;
  width: number;
  height: number;
  hRel: AnchorRel;
  vRel: AnchorRel;
  hAlign?: "left" | "center" | "right";
  vAlign?: "top" | "center" | "bottom";
  /** Clockwise degrees. */
  rotation: number;
  behind?: boolean;
}

export interface ShapeArt {
  type: "art";
  srcDrawing?: XmlElement;
  x: number;
  y: number;
  /** Percent-of-page offsets (wp14:pctPos*Offset), 0..1. */
  pctX?: number;
  pctY?: number;
  width: number;
  height: number;
  hRel: AnchorRel;
  vRel: AnchorRel;
  hAlign?: "left" | "center" | "right";
  behind?: boolean;
  lines: DrawingLine[];
  images: DrawingImage[];
  paths: DrawingPath[];
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
  /** Wrap distances px (wp:anchor distT/B/L/R); text clears the image by these. */
  dist?: { t: number; b: number; l: number; r: number };
  crop?: { l: number; t: number; r: number; b: number };
  rotation?: number;
  /** Source w:drawing element (editing). */
  srcDrawing?: XmlElement;
}

export type Shape = ShapeLine | ShapeTextbox | ShapeImage | ShapeArt | ShapeWordArt;

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
  /** Freeform vector shapes (a:custGeom), as SVG path data. */
  paths?: DrawingPath[];
  /** An INLINE wps text box (wp:inline wps:txbx): a fixed-extent box that
   * flows in the text (occupying its width x height like an inline image) and
   * carries a fill/border + its own block content. Distinct from the floating
   * ShapeTextbox, which is absolutely placed via a wp:anchor. */
  textbox?: {
    blocks: Block[];
    fill?: string;
    stroke?: { color: string; weight: number };
    insets?: { l: number; t: number; r: number; b: number };
    textAnchor?: "top" | "middle" | "bottom";
  };
  /** Source w:drawing element (select/move as a group). */
  srcDrawing?: XmlElement;
}

export interface DrawingPath {
  /** Position/size inside the drawing, px. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** SVG path in the `viewW x viewH` source coordinate space. */
  d: string;
  viewW: number;
  viewH: number;
  fill?: string;
  stroke?: { color: string; width: number };
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
  | MathContent
  | TextContent
  | BreakContent
  | TabContent
  | PTabContent
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
  /** w:bookmarkStart names in this paragraph (PAGEREF targets). */
  bookmarks?: string[];
  /** Final revision view: the paragraph mark AND all content are tracked
   * deletions — the paragraph does not exist (no line, no numbering). */
  revisionHidden?: boolean;
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
  /** w:tblLook flags controlling which conditional formats apply. */
  tblLook?: {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHBand: boolean;
    noVBand: boolean;
  };
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
  /** w:lnNumType: margin line numbering. distance px from the text edge. */
  lineNumbering?: {
    countBy: number;
    start: number;
    distance: number;
    restart: "continuous" | "newPage" | "newSection";
  };
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

/** One w:tblStylePr conditional-formatting block (firstRow, band1Horz, …). */
export interface TableCondFormat {
  /** Cell shading fill as CSS color. */
  shd?: string;
  /** Conditional cell borders. */
  borders?: {
    top?: Border;
    bottom?: Border;
    left?: Border;
    right?: Border;
    insideH?: Border;
    insideV?: Border;
  };
  /** Conditional run bold (firstRow/firstCol headers). */
  bold?: boolean;
}

export type TableCondType =
  | "wholeTable"
  | "band1Vert"
  | "band2Vert"
  | "band1Horz"
  | "band2Horz"
  | "firstRow"
  | "lastRow"
  | "firstCol"
  | "lastCol"
  | "nwCell"
  | "neCell"
  | "swCell"
  | "seCell";

export interface Style {
  id: string;
  type: "paragraph" | "character" | "table" | "numbering";
  name?: string;
  basedOn?: string;
  isDefault?: boolean;
  pPr?: ParaProps;
  rPr?: RunProps;
  tblPr?: TableProps;
  /** w:tblStylePr conditional formats, by type (table styles only). */
  condFormats?: Map<TableCondType, TableCondFormat>;
  /** w:tblStyleRowBandSize / ColBandSize (default 1). */
  rowBandSize?: number;
  colBandSize?: number;
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
