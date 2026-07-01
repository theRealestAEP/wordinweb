/**
 * Typed document model produced by the parser and consumed by the layout
 * engine. All lengths are CSS px (see units.ts) unless a field name says
 * otherwise. Property bags are Partial-style: `undefined` means "not set at
 * this level" so the style-inheritance merge can distinguish absence from an
 * explicit value.
 */

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
}
export interface FieldContent {
  kind: "field";
  /** Raw field instruction, e.g. ` PAGE \\* MERGEFORMAT `. */
  instruction: string;
  /** Last cached result text from the file, used for unsupported fields. */
  cachedResult: string;
}
export type RunContent = TextContent | BreakContent | TabContent | ImageContent | FieldContent;

export interface Run {
  type: "run";
  props: RunProps;
  content: RunContent[];
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
}

export interface Section {
  props: SectionProps;
  blocks: Block[];
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
