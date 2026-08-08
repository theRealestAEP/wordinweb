import type { DocxDocument } from "@wordinweb/core";
import type { IntentBody, PresencePosition } from "@wordinweb/collab/client";

export type AgentReference = string;

export interface AgentCursor {
  value: string;
}

export interface AgentOverview {
  revision: string;
  sections: number;
  mirrorMargins: boolean;
  sectionLayouts: AgentSectionLayout[];
  stories: Array<{ id: string; kind: "body" | "header" | "footer" | "footnote" | "endnote" | "textbox"; blocks: number }>;
  blocks: { paragraphs: number; tables: number };
  characters: number;
  comments: number;
  components: Record<string, number>;
  objectCounts: Record<string, number>;
  outline: OutlineEntry[];
}

export interface AgentSectionLayout {
  index: number;
  page: { width: number; height: number };
  margins: { top: number; right: number; bottom: number; left: number };
  headerDistance: number;
  footerDistance: number;
  gutter: number;
  columns: { count: number; space: number; widths?: number[]; spaces?: number[]; separator?: boolean };
  titlePage: boolean;
  pageNumberStart?: number;
  pageNumberFormat?: string;
  breakType?: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
  verticalAlignment?: "top" | "center" | "both" | "bottom";
  pageBorders?: AgentPageBorders;
  lineNumbering?: { countBy: number; start: number; distance: number; restart: "continuous" | "newPage" | "newSection" };
}

export interface AgentBorder {
  style: "none" | "single" | "double" | "dotted" | "dashed" | "thick" | "wave" | "dotDash" | "dotDotDash" | "thinThickSmallGap" | "triple";
  width: number;
  color: string;
  space: number;
  rawWidth?: number;
}

export interface AgentPageBorders {
  top?: AgentBorder;
  bottom?: AgentBorder;
  left?: AgentBorder;
  right?: AgentBorder;
  offsetFrom: "text" | "page";
}

export interface OutlineEntry {
  ref: AgentReference;
  level: number;
  text: string;
  styleId?: string;
}

export interface AgentRun {
  ref: AgentReference;
  editable: boolean;
  paragraphStart: number;
  paragraphEnd: number;
  wireLength: number;
  text: string;
  formatting: Record<string, unknown>;
  hyperlink?: { href?: string; anchor?: string };
  components: AgentComponentSummary[];
}

export interface AgentParagraph {
  type: "paragraph";
  ref: AgentReference;
  editable: boolean;
  story: string;
  styleId?: string;
  outlineLevel?: number;
  alignment?: string;
  list?: { numId: number; level: number } | null;
  text: string;
  textRange: { start: number; end: number; total: number };
  runs: AgentRun[];
  bookmarks: string[];
}

export interface AgentTable {
  type: "table";
  ref: AgentReference;
  editable: boolean;
  story: string;
  rows: number;
  columns: number;
  styleId?: string;
  floating: boolean;
  cells: Array<{ row: number; column: number; blocks: AgentReference[] }>;
}

export type AgentBlock = AgentParagraph | AgentTable;

export interface AgentReadResult {
  revision: string;
  story: string;
  blocks: AgentBlock[];
  next?: AgentCursor;
  truncated: boolean;
}

export interface AgentContextRun {
  ref: AgentReference;
  start: number;
  end: number;
  wireLength?: number;
}

export interface AgentContextParagraph {
  type: "paragraph";
  ref: AgentReference;
  text: string;
  range?: { start: number; end: number; total: number };
  runs: AgentContextRun[];
  styleId?: string;
  outlineLevel?: number;
  list?: { numId: number; level: number };
  bookmarks?: string[];
  objects?: AgentComponentSummary[];
}

export interface AgentContextTable {
  type: "table";
  ref: AgentReference;
  rows: number;
  columns: number;
}

export interface AgentContextStory {
  story: string;
  kind: "body" | "header" | "footer" | "footnote" | "endnote" | "textbox";
  blocks: Array<AgentContextParagraph | AgentContextTable>;
  next?: AgentCursor;
}

export interface AgentContextResult {
  revision: string;
  contents: AgentContextStory[];
  truncated: boolean;
  remainingStories?: string[];
}

export interface AgentSearchResult {
  revision: string;
  matches: Array<{
    blockRef: AgentReference;
    runRef: AgentReference;
    editable: boolean;
    start: number;
    end: number;
    excerpt: string;
  }>;
  truncated: boolean;
}

export interface AgentComponentSummary {
  ref: AgentReference;
  editRef?: AgentReference;
  type: string;
  label?: string;
}

export interface AgentObjectResult {
  revision: string;
  ref: AgentReference;
  editRef?: AgentReference;
  type: string;
  detail: Record<string, unknown>;
}

export interface AgentAsset {
  ref: AgentReference;
  bytes: Uint8Array;
  mediaType: string;
}

export interface SpatialBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface AgentSpatialObject {
  ref: AgentReference;
  editRef?: AgentReference;
  type: string;
  page: number;
  bounds: SpatialBounds;
  rotation?: number;
  layer: "behind-text" | "body" | "in-front-of-text";
  zOrder?: number;
}

export interface AgentOverlap {
  objects: [AgentReference, AgentReference];
  overlapBounds: SpatialBounds;
  overlapArea: number;
  topObject?: AgentReference;
}

export interface AgentSpatialResult {
  revision: string;
  layout: { quality: "exact" | "approximate"; profile: string };
  totalPages: number;
  pages: Array<{ index: number; width: number; height: number }>;
  objects: AgentSpatialObject[];
  overlaps: AgentOverlap[];
}

export type AgentInspectRequest =
  | { kind: "overview" }
  | { kind: "context"; stories?: string[]; maxBlocks?: number; maxCharacters?: number; include?: Array<"bookmarks" | "objects">; includeEmpty?: boolean }
  | { kind: "read"; story?: string; cursor?: AgentCursor; maxBlocks?: number; maxCharacters?: number }
  | { kind: "search"; query: string; maxResults?: number }
  | { kind: "object"; ref: AgentReference }
  | { kind: "spatial"; pages?: { start: number; count: number }; includeOverlaps?: boolean };

export type AgentInspectResult =
  | AgentOverview
  | AgentContextResult
  | AgentReadResult
  | AgentSearchResult
  | AgentObjectResult
  | AgentSpatialResult;

export type AgentProjectionMode = "text" | "md" | "outline";

/** One run of projected characters that came from one place in the document.
 * `start`/`end` are columns in the projection line; `wireStart`/`wireEnd` are
 * offsets in the run's wire space. An editable segment is backed by a `w:t`,
 * so the two spaces line up one-to-one across it and a patch can address any
 * column inside it. Everything else is an atom the patch path treats as
 * opaque. */
export interface AgentAnchorSegment {
  start: number;
  end: number;
  runRef: AgentReference;
  wireStart: number;
  wireEnd: number;
  editable: boolean;
}

export interface AgentAnchorLine {
  line: number;
  role: "paragraph" | "table" | "structure";
  blockRef?: AgentReference;
  /** Leading characters of the line that are markdown structure rather than
   * document text (`## `, `- `, `1. `). Always 0 in `text` mode. */
  marker: number;
  editable: boolean;
  segments: AgentAnchorSegment[];
}

export interface AgentProjectRequest {
  story?: string;
  mode?: AgentProjectionMode;
  cursor?: AgentCursor;
  maxBlocks?: number;
  maxCharacters?: number;
}

export interface AgentProjectResult {
  revision: string;
  story: string;
  mode: AgentProjectionMode;
  text: string;
  lines: number;
  window: { cursor: string | null; startBlock: number; endBlock: number; maxBlocks: number; maxCharacters: number };
  anchors: AgentAnchorLine[];
  next?: AgentCursor;
  truncated: boolean;
}

export interface AgentPatchHunk {
  startLine: number;
  endLine: number;
  newText: string;
}

export interface AgentPatchRequest {
  revision: string;
  story?: string;
  mode?: AgentProjectionMode;
  cursor?: AgentCursor;
  maxBlocks?: number;
  maxCharacters?: number;
  edits?: AgentPatchHunk[];
  diff?: string;
  suggest?: boolean;
}

export interface AgentPatchResult {
  revision: string;
  status: "applied" | "submitted";
  operations: string[];
  connection?: "local" | "live" | "offline" | "reconnecting";
  projection: AgentProjectResult;
}

export type AgentOperation = Record<string, unknown> & { kind: string };

export interface AgentEditRequest {
  revision: string;
  operations: AgentOperation[];
}

export interface AgentEditResult {
  revision: string;
  status: "applied" | "submitted";
  operations: string[];
  connection?: "local" | "live" | "offline" | "reconnecting";
}

export interface AgentComposeRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  highlight?: string;
  fontSizePt?: number;
  fontFamily?: string;
}

export interface AgentComposeCell {
  text?: string;
  runs?: AgentComposeRun[];
  fill?: string;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "center" | "bottom";
}

export interface AgentComposeTextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontSizePt?: number;
  fontFamily?: string;
  align?: "left" | "center" | "right";
}

export type AgentComposeBlock =
  | { type: "paragraph"; text?: string; runs?: AgentComposeRun[]; styleId?: "Normal" | "Title" | "Subtitle" | "Heading1" | "Heading2" | "Heading3"; align?: "left" | "center" | "right" | "both"; spacing?: { beforePt?: number; afterPt?: number; lineMultiple?: number }; list?: "bullet" | "number" }
  | { type: "heading"; level: 1 | 2 | 3; text: string }
  | { type: "table"; rows: Array<Array<string | AgentComposeCell>>; headerRows?: number; headerFill?: string; headerTextColor?: string; columnWidths?: number[] }
  | { type: "equation"; mathText: string; align?: "left" | "center" | "right" }
  | { type: "chart"; chart: { type: "column" | "bar" | "line" | "pie"; title?: string; categories: string[]; series: Array<{ name: string; values: number[] }> }; widthPx?: number; heightPx?: number; align?: "left" | "center" | "right" }
  | { type: "smartArt"; smartArt: { layout: "process" | "cycle" | "hierarchy" | "list"; items: string[] }; widthPx?: number; heightPx?: number; align?: "left" | "center" | "right" }
  | { type: "image"; assetRef: string; widthPx: number; heightPx: number; alt?: string; align?: "left" | "center" | "right"; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; position?: { xPx: number; yPx: number } }
  | { type: "shape"; preset: string; text?: string; textStyle?: AgentComposeTextStyle; widthPx?: number; heightPx?: number; position?: { xPx: number; yPx: number }; fill?: string | null; line?: { color: string; widthPx: number; dash: "solid" | "dashed" | "dotted" } | null; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; order?: "front" | "back" }
  | { type: "wordArt"; text: string; preset: "plain" | "archUp" | "archDown" | "wave" | "chevron"; widthPx?: number; heightPx?: number; position?: { xPx: number; yPx: number }; rotation?: number; fill?: string; opacity?: number; wrap?: "inline" | "square" | "topAndBottom" | "none" | "behind"; order?: "front" | "back" }
  | { type: "pageNumber"; fieldKind: "page" | "pageOfTotal"; align?: "left" | "center" | "right"; color?: string; fontSizePt?: number; fontFamily?: string; bold?: boolean }
  | { type: "pageBreak" };

export interface AgentComposeRequest {
  revision: string;
  body: AgentComposeBlock[];
  header?: AgentComposeBlock[];
  footer?: AgentComposeBlock[];
  firstHeader?: AgentComposeBlock[];
  firstFooter?: AgentComposeBlock[];
  page?: {
    margins?: { top?: number; right?: number; bottom?: number; left?: number };
    size?: { width: number; height: number };
    orientation?: "portrait" | "landscape";
    columns?: number;
    columnSeparator?: boolean;
    mirrorMargins?: boolean;
  };
}

export interface AgentComposeResult {
  revision: string;
  status: "applied";
  blocks: number;
  components: number;
  overview: AgentOverview;
  createdObjects: AgentComponentSummary[];
  createdObjectsTruncated: boolean;
}

export interface AgentCollaborativeTarget {
  getDocument(): DocxDocument | null;
  getRevision(): string | number;
  allocateIds(count: number): number[];
  submit(operation: IntentBody): void | Promise<void>;
  setPresence?(position: PresencePosition | null): void;
  uploadMedia?(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number; iv?: string } | null>;
  getConnectionState?(): "local" | "live" | "offline" | "reconnecting";
}

export interface AgentProvenance {
  author: string;
  now?: () => string;
  nextId?: () => string;
}

export interface AgentTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
  execute(input: unknown): Promise<unknown>;
}
