import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from "react";
import type { EditorIntent } from "@wordinweb/core";
import {
  computePresenceCarets,
  computePresenceSelections,
  drawPresenceCarets,
  drawPresenceSelections,
} from "./presence-cursors.js";
import type { PresencePosition } from "@wordinweb/collab/client";
import {
  type SelectionSegment,
  collectRevisions,
  DocxDocument,
  DocxEditor,
  EditHistory,
  LineNumberingPatch,
  PageLayoutPatch,
  ParagraphAlignment,
  RenderHandle,
  RunFormatPatch,
  SelectionFormat,
  TableOp,
  applyRunFormat,
  applyTableOp,
  attr,
  localName,
  cellContextOf,
  cellShadingAt,
  listTableStyles,
  readTableProperties,
  setTableBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableHeaderRows,
  setTableLayoutMode,
  setTableLook,
  sortTableRows,
  insertTableFormula,
  setTableStyle,
  setTableWidth,
  tableLookOf,
  addComment,
  adjustIndent,
  paragraphDividerAt,
  deleteComment,
  compileReplaceAll,
  compileReplaceMatch,
  findAll,
  linkAt,
  removeLink,
  replaceMatch,
  replaceAll,
  requestTextInputDialog,
  setLink,
  setParagraphDivider,
  setParagraphSpacing,
  setTabStops,
  tabStopsAt,
  type TabStopSpec,
  setParagraphBorders,
  paragraphBordersAt,
  type ParagraphBorderEdge,
  type ParagraphBordersPatch,
  setDropCapAt,
  transformCase,
  exactLineHeightAt,
  replyToComment,
  insertImageAt,
  setImageWrap,
  insertEndnote,
  insertFootnote,
  insertField,
  insertCitationField,
  insertDateTimeField,
  insertPageField,
  bibliographyEntryCount,
  citationText,
  documentBibliography,
  documentMergeFieldNames,
  findBibliographyFields,
  insertBibliography,
  refreshBibliographies,
  findIndexFields,
  indexEntryCount,
  insertIndex,
  insertIndexEntry,
  isValidIndexEntry,
  refreshIndexes,
  createCitationSource,
  editCitationSource,
  deleteCitationSource,
  setCitationStyle,
  type BibliographySource,
  type CitationSourcePatch,
  type CitationSourceSpec,
  type CitationStyle,
  listBuildingBlocks,
  createBuildingBlock,
  insertBuildingBlock,
  deleteBuildingBlock,
  buildingBlockNodeCount,
  selectionClipboardBlocks,
  encodeClipboardOoxml,
  type BuildingBlockInfo,
  listBookmarks,
  bookmarkTextTarget,
  insertBookmarkAroundSelection,
  insertBookmarkAt,
  insertCrossReference,
  insertCaptionAt,
  listCrossRefTargets,
  ensureRefBookmark,
  nextRefBookmarkName,
  type CrossRefTarget,
  insertMathAt,
  insertShapeAt,
  setDrawingLineStyle,
  insertWordArtAt,
  insertChartAt,
  setChartData,
  insertSmartArtAt,
  setSmartArtData,
  insertEmbeddedObjectAt,
  insertModel3DAt,
  insertWebVideoAt,
  insertBreakAt,
  insertBlankPageAt,
  insertCoverPage,
  insertSectionBreak,
  sectPrAt,
  sectionContextAt,
  setLineNumbering,
  lineNumberingAt,
  setTitlePage,
  titlePageEnabled,
  setEvenOddHeaders,
  setPageNumberFormat,
  pageNumberFormatAt,
  setFootnoteOptions,
  footnoteOptionsAt,
  setEndnoteOptions,
  endnoteOptionsAt,
  setCommentResolved,
  editCommentText,
  documentTextStatistics,
  type PageNumberFormat,
  type PageNumberFormatPatch,
  type FootnoteOptionsPatch,
  type EndnoteOptionsPatch,
  type XmlElement,
  type EncodedCaret,
  type WireRange,
  resolveWireRange,
  type ShapePreset,
  type WordArtPreset,
  type WordArtStyle,
  type DrawingTool,
  type DrawingLineDash,
  type CoverPageContent,
  type PageNumberGalleryPosition,
  type PageNumberGalleryAlign,
  type HeaderFooterPreset,
  insertPageNumberPosition,
  removePageNumberFields,
  insertHeaderFooterPreset,
  type ObjectArrangeAction,
  type SelectedObjectCommand,
  type SelectedObjectKind,
  type ParagraphDivider,
  type FindOptions,
  type ReplaceAllResult,
  type ChartData,
  type SmartArtData,
  type SmartArtTextFormat,
  insertTableAfter,
  convertTextToTable,
  convertTableToText,
  plainTextOf,
  operationBody,
  documentOperationBody,
  type RegisteredOperationArgs,
  type RegisteredOperationKind,
  applyFieldResults,
  computeFieldResults,
  findTocFields,
  insertToc,
  insertWatermark,
  removeWatermark,
  type WatermarkSpec,
  tocEntryCount,
  rebuildToc,
  type TocOptions,
  createMeasurer,
  type TextMeasurer,
  layoutDocument,
  layoutDocumentAsync,
  type MergeRecord,
  relayoutHeadersFooters,
  detectMissingFonts,
  type MissingFont,
  type LayoutResult,
  listTypeAt,
  topLevelBlockOf,
  paragraphOf,
  printPages,
  buildPrintHtml,
  setListType,
  paragraphStyleIdOf,
  renderToDom,
  selectionToSegments,
  setPageLayout,
  setParagraphAlignment,
  setParagraphStyle,
  createStyle,
  modifyStyle,
  deleteStyle,
  listStyles,
  setNumberingLevelAt,
  NUMBERING_PRESETS,
  type NumberingPresetId,
  restartNumberingAt,
  continueNumberingAt,
  formatPatchFrom,
  summarizeSelection,
  suggestMeta,
  runWireLength,
  wireOffsetOf,
  type HostCommand,
} from "@wordinweb/core";
import type {
  StyleGalleryEntry,
  StyleSpec,
  StylePatch,
  LevelPatch,
  CellMarginsPt,
  TableBorderEdge,
  TableBorderSpec,
  TableLookToggles,
  TablePropertiesPt,
  RevisionMeta,
} from "@wordinweb/core";

async function objectPoster(title: string, subtitle: string, glyph: string): Promise<Blob> {
  const canvas = document.createElement("canvas");
  canvas.width = 640;
  canvas.height = 360;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas is unavailable");
  context.fillStyle = "#f3f6fb";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#2e74b5";
  context.fillRect(0, 0, 12, canvas.height);
  context.fillStyle = "#2e74b5";
  context.font = "bold 92px Arial, sans-serif";
  context.fillText(glyph, 46, 143);
  context.fillStyle = "#172b4d";
  context.font = "bold 34px Arial, sans-serif";
  context.fillText(title, 46, 220);
  context.fillStyle = "#52616f";
  context.font = "22px Arial, sans-serif";
  context.fillText(subtitle, 46, 260, 540);
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!blob) throw new Error("Could not create object preview");
  return blob;
}

export interface DocxViewApi {
  /** Apply character formatting to the current browser selection. */
  applyFormat(patch: RunFormatPatch): void;
  /** Create a review comment on the current selection. False if no selection. */
  addComment(text: string): boolean;
  /** Insert a footnote at the caret. False without a caret. */
  addFootnote(text: string): boolean;
  /** Insert an endnote at the caret. False without a caret. */
  addEndnote(text: string): boolean;
  /** Insert a dynamic page-number field at the caret (body, header or footer). */
  insertPageNumber(kind?: "page" | "pageOfTotal"): boolean;
  /** Insert any Word field instruction supported by the renderer. */
  insertField(instruction: string, cachedResult?: string): boolean;
  /** Insert a live DATE or TIME field with an optional Word date picture. */
  insertDateTime(kind: "date" | "time", picture?: string): boolean;
  /** Named bookmarks in document order. */
  listBookmarks(): string[];
  /** Add a bookmark around the selection, or a zero-length bookmark at the caret. */
  addBookmark(name: string): boolean;
  /** Insert a live text or page reference to a bookmark. */
  insertCrossReference(bookmark: string, kind: "text" | "page"): boolean;
  /**
   * Cross-reference targets beyond plain bookmarks, in document order:
   * headings, captions (SEQ paragraphs), and numbered list items — what
   * Word's Cross-reference dialog lists per reference type. The keys stay
   * valid until the next call (or a document change).
   */
  listCrossRefTargets(): { key: string; kind: "heading" | "caption" | "numberedItem"; text: string }[];
  /**
   * Insert a live text or page reference to a listed target. A target with
   * no hidden `_Ref` bookmark gets one first (Word's own mechanism), so the
   * REF/PAGEREF has something durable to point at.
   */
  insertCrossRefToTarget(key: string, kind: "text" | "page"): boolean;
  /**
   * Insert a Word caption — "<label> <n>" in the Caption style, the number a
   * live SEQ field — below or above the selected object (or the caret's
   * block; a caret inside a table captions the table). updateFields
   * renumbers the whole label sequence.
   */
  insertCaption(label: string, text?: string, position?: "below" | "above"): boolean;
  /**
   * Recompute every supported field's cached result — Word's F9, and what a
   * host calls before printing or exporting so the file it hands on carries
   * current page numbers, dates and cross-references. False when nothing
   * changed.
   *
   * `fileName` and `author` are the host's to supply: the engine has no
   * filesystem of its own, so FILENAME and AUTHOR fields keep their cached
   * results without them.
   */
  updateFields(values?: { fileName?: string; author?: string }): boolean;
  /**
   * Insert a native Word table of contents at the caret.
   *
   * Outside a room the page numbers are filled in immediately, from a layout
   * this call runs. IN a room the entries land with placeholders: page numbers
   * come from a layout, a layout depends on the host's font metrics, and that
   * is precisely the value `updateFields` carries as data instead of letting
   * each replica recompute it. Run the update pass to fill them.
   */
  insertToc(options?: TocOptions): boolean;
  /**
   * Rebuild every table of contents from the document's current headings, then
   * repage it. Use after headings are added, removed or retitled; repaging an
   * unchanged heading set needs only updateFields. False in a shared document,
   * for the reason insertToc gives.
   */
  refreshTocs(): boolean;
  /**
   * Every data column this document's MERGEFIELD fields name, in document
   * order, once each — headers and footers included.
   *
   * A mail-merge UI compares this against its data source's headers to report
   * which fields the data does not supply. Those keep their «Name» placeholder
   * while previewing, so the user sees an unbound field instead of a blank.
   *
   * Read-only, like the whole preview path: no operation, no undo entry, no
   * collab traffic. Set the values through the `mergeRecord` prop.
   */
  listMergeFieldNames(): string[];
  /** The bibliography sources the document's sources part declares, in part
   * order — what a citation picker and the source manager list. */
  listCitationSources(): BibliographySource[];
  /** The selected citation style (b:Sources/@StyleName — "APA", "MLA", or
   * whatever an arriving document declares), or null with no sources part. */
  getCitationStyle(): string | null;
  /** Select the citation style (APA or MLA) for citations and the
   * bibliography, creating the sources part when the document has none.
   * Refreshes citation displays and bibliographies in the same call. */
  setCitationStyle(style: CitationStyle): boolean;
  /** Add a bibliography source (book, article, website, or report), creating
   * the sources part on first use. False when the tag is already taken. */
  createCitationSource(spec: CitationSourceSpec): boolean;
  /** Change a source's fields, by tag. Refreshes citation displays and
   * bibliographies in the same call. */
  editCitationSource(tag: string, patch: CitationSourcePatch): boolean;
  /** Delete a source. False while a CITATION field still cites it — remove
   * the citations first. */
  deleteCitationSource(tag: string): boolean;
  /** Insert a citation to an existing source at the caret, displayed in the
   * document's citation style. False for a tag the sources part lacks. */
  insertCitation(tag: string): boolean;
  /** Insert a bibliography at the caret: a BIBLIOGRAPHY field whose entries
   * are generated from the document's sources (regenerated by updateFields). */
  insertBibliography(): boolean;
  /** The Quick Parts (building blocks) the document's glossary part declares,
   * name and category, in part order — what the gallery lists. */
  listBuildingBlocks(): BuildingBlockInfo[];
  /** Save the current selection as a named Quick Part, creating the glossary
   * part on first use. False with nothing selected, a name already taken, or
   * a selection the paste-fragment gate refuses (the pasteBlocks gate). */
  createBuildingBlock(name: string, category?: string): boolean;
  /** Insert a named Quick Part at the caret, cloning its stored content.
   * False for a name the glossary part lacks. */
  insertBuildingBlock(name: string): boolean;
  /** Delete a Quick Part from the glossary part, by name. False when the name
   * names none. */
  deleteBuildingBlock(name: string): boolean;
  /** Mark an index entry: an invisible XE field after the selection (or the
   * caret). With no `entry` the selected text is the entry; a colon makes a
   * subentry ("Widgets:assembly"). False with nothing to mark. */
  addIndexEntry(entry?: string): boolean;
  /** Insert an index at the caret, built from the document's XE entry marks
   * (alphabetized; page numbers filled by the update pass). Rebuilt from the
   * current marks by updateFields. */
  insertIndex(): boolean;
  /** Insert an editable inline equation from WordInWeb's linear math syntax. */
  insertEquation(linear: string): boolean;
  /** Insert a Unicode symbol through the normal undo/suggestion-aware typing path. */
  insertSymbol(symbol: string): boolean;
  /** Insert a floating editable DrawingML shape at the caret. */
  insertShape(
    preset: ShapePreset,
    text?: string,
    lineStyle?: { color: string; width: number; dash: DrawingLineDash },
  ): boolean;
  /** Insert editable DrawingML WordArt at the caret. */
  insertWordArt(text: string, preset?: WordArtPreset, style?: WordArtStyle): boolean;
  /** Insert a native editable ChartML chart at the caret. */
  insertChart(data: ChartData): boolean;
  /** Replace the selected native chart's type and data. */
  updateSelectedChart(data: ChartData): boolean;
  /** Data for the selected native chart, or null when another object is selected. */
  getSelectedChart(): ChartData | null;
  /** Insert a native editable SmartArt diagram at the caret. */
  insertSmartArt(data: SmartArtData): boolean;
  /** Replace the selected native SmartArt diagram's layout and node text. */
  updateSelectedSmartArt(data: SmartArtData): boolean;
  /** Data for the selected native SmartArt diagram, or null when another object is selected. */
  getSelectedSmartArt(): SmartArtData | null;
  /** Text formatting for the selected SmartArt node, or the first node when the group is selected. */
  getSelectedSmartArtTextFormat(): SmartArtTextFormat | null;
  /** Format the selected SmartArt node, or every node when the group is selected. */
  setSelectedSmartArtTextFormat(patch: Partial<SmartArtTextFormat>): boolean;
  /** Insert a native Office 3D model with an optional custom poster image. */
  insertModel3D(file: Blob, poster?: Blob): Promise<boolean>;
  /** Insert Word online-video metadata with a browser-safe poster. */
  insertOnlineVideo(url: string): Promise<boolean>;
  /** Embed an arbitrary file as a native OLE Package object. */
  insertEmbeddedObject(file: Blob, filename?: string): Promise<boolean>;
  /** Activate a freehand pen, or return to selection mode with null. */
  setDrawingTool(tool: DrawingTool | null): void;
  /** Current freehand pen, or null while in selection mode. */
  getDrawingTool(): DrawingTool | null;
  /** Align, rotate, or reorder the selected image, shape, or ink group. */
  arrangeObject(action: ObjectArrangeAction): boolean;
  /** True while an image, shape, or ink group is selected. */
  hasSelectedObject(): boolean;
  /** Kind of the selected object, used by contextual formatting controls. */
  getSelectedObjectContext(): { kind: SelectedObjectKind; canEditText: boolean; smartArtNodeSelected?: boolean; smartArtNodeIndex?: number } | null;
  /** Run a formatting command against the current selected object. */
  runSelectedObjectCommand(command: SelectedObjectCommand): boolean;
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  canRedo(): boolean;
  /** Insert a rows×cols table at the caret's paragraph. */
  insertTable(rows: number, cols: number): void;
  /** Row/column/table operations on the table containing the caret. */
  tableOp(op: TableOp): void;
  /** Current table-cell fill, undefined when the caret is outside a table. */
  getTableCellFill(): string | null | undefined;
  /** What a Table Properties dialog prefills from — the table's width, its
   * grid columns, the caret's column, the default cell margins and the header
   * band, in points. Undefined when the caret is outside a table. */
  getTableProperties(): TablePropertiesPt | undefined;
  /** Set or clear border edges on the caret's cell or its whole table. A null
   * `border` removes the edges so they inherit again; `{ style: "none" }`
   * instead suppresses them, which is Word's No Border. */
  setTableBorders(scope: "cell" | "table", edges: TableBorderEdge[], border: TableBorderSpec | null): void;
  /** The table styles this document's styles.xml defines — the whole valid
   * domain for setTableStyle, since an undefined id renders nothing. */
  listTableStyles(): { id: string; name: string }[];
  /** Style id of the table containing the caret: null when it has none,
   * undefined when the caret is outside a table. */
  getTableStyleId(): string | null | undefined;
  /** Apply a table style, or with null remove the reference. */
  setTableStyle(styleId: string | null): void;
  /** Word's six table style options for the caret's table, undefined outside
   * a table. */
  getTableLook(): TableLookToggles | undefined;
  /** Set some of the six table style options. */
  setTableLook(patch: Partial<TableLookToggles>): void;
  /** Set the table's preferred width. `value` is points for "pt", 0-100 for
   * "pct", and ignored for "auto". */
  setTableWidth(unit: "pt" | "pct" | "auto", value?: number): void;
  /** Set one column to an exact width in points. */
  setTableColumnWidth(colIdx: number, widthPt: number): void;
  /** Switch the table between fixed column widths and autofit. Switching to
   * fixed freezes the columns as currently RENDERED, which is what the user
   * is looking at; switching to autofit measures them from content again. */
  setTableLayout(layout: "fixed" | "autofit"): void;
  /** Set the table's default cell margins, or the caret cell's override, in
   * points. A null patch drops the override. */
  setTableCellMargins(scope: "cell" | "table", margins: CellMarginsPt | null): void;
  /** Repeat the first `count` rows as a header band on every page. */
  setTableHeaderRows(count: number): void;
  /** Sort the caret table's body rows by one grid column, as text or as
   * numbers. The repeating header band always stays in place; `hasHeader`
   * additionally pins the first row. Refuses tables with merged cells. */
  sortTableRows(colIdx: number, order: "asc" | "desc", compare: "text" | "number", hasHeader?: boolean): void;
  /** Convert the selected paragraphs (or the caret paragraph) into a table,
   * one row per paragraph, cells split on the separator. */
  convertTextToTable(separator: "tab" | "comma"): boolean;
  /** Convert the caret's table into paragraphs, one per row, cell texts
   * joined by the separator. */
  convertTableToText(separator: "tab" | "comma"): boolean;
  /** Insert a table formula field ("=SUM(ABOVE)", "=A1+B2", …) at the caret's
   * cell paragraph, with an optional \# number format like "#,##0.00". The
   * result is evaluated immediately; updateFields recomputes it. */
  insertTableFormula(formula: string, numFmt?: string): boolean;
  /** Insert an image file at the caret (inline, natural size clamped to column).
   * The result is REPORTED rather than swallowed: a picker that accepts a file
   * and then does nothing is indistinguishable from a broken button. */
  insertImage(file: Blob): Promise<ImageInsertResult>;
  /** What this document's image picker should ADVERTISE, as a file-input
   * `accept` string — narrower in a shared document, where the format has to
   * survive the wire. Derived from the same list the insert guard consults, so
   * the picker cannot offer something the insert will refuse. */
  imageAccept(): string;
  /** The relay's published per-image byte limit, or null when there is none to
   * show (a local document, or a server that publishes no limit). Null means
   * "say nothing about size" — never substitute a default. */
  imageMaxBytes(): number | null;
  /** Capture a screen, window, or browser tab and insert the current frame as a PNG picture. */
  insertScreenshot(): Promise<ScreenshotInsertResult>;
  /** Align the paragraph(s) under the caret or selection. */
  setAlignment(align: ParagraphAlignment): void;
  /** Apply a named paragraph style (null clears back to Normal). */
  setParagraphStyle(styleId: string | null): void;
  /** Toggle bulleted/numbered list on the paragraph(s) under the selection. */
  toggleList(kind: "bullet" | "number"): void;
  /** Current list kind at the caret ("bullet" | "number" | null). */
  getListType(): "bullet" | "number" | null;
  /** Link the selection to a URL; null removes the link at the caret. */
  setLink(url: string | null): void;
  /** URL of the hyperlink at the caret/selection, or null. */
  getLinkAt(): string | null;
  /** Step paragraph indent by half an inch (Word's indent buttons). */
  adjustIndent(direction: 1 | -1): void;
  /** Line spacing multiple or exact point height, and/or space before/after (points). */
  setParagraphSpacing(patch: { lineMultiple?: number; exactLinePt?: number; beforePt?: number | null; afterPt?: number | null }): void;
  /** Create, customize, or remove the bottom-border divider on the selected paragraph(s). */
  setParagraphDivider(divider: ParagraphDivider | null): boolean;
  /** Direct bottom-border divider on the caret paragraph. */
  getParagraphDivider(): ParagraphDivider | null;
  /** Replace the selected paragraphs' direct tab stops (w:tabs). An empty
   * list removes them, so style stops and the default grid apply again. */
  setTabStops(stops: TabStopSpec[]): boolean;
  /** Direct tab stops on the caret paragraph, sorted by position. */
  getTabStops(): TabStopSpec[];
  /** Patch the selected paragraphs' border edges (w:pBdr) and/or shading
   * fill (w:shd). Each named edge is set (spec) or cleared (null); absent
   * edges stay. `shading: null` removes the fill. */
  setParagraphBorders(patch: ParagraphBordersPatch): boolean;
  /** Direct borders and shading on the caret paragraph. */
  getParagraphBorders(): { borders: Partial<Record<ParagraphBorderEdge, TableBorderSpec>>; shading: string | null };
  /** Apply or remove a native Word drop cap on the caret paragraph. */
  setDropCap(mode: "drop" | "margin" | null, lines?: number): boolean;
  /** Remove direct character formatting from the selection. */
  clearFormatting(): void;
  /** Change the selection's case. */
  changeCase(mode: "upper" | "lower" | "title"): void;
  /** Find matches for a query across every story (body, headers, footers,
   * footnotes, endnotes); selects the first and returns the count. */
  find(query: string, opts?: FindOptions): number;
  /** Select the next/previous match; returns 1-based index or 0. */
  findStep(delta: 1 | -1): number;
  /** Replace the current match; returns remaining match count. */
  replaceCurrent(replacement: string): number;
  /** Replace every match; reports how many were replaced, per story. */
  replaceAll(query: string, replacement: string, opts?: FindOptions): ReplaceAllResult;
  /** Go To: scroll the given 1-based page into view. */
  goToPage(page: number): boolean;
  /** Go To: move the caret to a named bookmark and scroll it into view. */
  goToBookmark(name: string): boolean;
  /**
   * Select an exact text range by stable address — the `{blockId, runId,
   * start, end}` shape the wire, presence, and suggestRevision already use
   * (offsets in the run's wire basis; see getEncodedCaret for the caret
   * half). Several ranges select together, for a word split across runs.
   * Selection and scroll ride find-navigation's machinery; the view only
   * scrolls when the range is off screen. In a local document the first call
   * populates the stable-id table (enableStableIds), so hosts — the desktop
   * spellcheck's select-and-replace — can address text without a collab
   * mount. False when no range resolves to text.
   */
  selectRange(range: WireRange | WireRange[]): boolean;
  /** Paragraph styles for the style menu (declared + Word built-ins). */
  listParagraphStyles(): { id: string; name: string }[];
  /** pStyle id of the caret paragraph (null = Normal). */
  getParagraphStyleId(): string | null;
  /**
   * Every declared style with the data a gallery entry needs: identity,
   * cascade parent, quick-style flag, usage count, and resolved preview props.
   * Richer than listParagraphStyles, which is the flat menu the toolbar's
   * existing style dropdown reads.
   */
  listStyles(filter?: { type?: StyleGalleryEntry["type"]; quickStyleOnly?: boolean }): StyleGalleryEntry[];
  /** Add a paragraph or character style definition. */
  createStyle(spec: StyleSpec): boolean;
  /** Patch an existing style definition; the cascade re-resolves live. */
  modifyStyle(styleId: string, patch: StylePatch): boolean;
  /** Delete a style definition; its users fall back to its basedOn. */
  deleteStyle(styleId: string): boolean;
  /**
   * Stamp a text watermark across every page. Word keeps a watermark in the
   * header parts, so a document with none gets one first.
   */
  insertWatermark(spec: WatermarkSpec): boolean;
  /** Take the watermark back off every page. False when there was none. */
  removeWatermark(): boolean;
  /** Apply (or with null remove) a character style over the selection. */
  setCharacterStyle(styleId: string | null): void;
  /** Change a list level's number format, label text, or indent. */
  setNumberingLevel(ilvl: number | null, patch: LevelPatch): boolean;
  /**
   * Apply a preset multilevel definition (Word's multilevel gallery) to the
   * caret's list — the paragraphs join a numbered list first when they are
   * not in one. Compiles onto the existing per-level patch operation, so it
   * rides the same wire ops and converges like any level edit.
   */
  applyNumberingPreset(preset: NumberingPresetId): boolean;
  /** Restart list numbering at the caret, or (null) continue the preceding list. */
  setNumberingRestart(start: number | null): boolean;
  /** Format painter, half one: the selection's formatting, or null. */
  copyFormatting(): SelectionFormat | null;
  /** Format painter, half two: paint a copied format over the selection. */
  applyCopiedFormatting(format: SelectionFormat): void;
  /** Change margins / page size / orientation (inches). */
  setPageLayout(patch: PageLayoutPatch, scope?: "document" | "section"): void;
  /** One-based logical section containing the caret or selection. */
  getSectionContext(): { index: number; count: number } | null;
  /** Insert a page/column break or a section break at the caret. */
  insertBreak(kind: "page" | "column" | "sectionNextPage" | "sectionContinuous"): boolean;
  /** Insert a full blank page at the caret (two consecutive page breaks). */
  insertBlankPage(): boolean;
  /** Insert an editable cover page before the current document. */
  insertCoverPage(content: CoverPageContent): boolean;
  /** Toggle/configure margin line numbers (Word's Layout > Line Numbers).
   * scope "section" targets the caret's section; "document" every section. */
  setLineNumbering(patch: LineNumberingPatch, scope?: "document" | "section"): void;
  /** Current line-numbering settings for the caret's section, or null (off). */
  getLineNumbering(): { countBy: number; restart: "continuous" | "newPage" | "newSection"; start: number } | null;
  /** Leave header/footer editing mode. */
  closeHeaderFooter(): void;
  /** Enter and, if needed, create the header or footer on the caret's page. */
  openHeaderFooter(kind: "header" | "footer"): boolean;
  /** Word's "Different First Page" (w:titlePg). Enabling creates the empty
   * first-page header/footer parts; disabling keeps them for re-enable. */
  setDifferentFirstPage(on: boolean): boolean;
  /** Whether any section requests a different first page. */
  getDifferentFirstPage(): boolean;
  /** Word's "Different Odd & Even Pages" (settings.xml w:evenAndOddHeaders).
   * Enabling creates the empty even-page header/footer parts. */
  setOddEvenHeaders(on: boolean): boolean;
  getOddEvenHeaders(): boolean;
  /** Page-number format and start-at (w:pgNumType): decimal / roman / letter
   * numbering, restart value. scope "section" targets the caret's section;
   * "document" (and any shared document) every section. */
  setPageNumberFormat(patch: PageNumberFormatPatch, scope?: "document" | "section"): boolean;
  /** Current page-number settings for the caret's section. */
  getPageNumberFormat(): { fmt: PageNumberFormat; start: number | null };
  /** Patch the automatic-hyphenation settings (w:autoHyphenation,
   * w:hyphenationZone in points, w:doNotHyphenateCaps). Round-trip state:
   * this engine's layout does not hyphenate automatically — the settings
   * govern Word's own rendering of the document. */
  setHyphenation(patch: { auto?: boolean; zonePt?: number | null; noCaps?: boolean }): boolean;
  /** Current hyphenation settings (zonePt null = Word's default 18pt). */
  getHyphenation(): { auto: boolean; zonePt: number | null; noCaps: boolean };
  /** Word's page-number position gallery: a single live PAGE field in the
   * header (top) or footer (bottom), aligned left/center/right. Creates the
   * part on demand and replaces its content — a gallery pick, not a merge. */
  insertPageNumberPosition(position: PageNumberGalleryPosition, align: PageNumberGalleryAlign): boolean;
  /** Word's "Remove Page Numbers": strips PAGE/NUMPAGES fields from every
   * header and footer part. */
  removePageNumbers(): boolean;
  /** The Header & Footer preset gallery: replace a header or footer's
   * content with a preset layout (blank / centered title / title + date /
   * three-column). */
  insertHeaderFooterPreset(kind: "header" | "footer", preset: HeaderFooterPreset): boolean;
  /** Footnote options (w:footnotePr): number format, restart rule, start-at,
   * and position (§17.11.17/19/21). scope "section" targets the caret's
   * section; "document" (and any shared document) every section. Format and
   * start drive the painted marks; restart honors "eachSect" only; position
   * round-trips without changing where footnotes lay out (always page
   * bottom in this engine). */
  setFootnoteOptions(patch: FootnoteOptionsPatch, scope?: "document" | "section"): boolean;
  /** Current footnote options for the caret's section. */
  getFootnoteOptions(): ReturnType<typeof footnoteOptionsAt>;
  /** Same as setFootnoteOptions, for endnotes — position vocabulary differs
   * (sectEnd / docEnd). */
  setEndnoteOptions(patch: EndnoteOptionsPatch, scope?: "document" | "section"): boolean;
  getEndnoteOptions(): ReturnType<typeof endnoteOptionsAt>;
  /** Resolve (true) or reopen (false) a comment thread. */
  resolveComment(id: string, resolved: boolean): boolean;
  /** Replace a comment's body text (Word's edit-my-comment). */
  editComment(id: string, text: string): boolean;
  /** Select and scroll to the next/previous commented range (Review-tab
   * navigation). Returns the focused thread's comment id, or null when the
   * document has no anchored comments. */
  stepComment(delta: 1 | -1): string | null;
  /** Word Count: body text statistics from the model plus the page count
   * from the latest layout. */
  wordCount(): { words: number; characters: number; charactersWithSpaces: number; paragraphs: number; pages: number };
  /** Effective formatting of the current selection (toolbar state), or null. */
  getSelectionFormat(): SelectionFormat | null;
  /** Print the rendered pages (browser print dialog / save as PDF). */
  print(): void;
  /**
   * The standalone print document (all pages plus styles) as an HTML string,
   * for hosts that produce the PDF themselves (e.g. a desktop shell).
   * Null before the first render.
   */
  exportPrintHtml(): string | null;
  /** Serialize the (edited) document back to .docx bytes. */
  save(): Uint8Array;
  /** Page count after the latest layout. */
  pageCount(): number;
  /**
   * Suggesting mode: when on, edits record as OOXML tracked changes (w:ins /
   * w:del) instead of mutating text directly, and the view switches to markup
   * so the suggestion shows live. `author` stamps the revision (defaults to the
   * commentAuthor prop). Turning it off restores the prior revision view.
   */
  setSuggesting(on: boolean, author?: string): void;
  isSuggesting(): boolean;
  /** Accept the tracked change at the caret (keep insertion / apply deletion). */
  acceptRevisionAtCaret(): boolean;
  /** Reject the tracked change at the caret (drop insertion / restore deletion). */
  rejectRevisionAtCaret(): boolean;
  /** How many tracked changes (suggestions) the document currently holds. */
  revisionCount(): number;
  /** Accept every tracked change (one undo step). Returns how many applied. */
  acceptAllRevisions(): number;
  /** Reject every tracked change (one undo step). Returns how many applied. */
  rejectAllRevisions(): number;
  /** Current caret as stable-id addresses, or null. The encoding survives a
   * reconciliation reload, so it can be captured from a view about to
   * remount and restored into its replacement. In a local document the first
   * call populates the stable-id table (like selectRange), so hosts can
   * capture addresses outside a collab mount too. */
  getEncodedCaret(): EncodedCaret | null;
  /** Restore a caret captured by getEncodedCaret. False when the position no
   * longer resolves (or outside collab mode). */
  setCaretFromEncoded(pos: EncodedCaret): boolean;
  /**
   * Scroll the view so `participant`'s live presence caret is on screen —
   * "jump to that person". A pure VIEW operation: it never relayouts, marks
   * nothing dirty, and mutates no document state. The jump is instant, never
   * smooth: the target can be hundreds of virtualized pages away, and a
   * smooth scroll across that distance forces continuous page mounting;
   * instant also satisfies prefers-reduced-motion by construction.
   *
   *  - "revealed"    scrolled to the caret (or, when the exact run was
   *                  deleted since it was reported, to the start of its
   *                  paragraph — the same fallback caret restore uses).
   *  - "no-position" that participant has no broadcast cursor right now
   *                  (never placed one, cleared it, or left).
   *  - "unresolved"  the position no longer maps to content in this replica.
   */
  revealPresence(participant: string): "revealed" | "no-position" | "unresolved";
  document: DocxDocument;
}

export type ScreenshotInsertResult = "inserted" | "unsupported" | "cancelled" | "error" | "no-caret";

/**
 * Word's multilevel-list gallery, expressed as per-level patches over the
 * existing deep numbering ops. One entry per preset; index = ilvl.
 */
export { NUMBERING_PRESETS } from "@wordinweb/core";
export type { NumberingPresetId } from "@wordinweb/core";

/**
 * Image formats a SHARED document accepts, and the single source of truth for
 * both halves of that promise: the guard that declines an intent and the
 * `accept` attribute of the file picker that offers one.
 *
 * They were separately written constants once, and drifted — the picker
 * advertised SVG while the guard refused it, so choosing an SVG in a shared
 * document did nothing at all, with no skeleton, no message and no log. The
 * user's report was "it just didn't work and then disappeared". Deriving the
 * offer from the acceptance is what makes that drift unrepresentable.
 *
 * Raster-only because the wire allowlist is (collab/src/validate.ts): widening
 * it is an intent-shape change with an ENGINE_VERSION bump, not a UI edit.
 */
export const COLLAB_IMAGE_EXTS = ["png", "jpg", "jpeg", "gif", "bmp", "webp"] as const;
const COLLAB_IMAGE_ACCEPT = "image/png,image/jpeg,image/gif,image/bmp,image/webp";
/** Local documents additionally take SVG — nothing has to agree with them. */
const LOCAL_IMAGE_ACCEPT = `${COLLAB_IMAGE_ACCEPT},image/svg+xml`;

/**
 * Why an image insert did not happen, so the caller can SAY so. Every one of
 * these was a bare `return` until a user hit the `unsupported-format` path and
 * had no way to tell a rejected file from a broken button.
 */
export type ImageInsertResult =
  | "inserted"
  | "no-caret"
  /** Not a format a shared document can carry (SVG today). */
  | "unsupported-format"
  /** Bigger than the relay's published per-blob limit. Refused BEFORE any
   * bytes were read, sealed, hashed or sent — see `imageMaxBytes()` for the
   * number to show the user. */
  | "too-large"
  /** Collab is wired but no media relay is configured — nothing can be
   * uploaded, and inserting locally would fork the room. */
  | "no-relay"
  /** The relay refused the bytes; nothing was reserved, nothing forked. */
  | "upload-failed"
  /** The bytes did not decode as an image at all. */
  | "error";

export interface DocxViewProps {
  /** The document: raw bytes, a File/Blob, or a URL to fetch. */
  source: ArrayBuffer | Uint8Array | Blob | string;
  /** Zoom factor, 1 = 100%. */
  zoom?: number;
  /**
   * Fit-to-width (Google-Docs mobile behavior): when the page is wider than the
   * viewport, auto-scale down so it fits with a small gutter and never scrolls
   * horizontally. The computed scale drives the real `zoom` (crisp text, not a
   * blurry transform) and is capped at `zoom`, so a wide desktop viewport is
   * unchanged. Recomputed on container resize. Default true.
   */
  fitWidth?: boolean;
  /**
   * Container width (px) at or below which the chrome switches to its compact
   * phone/tablet treatment — the comment rail collapses to tap-to-open cards so
   * balloons never force horizontal scroll. Default 820.
   */
  narrowWidth?: number;
  /**
   * Enable editing commands (selection-based formatting, save-back).
   * Default false: pure render-only viewer.
   */
  editable?: boolean;
  className?: string;
  style?: React.CSSProperties;
  onLoad?: (info: { pageCount: number; document: DocxDocument }) => void;
  /** Fires whenever editing changes the rendered page count. */
  onPageCountChange?: (pageCount: number) => void;
  /** Fires when the document is ready; the api is only usable while mounted. */
  onReady?: (api: DocxViewApi) => void;
  onError?: (error: Error) => void;
  /**
   * Optional collaborative session (from `wordinweb/collab`'s `useCollab`).
   * When provided, the editor forwards each local edit as an intent via
   * `collab.submit`. Typed structurally so the main `wordinweb` bundle carries
   * no runtime dependency on the collab engine (plan doc 07 tree-shaking) —
   * the app imports the session from the separate `wordinweb/collab` entry and
   * injects it here.
   */
  collab?: {
    /** clientId → display name for caret flags (doc 14 §2); text-node rendered. */
    participantNames?: Record<string, string>;
    submit: (intent: EditorIntent) => void;
    /** Remote participants' cursor/selection positions, drawn as colored
     * carets over the page (see presence-cursors). */
    presence?: Record<string, PresencePosition | null>;
    /** Allocate `n` fresh carried node ids (for sub-range format / split /
     * insert intents). Injected from the collab connection. */
    allocIds?: (n: number) => number[];
    /** The live reconciled document object to render DIRECTLY (skip the
     * bytes → parse round-trip). The collab replica mutates this same instance
     * in place on each broadcast; DocxView repaints it when `renderSignal`
     * bumps, so a remote edit costs one repaint — no re-serialize, no re-parse,
     * no caret reset. When present, `source` is only a placeholder. */
    doc?: DocxDocument;
    /** Monotonic counter that bumps whenever `doc` was mutated in place; a
     * change triggers an in-place repaint of `doc`. */
    renderSignal?: number;
    /** Drain the union of the dirty scopes behind `renderSignal` since the
     * last take. A narrow scope lets the repaint relayout one paragraph
     * incrementally (the same path local typing takes) instead of the whole
     * document; `doc` scope (or an absent method) keeps the whole-document
     * repaint; null means nothing is dirty and the repaint is skipped.
     * Consumed at the repaint so a coalesced repaint covers every batched
     * remote intent. */
    takeRenderScope?: () =>
      | { kind: "doc" }
      | { kind: "block"; blocks: XmlElement[] }
      | { kind: "split"; before: XmlElement; after: XmlElement }
      | null;
    /** Broadcast the local caret so remote participants draw this user's
     * cursor. Called with the caret's stable-id address on every caret move
     * (null when the caret leaves id-tracked content). */
    setPresence?: (pos: PresencePosition | null) => void;
    /** Submit a toolbar/API operation NOT yet applied locally. The connection
     * applies it optimistically through the same canonical code the server
     * runs, so the local result is byte-identical to every replica. When set,
     * DocxView routes its imperative commands (insert chart/table/equation,
     * set link/page layout, comments, ...) through this instead of mutating
     * the document itself. */
    submitOp?: (intent: { kind: string } & Record<string, unknown>) => void;
    /** Upload image bytes to the media relay and return the address fields
     * the insertImage intent must carry (plan doc 16 §5.1). Null means the
     * relay REFUSED — the caller must then not reserve anything, or the room
     * gets a skeleton nobody can ever fill. Absent when the app supplied no
     * relay origin, in which case images stay a local-only feature. */
    uploadMedia?: (bytes: Uint8Array) => Promise<{ blobSha: string; bytesLen: number; iv?: string } | null>;
    /**
     * Largest single upload the RELAY will accept, in bytes, as published in
     * the welcome. Lets the insert refuse an oversized file locally instead of
     * discovering it after sealing, hashing and a full upload.
     *
     * NULL MEANS SKIP THE CHECK — not "no limit" and not "use a default". An
     * older server publishes nothing, and a client that invents a number
     * either blocks uploads the server would have taken or promises the user
     * one it will refuse. The server enforces the real limit either way, so
     * skipping is safe and guessing is not.
     */
    mediaMaxBlobBytes?: number | null;
    /** Reverse this user's last SEQUENCED action (plan doc 03 Phase 8). The
     * editor routes Cmd+Z here in a room, because replaying the LOCAL history
     * stack would edit this replica with nothing on the wire. Absent ⇒ undo
     * declines rather than mutating. */
    undoLast?: () => void;
    /** Single-process session marker (useAgentDocumentSession over a
     * LocalDocumentSession). Its presence tells the editor there are NO peers,
     * so undo/redo replay the local history stack; it fires after each applied
     * undo/redo so the session can bump its revision. A REAL room must leave
     * this unset, or Cmd+Z would fork it. */
    noteLocalHistory?: () => void;
  };
  /** Author name stamped on comment replies (default "You"). */
  commentAuthor?: string;
  /** Connected collaborator names offered as @mention shortcuts in comments. */
  commentMentions?: string[];
  /** Render review comments (range highlights + margin balloons). Default true. */
  showComments?: boolean;
  /** Tracked-changes display: "final" (default) or "markup". */
  revisions?: "final" | "markup";
  /** Fires after render with document-requested font faces the browser cannot
   * render (unavailable, or lacking the document's script) — the page is
   * silently substituting and may differ from Word. Empty array = all good. */
  onMissingFonts?: (missing: MissingFont[]) => void;
  /**
   * Mail-merge PREVIEW: the active record's column values, painted into the
   * document's MERGEFIELD fields. Absent (or undefined) shows the «Name»
   * placeholders — that is preview "off".
   *
   * Nothing is written to the document. The values are resolved as the pages
   * are laid out, so they cannot reach a saved file, an undo entry or the
   * collab wire; preview is per-viewer state by construction. Changing this
   * prop relays the pages without re-parsing the .docx.
   *
   * A column the record does NOT carry keeps its «Name» placeholder, so the
   * user can see which fields the data leaves unbound. Word renders such a
   * field blank; this is a deliberate divergence. A column that IS present but
   * empty renders empty and suppresses the field's \b and \f switch texts.
   */
  mergeRecord?: MergeRecord;
}

export interface AgentDocumentViewBinding {
  subscribe(listener: () => void): () => void;
  getSnapshot(): number;
  doc: DocxDocument;
  submit(intent: EditorIntent): void;
  submitOp(intent: { kind: string } & Record<string, unknown>): void;
  allocIds(count: number): number[];
  uploadMedia?(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number; iv?: string } | null>;
  /** Bump the session revision after a local history undo/redo mutated the
   * document outside the intent path. See collab.noteLocalHistory. */
  noteHistory?(): void;
  takeRenderScope():
    | { kind: "doc" }
    | { kind: "block"; blocks: XmlElement[] }
    | { kind: "split"; before: XmlElement; after: XmlElement }
    | null;
}

const AGENT_VIEW_SOURCE = new Uint8Array(0);

/** Connect a framework-neutral local document session to DocxView. */
export function useAgentDocumentSession(binding: AgentDocumentViewBinding): Pick<DocxViewProps, "source" | "collab"> {
  const renderSignal = useSyncExternalStore(binding.subscribe, binding.getSnapshot, binding.getSnapshot);
  return {
    source: AGENT_VIEW_SOURCE,
    collab: {
      doc: binding.doc,
      renderSignal,
      submit: binding.submit,
      submitOp: binding.submitOp,
      allocIds: binding.allocIds,
      uploadMedia: binding.uploadMedia,
      takeRenderScope: binding.takeRenderScope,
      // A local session has no peers: undo/redo replay the editor's own
      // history stack, and the session only needs its revision bumped.
      noteLocalHistory: binding.noteHistory,
    },
  };
}

async function toBytes(source: DocxViewProps["source"]): Promise<Uint8Array> {
  if (typeof source === "string") {
    const res = await fetch(source);
    if (!res.ok) throw new Error(`Failed to fetch ${source}: ${res.status}`);
    return new Uint8Array(await res.arrayBuffer());
  }
  if (source instanceof Blob) return new Uint8Array(await source.arrayBuffer());
  if (source instanceof Uint8Array) return source;
  return new Uint8Array(source);
}

const BACKGROUND_LAYOUT_PAGE_THRESHOLD = 50;

/** Word's font-size ladder, which grow/shrink steps along. */
const FONT_SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

function steppedFontSize(current: number, direction: 1 | -1): number {
  if (direction === 1) return FONT_SIZE_STEPS.find((size) => size > current) ?? Math.min(current + 10, 1638);
  return [...FONT_SIZE_STEPS].reverse().find((size) => size < current) ?? Math.max(current - 1, 1);
}

/**
 * Run a keyboard command through the public API.
 *
 * Every case calls exactly what the matching toolbar button calls, which is
 * what makes a shortcut and its button emit the same collaboration intent and
 * record the same tracked change in suggesting mode. `painter` is the format
 * painter's clipboard, shared with nothing else — the toolbar keeps its own.
 */
function runHostCommand(
  api: DocxViewApi,
  command: Exclude<HostCommand, "link" | "comment" | "goToPage">,
  painter: { current: SelectionFormat | null },
): void {
  const format = api.getSelectionFormat();
  switch (command) {
    case "bullet":
    case "number":
      api.toggleList(command);
      break;
    case "alignLeft":
      api.setAlignment("left");
      break;
    case "alignCenter":
      api.setAlignment("center");
      break;
    case "alignRight":
      api.setAlignment("right");
      break;
    case "justify":
      api.setAlignment("justify");
      break;
    case "indentIn":
      api.adjustIndent(1);
      break;
    case "indentOut":
      api.adjustIndent(-1);
      break;
    case "lineSpacingSingle":
      api.setParagraphSpacing({ lineMultiple: 1 });
      break;
    case "lineSpacingOneAndHalf":
      api.setParagraphSpacing({ lineMultiple: 1.5 });
      break;
    case "lineSpacingDouble":
      api.setParagraphSpacing({ lineMultiple: 2 });
      break;
    case "strikethrough":
      api.applyFormat({ strike: !format?.strike });
      break;
    case "superscript":
      api.applyFormat({ verticalAlign: format?.verticalAlign === "superscript" ? null : "superscript" });
      break;
    case "subscript":
      api.applyFormat({ verticalAlign: format?.verticalAlign === "subscript" ? null : "subscript" });
      break;
    case "clearFormatting":
      api.applyFormat({ clear: true });
      break;
    case "copyFormatting":
      painter.current = api.copyFormatting();
      break;
    case "pasteFormatting":
      if (painter.current) api.applyCopiedFormatting(painter.current);
      break;
    case "growFont":
    case "shrinkFont":
      // A selection of mixed sizes reports none, and there is no one size to
      // step from — Word would grow each run separately, which no single
      // patch expresses. Declining beats picking a size the user never had.
      if (format?.fontSizePt !== undefined) {
        api.applyFormat({ fontSizePt: steppedFontSize(format.fontSizePt, command === "growFont" ? 1 : -1) });
      }
      break;
    case "nextComment":
      api.stepComment(1);
      break;
    case "previousComment":
      api.stepComment(-1);
      break;
    case "trackChanges":
      api.setSuggesting(!api.isSuggesting());
      break;
    case "tableRowBelow":
      api.tableOp("rowBelow");
      break;
  }
  // Tell the toolbar its buttons may have changed state (the same
  // announcement the editor makes after a selection-changing edit).
  document.dispatchEvent(new CustomEvent("dxw-selection"));
}

/**
 * High-fidelity paginated DOCX viewer (and, with `editable`, editor).
 *
 * ```tsx
 * <DocxView source="/report.docx" />                          // render-only
 * <DocxView source="/report.docx" editable onReady={setApi} /> // editing
 * ```
 */
export function DocxView({
  source,
  zoom = 1,
  fitWidth = true,
  narrowWidth = 820,
  editable = false,
  className,
  style,
  onLoad,
  onPageCountChange,
  onReady,
  onError,
  commentAuthor = "You",
  commentMentions = [],
  showComments = true,
  revisions = "final",
  onMissingFonts,
  mergeRecord,
  collab,
}: DocxViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const commentMentionsRef = useRef(commentMentions);
  commentMentionsRef.current = commentMentions;
  /** Mail-merge preview record, read by every layout call below. Held in a ref
   * so a record step never re-runs the load effect (which would re-parse the
   * .docx); the repaint effect further down relays the pages instead. */
  const mergeRecordRef = useRef(mergeRecord);
  mergeRecordRef.current = mergeRecord;
  /** The keyboard format painter's clipboard (⌥⌘C copies, ⌥⌘V pastes). */
  const painterRef = useRef<SelectionFormat | null>(null);
  const [error, setError] = useState<Error | null>(null);
  const [layoutBusy, setLayoutBusy] = useState(false);
  // Fit-to-width scale (page-width / container-width) recomputed on resize, and
  // the compact-chrome flag. The effective zoom driven into the renderer caps
  // the fit at the user's zoom so a roomy desktop viewport stays at 100%.
  const [fitZoom, setFitZoom] = useState<number | null>(null);
  const [narrow, setNarrow] = useState(false);
  const basePageWidthRef = useRef(816);
  const effectiveZoom = fitWidth && fitZoom != null ? Math.min(zoom, fitZoom) : zoom;
  // Latest effective zoom the render loop should paint at. Held in a ref so a
  // zoom change re-renders the pages in place (crisp) WITHOUT re-parsing the
  // document or tearing down the editor — the big effect below reads it as its
  // starting zoom and installs applyZoomRef to update it live.
  const effZoomRef = useRef(effectiveZoom);
  effZoomRef.current = effectiveZoom;
  const applyZoomRef = useRef<((z: number) => void) | null>(null);
  const onPageCountChangeRef = useRef(onPageCountChange);
  onPageCountChangeRef.current = onPageCountChange;

  const recomputeFit = useCallback(() => {
    const c = containerRef.current;
    if (!c) return;
    const cw = c.clientWidth;
    if (!cw) return;
    const gutter = 24; // breathing room so the page edge isn't flush
    setFitZoom((cw - gutter) / basePageWidthRef.current);
    setNarrow(cw <= narrowWidth);
  }, [narrowWidth]);

  // Recompute fit whenever the container resizes (orientation change, split
  // view, window drag). Container width is independent of the page's own zoom,
  // so scaling the page never feeds back into this — no resize loop.
  useEffect(() => {
    const c = containerRef.current;
    if (!c || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(() => recomputeFit());
    ro.observe(c);
    recomputeFit();
    return () => ro.disconnect();
  }, [recomputeFit]);

  // Push a zoom change into the live render loop (no reparse).
  useEffect(() => {
    applyZoomRef.current?.(effectiveZoom);
  }, [effectiveZoom]);

  // Reflect the compact flag on the container so the renderer's CSS (rail
  // collapse, tap-to-open cards) can key off it.
  useEffect(() => {
    containerRef.current?.classList.toggle("dxw-narrow", narrow);
  }, [narrow]);

  // Tap-to-open comments in compact mode: tapping commented text surfaces that
  // comment's card as a floating sheet (the rail is hidden on narrow); tapping
  // empty space or another comment dismisses it. Additive listener — it never
  // preventDefaults, so the editor's own caret-placement path is untouched.
  useEffect(() => {
    const c = containerRef.current;
    if (!c) return;
    const onTap = (e: Event) => {
      if (!c.classList.contains("dxw-narrow")) return;
      const root = c.querySelector<HTMLElement>(".dxw-pages");
      if (!root) return;
      const target = e.target as HTMLElement;
      const closeAll = () =>
        root.querySelectorAll(".dxw-comment-card.dxw-open").forEach((el) => el.classList.remove("dxw-open"));
      if (target.closest?.(".dxw-comment-card")) return; // interacting inside a card
      const span = target.closest?.("[data-dxw-comment]") as HTMLElement | null;
      closeAll();
      if (!span) return;
      const id = span.dataset.dxwComment!.split(" ")[0];
      const cards = Array.from(root.querySelectorAll<HTMLElement>(".dxw-comment-card"));
      cards.find((el) => el.dataset.dxwCommentId === id)?.classList.add("dxw-open");
    };
    c.addEventListener("click", onTap);
    return () => c.removeEventListener("click", onTap);
  }, []);
  // Contextual header/footer hotbar: the editor announces hf-mode via a
  // bubbled dxw-hfmode event; the tools that only make sense there (page
  // numbers, close) surface right where the user is editing.
  const [hfMode, setHfMode] = useState(false);
  const apiRef = useRef<DocxViewApi | null>(null);
  const handleRef = useRef<RenderHandle | null>(null);
  // Set by the main effect so the live-collab-doc repaint effect can trigger an
  // in-place re-render without re-running the whole load effect.
  const rerenderRef = useRef<
    | ((
        doc: DocxDocument,
        scope?:
          | { kind: "doc" }
          | { kind: "block"; blocks: XmlElement[] }
          | { kind: "split"; before: XmlElement; after: XmlElement },
      ) => void)
    | null
  >(null);
  const redrawPresenceRef = useRef<(() => void) | null>(null);
  // Current presence, read by the imperative draw (which closes over an older
  // render's props otherwise).
  const presenceRef = useRef<Record<string, PresencePosition | null> | undefined>(undefined);
  presenceRef.current = collab?.presence;
  // The editor and imperative API survive prop updates. Their callbacks must
  // read the current session, especially when a live room becomes offline.
  const collabRef = useRef(collab);
  collabRef.current = collab;

  // Redraw remote presence carets whenever the presence prop changes.
  useEffect(() => {
    redrawPresenceRef.current?.();
  }, [collab?.presence]);
  // The parsed document (and its undo history) survives re-runs of the main
  // effect: toggling `editable` (Editing↔Viewing), `revisions` (markup/final)
  // or `commentAuthor` must re-render the SAME document, not re-parse the
  // original bytes — a reparse silently discarded every unsaved edit,
  // including pending suggestions. Only a new `source` parses fresh.
  const docCacheRef = useRef<{ source: DocxViewProps["source"]; doc: DocxDocument; history: EditHistory } | null>(null);
  useEffect(() => {
    const c = containerRef.current;
    if (!c) return;
    const onHf = (e: Event) => setHfMode(!!(e as CustomEvent<{ active: boolean }>).detail?.active);
    c.addEventListener("dxw-hfmode", onHf);
    return () => c.removeEventListener("dxw-hfmode", onHf);
  }, []);

  useEffect(() => {
    let cancelled = false;
    let handle: RenderHandle | null = null;
    let editor: DocxEditor | null = null;
    let detachPresenceSender: (() => void) | null = null;
    let onDeleteComment: ((id: string) => void) | undefined;
    let onReplyComment: ((id: string, text: string) => void) | undefined;
    let onResolveComment: ((id: string, resolved: boolean) => void) | undefined;
    let onEditComment: ((id: string, text: string) => void) | undefined;
    // Mutable current zoom for this document's lifetime: rerender() reads it and
    // applyZoomRef updates it in place, so a zoom change re-paints without
    // re-running this effect (which would reparse and drop editor/undo state).
    let curZoom = effZoomRef.current;
    // Kept as one object so applyZoom can update the editor's live zoom (the
    // editor reads host.zoom on every hit-test); reassigned in the editable path.
    let editorConfig: ConstructorParameters<typeof DocxEditor>[0] | null = null;
    setError(null);

    // One measurer for the document's lifetime: its width/metrics caches survive
    // across keystrokes so unchanged text is not re-measured on every relayout
    // (the default path builds a fresh, cold measurer per layoutDocument call).
    // Cache hits return the exact same values, so layout output is unchanged.
    const measurer: TextMeasurer = createMeasurer();
    // Previous layout result, fed back so the engine can reuse the pages of an
    // edit's unchanged prefix/suffix (incremental pagination). Same document, so
    // it stays valid across keystrokes; the engine falls back to a full layout
    // whenever it can't prove reuse is byte-identical.
    let prevLayout: LayoutResult | null = null;
    let paintedModelVersion = -1;
    let pages = 0;
    let layoutJob = 0;
    let layoutAbort: AbortController | null = null;
    let layoutTimer: ReturnType<typeof setTimeout> | null = null;
    let restoreEditorFocus = false;
    /** An async background layout is currently in flight. */
    let layoutRunning = false;
    /** Changes arrived while it ran (folded queue requests, or a sync paint
     * superseding a dying job) — the completion repair must re-validate. */
    let layoutDirty = false;
    /** Background-layout lifecycle counters, published when the host arms
     * `__dxwPerf` (benchmarks, the perf HUD, the e2e stress suite). The
     * inert-editor livelock was invisible precisely because nothing counted
     * jobs started against jobs that ever landed. */
    const countJob = (k: string): void => {
      const p = (globalThis as { __dxwPerf?: { jobs?: Record<string, number> } }).__dxwPerf;
      if (!p) return;
      p.jobs ??= {};
      p.jobs[k] = (p.jobs[k] ?? 0) + 1;
    };

    const setLayoutPending = (pending: boolean): void => {
      const container = containerRef.current;
      if (pending && container) {
        restoreEditorFocus ||= document.activeElement === container || container.contains(document.activeElement);
      }
      if (container) {
        container.inert = pending;
        container.toggleAttribute("data-dxw-layout-busy", pending);
        container.setAttribute("aria-busy", String(pending));
      }
      setLayoutBusy(pending);
      if (!pending && restoreEditorFocus && container) {
        restoreEditorFocus = false;
        container.focus({ preventScroll: true });
      }
    };

    const paintLayout = (doc: DocxDocument, layout: LayoutResult, layoutMs: number): number => {
      const perf = (globalThis as { __dxwPerf?: { last?: Record<string, number> } }).__dxwPerf;
      prevLayout = layout;
      paintedModelVersion = doc.modelVersion;
      const container = containerRef.current;
      if (!container) return 0;
      if (
        editable
        && !customElements.get("model-viewer")
        && (layout._hasModel3D ||
          layout.pages.some((page) => page.items.some((item) => item.kind === "image" && item.model3D)))
      ) {
        void import("@google/model-viewer");
      }
      // Re-rendering replaces the page DOM; keep the user's scroll position
      // (destroy-then-append clamps scrollTop to 0 otherwise). The previous
      // handle is handed to renderToDom so it can adopt the DOM of unchanged
      // pages and tear down only what actually changed.
      const tScroll0 = perf ? performance.now() : 0;
      const { scrollTop, scrollLeft } = container;
      const prev = handle;
      const t2 = perf ? performance.now() : 0;
      handle = renderToDom(doc, layout, container, {
        zoom: curZoom,
        interactive: editable,
        virtualize: true,
        comments: showComments,
        onDeleteComment,
        onReplyComment,
        onResolveComment,
        onEditComment,
        commentAuthor,
        onViewportChange: () => editor?.afterViewportChange(),
      }, prev ?? undefined);
      const tDom = perf ? performance.now() : 0;
      handleRef.current = handle;
      if (presenceRef.current) drawCollabPresence();
      container.scrollTop = scrollTop;
      container.scrollLeft = scrollLeft;
      const t3 = perf ? performance.now() : 0;
      editor?.afterRender();
      if (perf) {
        perf.last = {
          layout: layoutMs,
          destroy: 0,
          render: t3 - t2,
          scrollRead: t2 - tScroll0,
          renderDom: tDom - t2,
          scrollWrite: t3 - tDom,
          afterRender: performance.now() - t3,
          totalPages: layout.totalPages,
        };
        // The FIRST paint on this page — the mount paint — recorded once, so
        // a bench can read the mount's layout/render breakdown even after
        // later incremental paints have overwritten `last`.
        const p = perf as { mount?: Record<string, number> };
        p.mount ??= { ...perf.last };
      }
      if (pages !== layout.totalPages) onPageCountChangeRef.current?.(layout.totalPages);
      pages = layout.totalPages;
      return layout.totalPages;
    };

    const queueGlobalLayout = (doc: DocxDocument, delayMs = 0, preferHeadersOnly = false): void => {
      countJob("bgQueued");
      /**
       * FOLD INTO THE RUNNING JOB, never restart it. The old abort-and-restart
       * had no progress guarantee, and on a document whose full layout takes
       * seconds it was a LIVELOCK: every mid-flight change (a remote apply, a
       * coalesced repaint) restarted the layout from block zero, so under a
       * steady stream of broadcasts the layout never completed, the container
       * stayed `inert` forever, and the editor was dead while looking fine —
       * keystrokes fell through to the page (space scrolled). The running job
       * reads the LIVE tree, and the completion repair below re-validates
       * every block against the final tree, so folding loses nothing.
       */
      if (layoutRunning) {
        layoutDirty = true;
        countJob("bgFolded");
        return;
      }
      layoutAbort?.abort(); // a delayed timer job at most — a running one folds above
      if (layoutTimer) clearTimeout(layoutTimer);
      const abort = new AbortController();
      layoutAbort = abort;
      const job = ++layoutJob;
      const modelVersion = doc.modelVersion;
      const start = () => {
        layoutTimer = null;
        if (cancelled || job !== layoutJob || abort.signal.aborted) return;
        const started = performance.now();
        if (preferHeadersOnly && prevLayout) {
          const fast = relayoutHeadersFooters(doc, prevLayout, measurer, mergeRecordRef.current);
          if (fast) {
            if (doc.modelVersion === modelVersion) paintLayout(doc, fast, performance.now() - started);
            layoutAbort = null;
            setLayoutPending(false);
            return;
          }
        }
        countJob("bgStarted");
        const startModelVersion = doc.modelVersion;
        layoutRunning = true;
        layoutDirty = false;
        setLayoutPending(true);
        void layoutDocumentAsync(doc, {
          measurer,
          signal: abort.signal,
          windowModel: true,
          mergeRecord: mergeRecordRef.current,
        }).then((layout) => {
          if (cancelled || job !== layoutJob || abort.signal.aborted) return;
          if (doc.modelVersion !== startModelVersion || layoutDirty) {
            /**
             * The tree changed while the async layout ran (remote applies —
             * local input is gated by `inert`). Do NOT discard the result:
             * repair it synchronously instead. The engine's signature scan
             * relays exactly the blocks whose content no longer matches what
             * this run laid out (each run records per-block signatures), and
             * falls back to a full layout only when it cannot prove reuse —
             * either way the painted result matches the CURRENT tree, and the
             * job LANDS instead of being thrown away and restarted forever.
             */
            countJob("bgRepaired");
            layout = layoutDocument(doc, {
              measurer,
              windowModel: true,
              prev: layout,
              mergeRecord: mergeRecordRef.current,
            });
          }
          layoutDirty = false;
          paintLayout(doc, layout, performance.now() - started);
          countJob("bgCompleted");
        }).catch((cause: unknown) => {
          if (abort.signal.aborted || cancelled) return;
          const err = cause instanceof Error ? cause : new Error(String(cause));
          setError(err);
          onError?.(err);
        }).finally(() => {
          layoutRunning = false;
          if (job === layoutJob) {
            layoutAbort = null;
            setLayoutPending(false);
          } else if (layoutDirty && !cancelled) {
            // A queue request folded into this job while it was being torn
            // down (aborted by a sync paint, then a new request arrived before
            // this cleanup ran). It has no owner now — re-queue it, or the
            // change it represents never paints.
            countJob("bgRequeued");
            queueGlobalLayout(doc, 0, false);
          }
        });
      };
      if (delayMs > 0) {
        setLayoutPending(false);
        layoutTimer = setTimeout(start, delayMs);
      } else {
        start();
      }
    };

    // Collaboration: emit a block-addressed intent (formatParagraph/setListType)
    // for each distinct paragraph the given text targets belong to.
    const emitBlockIntents = (targets: XmlElement[], make: (blockId: number) => EditorIntent): void => {
      const d = docCacheRef.current?.doc;
      const current = collabRef.current;
      if (!current || !d?.stableIds) return;
      const seen = new Set<number>();
      for (const t of targets) {
        const p = paragraphOf(d, t);
        const blockId = p ? d.stableIds.idOf(p) : undefined;
        if (blockId !== undefined && !seen.has(blockId)) {
          seen.add(blockId);
          current.submit(make(blockId));
        }
      }
    };

    // Collaboration: emit ONE intent covering every distinct paragraph the
    // targets belong to (in selection order). setListType uses this instead of
    // emitBlockIntents because its apply MINTS a numbering definition: one
    // intent per paragraph mints one per apply, so replicas diverge from the
    // originator's single shared-definition mutation (see toggleList).
    const emitListIntent = (targets: XmlElement[], make: (blockIds: number[]) => EditorIntent): void => {
      const d = docCacheRef.current?.doc;
      const current = collabRef.current;
      if (!current || !d?.stableIds) return;
      const blockIds: number[] = [];
      for (const t of targets) {
        const p = paragraphOf(d, t);
        const blockId = p ? d.stableIds.idOf(p) : undefined;
        if (blockId !== undefined && !blockIds.includes(blockId)) blockIds.push(blockId);
      }
      if (blockIds.length > 0) current.submit(make(blockIds));
    };

    // Draw remote participants' carets over the current render (collab mode).
    const drawCollabPresence = (): void => {
      const handle = handleRef.current;
      const d = docCacheRef.current?.doc;
      const presence = presenceRef.current;
      if (!handle || !d || !presence) return;
      // Presence carets are drawn INTO each page surface (see drawPresenceCarets),
      // so their surface-local geometry aligns and scales like the local caret.
      // Highlights first, carets on top (each draw clears only its own class).
      drawPresenceSelections(handle.root, computePresenceSelections(handle, d, presence));
      drawPresenceCarets(handle.root, computePresenceCarets(handle, d, presence), collabRef.current?.participantNames);
    };
    redrawPresenceRef.current = drawCollabPresence;

    const rerender = (
      doc: DocxDocument,
      dirtyBlock?: XmlElement,
      scope: "local" | "global" | "background" = "local",
      dirtySource?: XmlElement,
    ): number => {
      const modelChanged = paintedModelVersion !== doc.modelVersion;
      const headerFooterOnly = scope === "global" && !modelChanged;
      const globalChange = scope !== "local" || modelChanged;
      const queue =
        editable &&
        prevLayout !== null &&
        (scope === "background" ||
          (prevLayout.totalPages > BACKGROUND_LAYOUT_PAGE_THRESHOLD && globalChange));
      if (queue) {
        queueGlobalLayout(doc, headerFooterOnly ? 120 : 0, headerFooterOnly);
        return prevLayout!.totalPages;
      }

      layoutAbort?.abort();
      layoutAbort = null;
      if (layoutTimer) clearTimeout(layoutTimer);
      layoutTimer = null;
      layoutJob++;
      // This synchronous layout reads the current tree, so it also covers any
      // change that folded into the (just-aborted) background job; its dying
      // `finally` must not re-queue a background layout for work painted here.
      layoutDirty = false;
      setLayoutPending(false);
      const started = performance.now();
      const layout =
        headerFooterOnly && prevLayout
          ? relayoutHeadersFooters(doc, prevLayout, measurer, mergeRecordRef.current) ??
            layoutDocument(doc, { measurer, windowModel: true, mergeRecord: mergeRecordRef.current })
          : layoutDocument(doc, {
              measurer,
              windowModel: true,
              mergeRecord: mergeRecordRef.current,
              prev: globalChange ? undefined : prevLayout ?? undefined,
              dirtyHint: globalChange ? undefined : dirtyBlock,
              dirtySource: globalChange ? undefined : dirtySource,
            });
      return paintLayout(doc, layout, performance.now() - started);
    };

    rerenderRef.current = null;

    (async () => {
      // Live-collab path: render the replica's own document object directly.
      const liveDoc = collab?.doc ?? null;
      const cached = liveDoc
        ? (docCacheRef.current?.doc === liveDoc ? docCacheRef.current : null)
        : (docCacheRef.current?.source === source ? docCacheRef.current : null);
      const bytes = cached || liveDoc ? null : await toBytes(source);
      if (cancelled) return;
      if (!cached && typeof document !== "undefined" && document.fonts?.ready) {
        try {
          // Canvas measurement doesn't trigger webfont loads; request the
          // metric-compatible substitutes explicitly if the host provides them.
          // Real Office faces (Cambria Math, real Calibri/Times/Arial, the CJK
          // families) are registered dev-only via @font-face over /fonts-local/;
          // load() 404s fast (and .catch swallows it) when they're absent, so
          // machines without the fonts fall back to the substitutes seamlessly.
          const loads: Promise<unknown>[] = [];
          // Latin faces measured on canvas (widths must be real before layout).
          const latin = [
            "Carlito", "Caladea", "Cambria", "Times New Roman", "Arial",
            "Calibri", "Calibri Light", "Tahoma", "Franklin Gothic Medium",
          ];
          for (const fam of latin) {
            for (const variant of ["", "italic ", "bold ", "bold italic "]) {
              loads.push(document.fonts.load(`${variant}16px "${fam}"`).catch(() => []));
            }
          }
          const complex = [
            ["Mangal", "अ"],
            ["Latha", "அ"],
            ["Noto Sans Lao Looped", "ກ"],
          ] as const;
          for (const [fam, sample] of complex) {
            for (const variant of ["", "bold "]) {
              loads.push(document.fonts.load(`${variant}16px "${fam}"`, sample).catch(() => []));
            }
          }
          loads.push(document.fonts.load('16px "Cambria Math"').catch(() => []));
          loads.push(document.fonts.load('16px "Segoe UI Emoji"', "🚀").catch(() => []));
          // CJK faces only affect PAINT (widths are em-based, line pitch comes
          // from a metrics table), so they don't gate layout — but load them so
          // the screenshot/paint uses the real glyphs when available.
          const cjk = [
            "MS Mincho", "MS Gothic", "Meiryo", "Yu Gothic", "Yu Mincho",
            "SimSun", "SimHei", "Microsoft JhengHei", "Microsoft YaHei",
            "Malgun Gothic",
          ];
          for (const fam of cjk) {
            for (const variant of ["", "bold "]) {
              loads.push(document.fonts.load(`${variant}16px "${fam}"`, "A漢").catch(() => []));
            }
          }
          await Promise.all(loads);
          await document.fonts.ready;
        } catch {
          /* non-fatal */
        }
      }
      if (cancelled) return;
      const doc = liveDoc ?? (cached ? cached.doc : DocxDocument.load(bytes!));
      if (!cached) docCacheRef.current = { source, doc, history: new EditHistory(doc) };
      // Expose an in-place repaint for the live-collab renderSignal effect.
      //
      // A NARROW scope (a remote text edit's one paragraph, or a split's two
      // halves) repaints through the same incremental path local typing takes:
      // prev layout retained, the paragraph as the dirty hint, synchronous —
      // so it never trips the input-blocking background layout, and watching
      // a collaborator type on page 100 of a 500-page document costs one
      // paragraph, not the document. The scoped resync left modelVersion
      // unchanged, which is exactly what rerender's "local" path needs; if a
      // resync FELL BACK to a full refresh despite the narrow scope, the
      // modelVersion bump makes rerender treat it as a global change — the
      // scope can never leave the view stale.
      //
      // DOC scope (structural/unverifiable intents, reloads, media installs)
      // keeps the old behavior: force modelChanged (the collab apply may not
      // bump modelVersion) so the body — not just headers/footers — is relaid
      // out against the mutated tree, async past the page threshold.
      rerenderRef.current = (d, scope) => {
        if (!scope || scope.kind === "doc") {
          paintedModelVersion = -1;
          rerender(d, undefined, "global");
          return;
        }
        // Split → hint the FINAL new paragraph (the engine's insertion fast
        // path re-hashes the split pair + neighbours and reflows forward,
        // re-attaching the shifted page suffix). Single block → in-place
        // hint. Several blocks → no hint: the engine's full signature scan
        // still relays only from the first dirty block.
        const hint =
          scope.kind === "split" ? scope.after : scope.blocks.length === 1 ? scope.blocks[0] : undefined;
        rerender(d, hint, "local");
      };
      // Collaboration: populate the stable-id side table so the editor can
      // encode intent positions and emit them (no-op / zero-cost otherwise).
      if (collab) doc.enableStableIds();
      const wantView = revisions === "markup" ? "markup" : "final";
      if (doc.revisionView !== wantView) doc.setRevisionView(wantView);
      // Feed the real page width into fit-to-width now that we know it, then
      // recompute (the first ResizeObserver pass ran against the 816px default).
      basePageWidthRef.current = doc.sections[0]?.props.pageWidth ?? 816;
      curZoom = effZoomRef.current;
      const pageCount = rerender(doc);
      // That paint covered the whole document, so any dirty scope accumulated
      // before it (the welcome's doc replacement, the seed tail's applies) is
      // already on screen — drain it, or the FIRST remote keystroke would
      // repaint globally for changes this paint already showed.
      collab?.takeRenderScope?.();
      pages = pageCount;
      recomputeFit();
      onLoad?.({ pageCount, document: doc });
      if (onMissingFonts && prevLayout) {
        const missing = detectMissingFonts(prevLayout);
        if (!cancelled) onMissingFonts(missing);
      }

      if (editable && containerRef.current) {
        // Undo history survives mode switches with the cached document.
        const history = docCacheRef.current?.doc === doc ? docCacheRef.current.history : new EditHistory(doc);

        // ---- Outbound presence: caret + selection highlight, one payload ----
        // Two triggers feed it: the editor's onCaretMove (deduped upstream, so
        // it fires only when the caret actually moved) and the document-level
        // "dxw-selection" event (fires for every selection change, including
        // the ones that leave the caret put, e.g. select-all). Both land in the
        // SAME task — notifySelection dispatches the event and then reports the
        // caret — so the send is coalesced onto a microtask: remote tabs get one
        // payload per settled state, never a caret-without-its-highlight flash.
        const MAX_SENT_RANGES = 64;
        let presenceCaret: { blockId: number; runId: number; offset: number } | null = null;
        // Seeded with the empty payload's key so a mount (or another mounted
        // view's dxw-selection — the event is document-level) never emits a
        // redundant "I have no cursor" message.
        let presenceKey: string | null = "null";
        let presenceQueued = false;

        /** The local selection as wire ranges: one per selection segment, each
         * inside a single w:t (and therefore a single run) by construction.
         * `encodeCaret` converts the segment START to the run's wire basis
         * (cumulative, separators counted), and the segment's own length —
         * which is w:t-local and separator-free — extends it to the end. */
        const selectionRanges = (): { blockId: number; runId: number; start: number; end: number }[] | undefined => {
          const ids = doc.stableIds;
          if (!ids || !editor) return undefined;
          const segs = editor.getSelectionSegments();
          if (segs.length === 0) return undefined;
          const out: { blockId: number; runId: number; start: number; end: number }[] = [];
          for (const seg of segs) {
            if (out.length >= MAX_SENT_RANGES) break;
            if (!seg.t || seg.end <= seg.start) continue;
            const enc = ids.encodeCaret(seg.t, seg.start, (el) => doc.findParentOf(el) ?? null);
            if (!enc) continue; // not id-tracked (math internals): not addressable
            out.push({ blockId: enc.blockId, runId: enc.runId, start: enc.offset, end: enc.offset + (seg.end - seg.start) });
          }
          return out.length > 0 ? out : undefined;
        };

        const flushPresence = (): void => {
          presenceQueued = false;
          const current = collabRef.current;
          if (!current?.setPresence) return;
          const ranges = selectionRanges();
          // A drag-selection can leave the editor without a caret; the payload
          // still needs an anchor (the pre-ranges shape old clients read), so
          // fall back to the selection's end — where the focus visually is.
          const last = ranges?.[ranges.length - 1];
          const anchor =
            presenceCaret ?? (last ? { blockId: last.blockId, runId: last.runId, offset: last.end } : null);
          const pos: PresencePosition | null = anchor ? (ranges ? { anchor, ranges } : { anchor }) : null;
          const key = pos ? JSON.stringify(pos) : "null";
          if (key === presenceKey) return; // dedup, like the caret path
          presenceKey = key;
          current.setPresence(pos);
        };
        const queuePresence = (): void => {
          if (presenceQueued) return;
          presenceQueued = true;
          queueMicrotask(flushPresence);
        };
        if (collab?.setPresence && typeof document !== "undefined") {
          const onSelectionEvent = () => queuePresence();
          document.addEventListener("dxw-selection", onSelectionEvent);
          detachPresenceSender = () => document.removeEventListener("dxw-selection", onSelectionEvent);
        }

        editorConfig = {
          doc,
          container: containerRef.current,
          getHandle: () => handle,
          rerender: (dirtyBlock?: XmlElement, scope?: "local" | "global" | "background", dirtySource?: XmlElement) => {
            pages = rerender(doc, dirtyBlock, scope, dirtySource);
          },
          zoom: curZoom,
          history,
          onFormatShortcut: (kind) => {
            const segs = editor?.getSelectionSegments() ?? [];
            if (segs.length === 0) return;
            const selectedAll = editor?.isEntireDocumentSelected() ?? false;
            const fmt = summarizeSelection(segs);
            const patch =
              kind === "bold" ? { bold: !fmt?.bold } :
              kind === "italic" ? { italic: !fmt?.italic } :
              { underline: !fmt?.underline };
            history.checkpoint();
            // Capture ids BEFORE the mutation (a sub-range format splits the
            // run and prunes the original id), then emit via the shared
            // helper so the shortcut and the toolbar behave identically.
            const preIds = captureFormatIds(segs);
            // One draw per gesture: the same author+date go into the local
            // w:rPrChange and into the intent every replica applies.
            const suggest = editor?.suggestionMeta();
            const formatted = applyRunFormat(doc, segs, patch, suggestMeta(doc, suggest));
            emitFormatIntents(segs, preIds, formatted, patch as Record<string, unknown>, suggest);
            pages = rerender(doc);
            if (selectedAll) editor?.selectAll();
            else if (formatted.length > 0) editor?.selectRanges(formatted);
            document.dispatchEvent(new CustomEvent("dxw-selection"));
          },
          // Through the public API, exactly like the style gallery button.
          // A private copy of the mutation used to run here instead, and it
          // emitted NO intent — ⌘⌥1 changed the heading for the person who
          // pressed it and nobody else.
          onStyleShortcut: (styleId) => {
            apiRef.current?.setParagraphStyle(styleId);
            document.dispatchEvent(new CustomEvent("dxw-selection"));
          },
          // Collaboration: forward each local edit as an intent. The editor
          // emits only when doc.stableIds is populated, so enable it here.
          onIntent: collab ? (intent) => collabRef.current?.submit(intent) : undefined,
          // Cmd+Z in a room reverses the last SEQUENCED action over the wire;
          // without this hook the editor declines rather than replaying local
          // history (which would fork the room silently).
          onCollabUndo: collab?.undoLast ? () => collabRef.current?.undoLast?.() : undefined,
          // A single-process session (LocalDocumentSession) instead replays
          // the local history stack; this tells the editor so and lets the
          // session bump its revision after each applied undo/redo.
          onLocalHistory: collab?.noteLocalHistory ? () => collabRef.current?.noteLocalHistory?.() : undefined,
          // Presence: broadcast the local caret so remote participants can
          // draw this user's cursor (the anchor mirrors the intent addressing).
          // The caret is recorded here and sent by flushPresence together with
          // the current selection ranges — one payload, one highlight+cursor.
          onCaretMove: collab?.setPresence
            ? (pos) => {
                presenceCaret = pos;
                queuePresence();
              }
            : undefined,
          // Disjoint carried-id blocks for the editor's node-creating intents
          // (paragraph split); see EditorHost.allocIds.
          allocIds: collab?.allocIds ? (n) => collabRef.current?.allocIds?.(n) ?? [] : undefined,
          onTextCommand: (command) => {
            const current = apiRef.current;
            if (!current) return;
            if (command === "link") {
              const anchor = containerRef.current;
              if (!anchor) return;
              void requestTextInputDialog(anchor, {
                title: "Link address",
                label: "URL",
                value: current.getLinkAt() ?? "https://",
                inputType: "url",
              }).then((next) => {
                if (next !== null) current.setLink(next.trim() || null);
              });
            } else if (command === "comment") {
              const anchor = containerRef.current;
              if (!anchor) return;
              void requestTextInputDialog(anchor, {
                title: "New comment",
                label: "Comment",
                submitLabel: "Comment",
                multiline: true,
                mentions: commentMentionsRef.current,
              }).then((text) => {
                if (text?.trim()) current.addComment(text.trim());
              });
            } else if (command === "goToPage") {
              const anchor = containerRef.current;
              if (!anchor) return;
              void requestTextInputDialog(anchor, {
                title: "Go to page",
                label: `Page number (1–${current.pageCount()})`,
                value: "",
              }).then((text) => {
                const page = Number(text?.trim());
                if (Number.isInteger(page) && page > 0) current.goToPage(page);
              });
            } else {
              runHostCommand(current, command, painterRef);
            }
          },
        };
        editor = new DocxEditor(editorConfig);
        editor.attach();
        onDeleteComment = (id) => {
          if (collabDocOp(() => ({ kind: "deleteComment", commentId: id }))) return;
          history.checkpoint();
          if (deleteComment(doc, id)) pages = rerender(doc);
        };
        onResolveComment = (id, resolved) => {
          if (collabDocOp(() => ({ kind: "resolveComment", commentId: id, resolved, paraId: hex8() }))) return;
          history.checkpoint();
          if (setCommentResolved(doc, id, resolved)) pages = rerender(doc);
        };
        onEditComment = (id, text) => {
          if (collabDocOp(() => ({ kind: "editComment", commentId: id, text }))) return;
          history.checkpoint();
          if (editCommentText(doc, id, text)) pages = rerender(doc);
        };
        onReplyComment = (id, text) => {
          if (collabDocOp(() => ({
            kind: "replyComment", parentId: id, text, author: commentAuthor,
            initials: commentAuthor.split(/\s+/).map((part) => part[0] ?? "").join("").slice(0, 2).toUpperCase() || undefined,
            date: new Date().toISOString(), paraIds: [hex8()],
          }))) return;
          history.checkpoint();
          const initials = commentAuthor
            .split(/\s+/)
            .map((part) => part[0] ?? "")
            .join("")
            .slice(0, 2)
            .toUpperCase();
          if (replyToComment(doc, id, text, commentAuthor, initials || undefined)) {
            pages = rerender(doc);
          }
        };
        pages = rerender(doc); // re-render with the delete affordance wired
        let findState: { matches: ReturnType<typeof findAll>; index: number } = { matches: [], index: 0 };
        // Cross-reference dialog targets (listCrossRefTargets), keyed by index.
        let crossRefTargets: CrossRefTarget[] = [];
        // Review-tab comment navigation cursor (-1 = not started).
        let commentNav = -1;
        const selectMatch = (i: number) => {
          const m = findState.matches[i];
          if (!m || !editor) return;
          const restore = handle?._virtualized ? handle.materializeAll?.() : undefined;
          editor.selectRanges(m.ranges);
          // Bring the hit into view.
          const t = m.ranges[0]?.t;
          const el = handle?.bindingsByText.get(t)?.[0]?.el;
          el?.scrollIntoView({ block: "center", behavior: "instant" as ScrollBehavior });
          restore?.();
          handle?.updateViewport?.();
          editor.selectRanges(m.ranges);
        };
        const insertionTarget = () => {
          const caret = editor?.getCaretTarget();
          if (caret) return caret;
          const last = [...(editor?.getSelectionSegments() ?? [])].reverse().find((segment) => segment.t);
          return last?.t ? { t: last.t, offset: last.end } : null;
        };
        /** The very start of the document's text, as a last-resort insertion
         * position. Reaching for the toolbar BEFORE clicking into the page is
         * the obvious first move, and every insert command used to answer it
         * with silence — the shape of the "images don't work at all" report.
         * Only a document with no text at all is genuinely unaddressable. */
        const documentStart = (): { t: XmlElement; offset: number } | null => {
          for (const section of doc.sections) {
            for (const block of section.blocks) {
              if (block.type !== "paragraph") continue;
              for (const item of block.children) {
                for (const run of item.type === "run" ? [item] : item.runs) {
                  const text = run.content.find((c) => c.kind === "text");
                  if (text?.srcT) return { t: text.srcT, offset: text.srcT.text.length };
                }
              }
            }
          }
          return null;
        };
        // ---- Collaborative op routing (toolbar/API commands) ----
        // In collab mode these commands are routed through the CANONICAL
        // intent apply (collab.submitOp → the same applyIntent code every
        // replica and the server run) instead of the local core mutation:
        // the optimistic local result is byte-identical everywhere by
        // construction, with no per-command convergence analysis. The repaint
        // arrives via renderSignal (submit → onChange → coalesced repaint).
        const collabAnchor = (t: XmlElement, offset: number) =>
          doc.stableIds ? doc.stableIds.encodeCaret(t, offset, (el) => doc.findParentOf(el) ?? null) : null;
        const collabOp = (
          make: (
            anchor: { blockId: number; runId: number; offset: number },
            alloc: (n: number) => number[],
          ) => ({ kind: string } & Record<string, unknown>) | null,
          at?: { t: XmlElement; offset: number } | null,
        ): boolean => {
          const current = collabRef.current;
          if (!current?.submitOp || !doc.stableIds) return false;
          // Past this gate we are IN collab mode and must return true no
          // matter what: returning false would drop the caller into its
          // LOCAL mutation fallback, which never rides the wire — a silent
          // permanent divergence that even poisons the persisted bundle
          // (found live: an unaddressable insert-shape landed locally,
          // desynced the room, and survived reload). An unaddressable
          // command in collab mode is an honest no-op instead.
          const target = at !== undefined ? at : insertionTarget();
          if (!target) return true;
          const anchor = collabAnchor(target.t, target.offset);
          if (!anchor) return true;
          const intent = make(anchor, (n) => current.allocIds?.(n) ?? []);
          if (!intent) return true;
          history.checkpoint();
          current.submitOp(intent);
          return true;
        };
        /** Submit a REGISTERED run-addressed operation. The wire field the
         * address goes in, and the number of carried ids the mutation needs,
         * both come from the registry declaration — so this call site restates
         * neither, and it cannot drift from the agent compiler, which sizes
         * its allocation from that same declaration. */
        const collabRunOperation = <K extends RegisteredOperationKind>(
          kind: K,
          args: RegisteredOperationArgs<K>,
        ): boolean =>
          collabOp((anchor, alloc) => operationBody(kind, anchor.runId, args, alloc) as never);
        /** The w:tbl the caret is in, or null when it is outside a table. */
        const caretTable = (): XmlElement | null => {
          const caret = editor?.getCaretTarget();
          if (!caret) return null;
          return cellContextOf(doc, caret.t)?.tbl ?? null;
        };
        /** The column widths a table was last PAINTED at, read off its resize
         * grips. An autofit table's grid rarely agrees with them, which is
         * exactly why freezing to fixed widths needs these and not the grid. */
        const renderedColumnWidths = (tblEl: XmlElement): number[] | undefined => {
          for (const page of prevLayout?.pages ?? []) {
            for (const item of page.items) {
              if (item.kind === "grip" && item.axis === "col" && item.tbl === tblEl && item.renderedWidths) {
                return item.renderedWidths;
              }
            }
          }
          return undefined;
        };
        /** Submit a REGISTERED cell-addressed operation, falling back to the
         * local mutation outside a room. The caret's paragraph is the address;
         * the registry decides whether the operation reads the cell or widens
         * to the table. */
        const runTableOperation = <K extends RegisteredOperationKind>(
          kind: K,
          args: RegisteredOperationArgs<K>,
          local: (caretT: XmlElement, meta: RevisionMeta | undefined) => boolean,
        ): void => {
          const caret = editor?.getCaretTarget();
          if (!caret) return;
          // One draw per gesture (doc 05 rule a): the same author+date go into
          // the local *PrChange and into the intent every replica applies.
          // Every kind routed here declares an optional suggest payload.
          const suggest = editor?.suggestionMeta();
          const payload = (suggest ? { ...args, suggest } : args) as RegisteredOperationArgs<K>;
          if (
            collabOp(
              (anchor, alloc) => operationBody(kind, anchor.blockId, payload, alloc) as never,
              { t: caret.t, offset: 0 },
            )
          ) {
            return;
          }
          history.checkpoint();
          if (local(caret.t, suggestMeta(doc, suggest))) pages = rerender(doc);
        };
        /** Document-level ops (page layout, line numbering, cover page). */
        const collabDocOp = (
          make: (alloc: (n: number) => number[]) => ({ kind: string } & Record<string, unknown>) | null,
        ): boolean => {
          const current = collabRef.current;
          if (!current?.submitOp || !doc.stableIds) return false;
          // In collab mode never fall through to the local path (see collabOp).
          const intent = make((n) => current.allocIds?.(n) ?? []);
          if (!intent) return true;
          history.checkpoint();
          current.submitOp(intent);
          return true;
        };
        /** Per-block ops over the current selection/caret paragraphs. */
        const collabBlockOp = (
          targets: XmlElement[],
          make: (blockId: number) => ({ kind: string } & Record<string, unknown>) | null,
        ): boolean => {
          const current = collabRef.current;
          if (!current?.submitOp || !doc.stableIds) return false;
          // In collab mode never fall through to the local path (see collabOp):
          // even with zero addressable targets this returns true (honest no-op).
          const submitted = new Set<number>();
          let any = false;
          for (const t of targets) {
            const a = collabAnchor(t, 0);
            if (!a || submitted.has(a.blockId)) continue;
            submitted.add(a.blockId);
            const intent = make(a.blockId);
            if (!intent) continue;
            if (!any) history.checkpoint();
            any = true;
            current.submitOp(intent);
          }
          return true;
        };
        /** Capture per-seg stable ids BEFORE a run format (a sub-range format
         * splits the run — the original id is pruned by the mutation). Wire
         * offsets are in the run's WIRE space (cumulative, separators counted
         * — checkpoint B1 rev 2): capture the segment's start in that basis
         * plus the run's total wire length. */
        const captureFormatIds = (segs: SelectionSegment[]) =>
          collabRef.current && doc.stableIds
            ? segs.map((seg) => ({
                runId: seg.run.src ? doc.stableIds!.idOf(seg.run.src) : undefined,
                blockId: seg.run.srcParent ? doc.stableIds!.idOf(seg.run.srcParent) : undefined,
                runLen: seg.run.src ? runWireLength(seg.run.src) : 0,
                cumStart: seg.run.src && seg.t ? (wireOffsetOf(seg.run.src, seg.t, seg.start) ?? seg.start) : seg.start,
              }))
            : [];
        /** Emit formatRun/formatRange intents for an applied run format and
         * re-key the local split pieces to the carried ids (the same walk the
         * server's apply performs). Shared by the Ctrl+B/I/U shortcut and the
         * toolbar's applyFormat. */
        const emitFormatIntents = (
          segs: SelectionSegment[],
          preIds: { runId?: number; blockId?: number; runLen: number; cumStart: number }[],
          formatted: { t: XmlElement | null }[],
          patch: Record<string, unknown>,
          suggest?: { author: string; date: string },
        ): void => {
          const current = collabRef.current;
          if (!current || !doc.stableIds) return;
          // A selection over a run with several w:t yields one seg PER w:t,
          // all sharing the runId. Group them: the whole-run check must use
          // the COMBINED cumulative span (each seg alone is partial), else a
          // plain Ctrl+B over such a run would emit a partial formatRange and
          // silently drop the later w:t's formatting on the wire.
          const spanByRun = new Map<number, { lo: number; hi: number }>();
          segs.forEach((seg, i) => {
            const pre = preIds[i];
            if (!pre || pre.runId === undefined) return;
            const lo = pre.cumStart;
            const hi = pre.cumStart + (seg.end - seg.start);
            const cur = spanByRun.get(pre.runId);
            spanByRun.set(pre.runId, cur ? { lo: Math.min(cur.lo, lo), hi: Math.max(cur.hi, hi) } : { lo, hi });
          });
          const segCountByRun = new Map<number, number>();
          for (const pre of preIds) {
            if (pre?.runId !== undefined) segCountByRun.set(pre.runId, (segCountByRun.get(pre.runId) ?? 0) + 1);
          }
          const seenRuns = new Set<number>();
          segs.forEach((seg, i) => {
            const { runId, blockId, runLen, cumStart } = preIds[i] ?? { runLen: 0, cumStart: 0 };
            if (runId === undefined || blockId === undefined || seenRuns.has(runId)) return;
            seenRuns.add(runId);
            const span = spanByRun.get(runId) ?? { lo: cumStart, hi: cumStart + (seg.end - seg.start) };
            const cumEnd = cumStart + (seg.end - seg.start);
            // Whole-run when the segs cover the run — AND when the selection
            // spans SEVERAL w:t of one run: applyRunFormat's local fallback
            // (tTargets.size !== 1) formats the whole run IN PLACE with no
            // split, so a partial formatRange here would diverge (remote
            // splits, local doesn't — and the reassign below would stamp the
            // UNSPLIT run + an unrelated sibling; review bug 2). formatRun
            // matches the local mutation exactly.
            const whole = !seg.t || (span.lo <= 0 && span.hi >= runLen) || (segCountByRun.get(runId) ?? 1) > 1;
            if (whole) {
              current.submit({ kind: "formatRun", blockId, runId, patch, ...(suggest ? { suggest } : {}) } as never);
              return;
            }
            const before = cumStart > 0;
            const after = cumEnd < runLen;
            const alloc = current.allocIds?.((before ? 1 : 0) + 1 + (after ? 1 : 0)) ?? [];
            let k = 0;
            const beforeId = before ? alloc[k++] : undefined;
            const middleId = alloc[k++];
            const afterId = after ? alloc[k++] : undefined;
            if (middleId === undefined) return;
            current.submit({ kind: "formatRange", blockId, runId, start: cumStart, end: cumEnd, patch, beforeId, middleId, afterId, ...(suggest ? { suggest } : {}) } as never);
            const middleT = formatted[i]?.t;
            const middleRun = middleT ? doc.findParentOf(middleT) : null;
            const parent = middleRun ? doc.findParentOf(middleRun) : null;
            if (middleRun && parent && doc.stableIds) {
              // Guard the neighbors like the server's apply does (localRun):
              // with mIdx 0 the preceding sibling can be w:pPr — stamping it
              // with a run id retires an unrelated element's identity.
              const isRun = (el: XmlElement | undefined) => !!el && (el.name === "w:r" || el.name.endsWith(":r") || el.name === "r");
              const mIdx = parent.children.indexOf(middleRun);
              doc.stableIds.reassign(middleRun, middleId);
              if (before && beforeId !== undefined && isRun(parent.children[mIdx - 1])) {
                doc.stableIds.reassign(parent.children[mIdx - 1], beforeId);
              }
              if (after && afterId !== undefined && isRun(parent.children[mIdx + 1])) {
                doc.stableIds.reassign(parent.children[mIdx + 1], afterId);
              }
            }
          });
        };
        /** Word-style 8-hex provenance token (comment paraId). */
        const hex8 = () => {
          const b = new Uint8Array(4);
          crypto.getRandomValues(b);
          return [...b].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
        };
        const api: DocxViewApi = {
          document: doc,
          pageCount: () => pages,
          getSelectionFormat: () => {
            const segs = editor?.getSelectionSegments() ?? [];
            return summarizeSelection(segs.length > 0 ? segs : handle ? selectionToSegments(handle.bindings) : []);
          },
          applyFormat: (patch) => {
            if (!handle) return;
            const own = editor?.getSelectionSegments() ?? [];
            const segments = own.length > 0 ? own : selectionToSegments(handle.bindings);
            if (segments.length === 0) return;
            const selectedAll = editor?.isEntireDocumentSelected() ?? false;
            history.checkpoint();
            const preIds = captureFormatIds(segments);
            const suggest = editor?.suggestionMeta();
            const formatted = applyRunFormat(doc, segments, patch, suggestMeta(doc, suggest));
            emitFormatIntents(segments, preIds, formatted, patch as Record<string, unknown>, suggest);
            pages = rerender(doc);
            // Keep the formatted text selected so toolbar actions compose.
            if (selectedAll) editor?.selectAll();
            else if (formatted.length > 0) editor?.selectRanges(formatted);
          },
          addFootnote: (text) => {
            if (collabOp((a, ids) => ({ kind: "insertFootnote", runId: a.runId, text, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (insertFootnote(doc, target.t, target.offset, text) !== null) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
          addEndnote: (text) => {
            if (collabRunOperation("insertEndnote", { text })) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (insertEndnote(doc, target.t, target.offset, text) !== null) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
          insertPageNumber: (kind = "page") => {
            if (collabOp((a, ids) => ({ kind: "insertPageField", runId: a.runId, fieldKind: kind, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (insertPageField(doc, target.t, target.offset, kind)) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
          insertField: (instruction, cachedResult) => {
            if (collabOp((a, ids) => ({ kind: "insertField", runId: a.runId, instruction, cachedResult, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertField(doc, target.t, target.offset, instruction, cachedResult)) return false;
            pages = rerender(doc);
            return true;
          },
          insertDateTime: (kind, picture) => {
            const fmt = picture ?? (kind === "time" ? "h:mm am/pm" : "M/d/yyyy");
            if (collabOp((a, ids) => ({ kind: "insertDateTimeField", runId: a.runId, dtKind: kind, picture: fmt, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            const format = fmt;
            if (!insertDateTimeField(doc, target.t, target.offset, kind, format)) return false;
            pages = rerender(doc);
            return true;
          },
          listBookmarks: () => listBookmarks(doc),
          addBookmark: (name) => {
            {
              // Collab: a selection becomes a range bookmark on its first
              // segment; a bare caret becomes a point bookmark.
              const segs0 = editor?.getSelectionSegments() ?? [];
              const rangeSeg = segs0.find((sg) => sg.t && sg.end > sg.start);
              if (rangeSeg?.t) {
                // a.offset is the segment start in the wire basis (cumulative
                // within the run); the range keeps its local length.
                if (collabOp(
                  (a) => ({ kind: "insertBookmarkRange", runId: a.runId, name, start: a.offset, end: a.offset + (rangeSeg.end - rangeSeg.start) }),
                  { t: rangeSeg.t, offset: rangeSeg.start },
                )) return true;
              } else if (collabOp((a) => ({ kind: "insertBookmark", runId: a.runId, name }))) return true;
            }
            const segments = editor?.getSelectionSegments() ?? [];
            const target = editor?.getCaretTarget();
            history.checkpoint();
            const done = segments.length > 0
              ? insertBookmarkAroundSelection(doc, segments, name)
              : target
                ? insertBookmarkAt(doc, target.t, target.offset, name)
                : false;
            if (done) pages = rerender(doc);
            return done;
          },
          insertCrossReference: (bookmark, kind) => {
            if (collabOp((a, ids) => ({ kind: "insertCrossRef", runId: a.runId, bookmark, refKind: kind, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertCrossReference(doc, target.t, target.offset, bookmark, kind)) return false;
            pages = rerender(doc);
            return true;
          },
          listCrossRefTargets: () => {
            crossRefTargets = listCrossRefTargets(doc);
            return crossRefTargets.map((target, i) => ({
              key: String(i),
              kind: target.kind,
              text: target.text.length > 80 ? `${target.text.slice(0, 79)}…` : target.text,
            }));
          },
          insertCrossRefToTarget: (key, kind) => {
            const target = crossRefTargets[Number(key)];
            if (!target) return false;
            let name = target.bookmark;
            const current = collabRef.current;
            if (current?.submitOp && doc.stableIds) {
              // The hidden bookmark rides its own registered intent first;
              // the REF then names it like any user bookmark. Both carry the
              // originator's values, so every replica writes the same XML.
              if (!name) {
                name = nextRefBookmarkName(doc);
                const blockId = doc.stableIds.idOf(target.paragraph);
                if (blockId === undefined) return true; // honest no-op (see collabOp)
                history.checkpoint();
                current.submitOp(operationBody("ensureRefBookmark", blockId, { name }) as never);
              }
              const bookmark = name;
              return collabOp((a, ids) => ({ kind: "insertCrossRef", runId: a.runId, bookmark, refKind: kind, nodeIds: ids(8) }));
            }
            const at = insertionTarget();
            if (!at) return false;
            history.checkpoint();
            if (!name) {
              name = nextRefBookmarkName(doc);
              if (!ensureRefBookmark(doc, target.paragraph, name)) return false;
            }
            if (!insertCrossReference(doc, at.t, at.offset, name, kind)) return false;
            pages = rerender(doc);
            return true;
          },
          insertCaption: (label, text = "", position = "below") => {
            // The caption anchors at the SELECTED OBJECT's paragraph when an
            // object is selected, else at the caret's block (a caret inside a
            // table captions the table — the mutation hoists).
            const drawing = editor?.getSelectedDrawingSource();
            const caret = editor?.getCaretTarget();
            const anchorTarget = drawing ?? caret?.t;
            if (!anchorTarget) return false;
            let pEl: XmlElement | null = null;
            for (let cur: XmlElement | null = anchorTarget; cur; cur = doc.findParentOf(cur) ?? null) {
              if (localName(cur.name) === "p") {
                pEl = cur;
                break;
              }
            }
            const current = collabRef.current;
            if (current?.submitOp && doc.stableIds) {
              const blockId = pEl ? doc.stableIds.idOf(pEl) : undefined;
              if (blockId === undefined) return true; // honest no-op (see collabOp)
              history.checkpoint();
              current.submitOp(operationBody(
                "insertCaption",
                blockId,
                { label, ...(text ? { text } : {}), position },
                (n) => current.allocIds?.(n) ?? [],
              ) as never);
              return true;
            }
            history.checkpoint();
            if (!insertCaptionAt(doc, anchorTarget, label, text, position)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          updateFields: (values) => {
            // Bibliographies first — Word's F9 regenerates them too. Their
            // rebuild is STRUCTURAL (paragraphs replaced), so it is its own
            // operation rather than part of the string-carrying result pass;
            // the entries derive from the sources part on every replica, so
            // unlike TOC rebuilds it works in a room.
            const bibliographies = findBibliographyFields(doc).length;
            // Indexes rebuild the same way: structural, entries derived from
            // the XE marks on every replica; only the page numbers travel
            // (as PAGEREF results in the pass below). A structurally
            // unchanged index applies nothing and keeps its numbers.
            const indexes = findIndexFields(doc).length;
            if (collabRef.current?.submitOp) {
              if (bibliographies > 0) {
                collabDocOp((ids) =>
                  documentOperationBody(
                    "refreshBibliography",
                    { entryCount: Math.min(10000, bibliographyEntryCount(doc) * bibliographies) },
                    ids,
                  ),
                );
              }
              if (indexes > 0) {
                collabDocOp((ids) =>
                  documentOperationBody(
                    "refreshIndex",
                    { entryCount: Math.min(10000, indexEntryCount(doc) * indexes) },
                    ids,
                  ),
                );
              }
              // The results are computed HERE, on the acting client, and
              // carried on the wire. Page numbers come out of a layout and
              // layout depends on the host's font metrics, so a replica that
              // recomputed could install different text; see the registry's
              // updateFields comment. (A bibliography contributes no field
              // site — its result renders verbatim. An index DOES contribute
              // sites — its PAGEREF placeholders — but submitOp applies the
              // refresh optimistically and synchronously, so the site walk
              // below already sees the refreshed document, exactly as every
              // replica will at its sequenced position.)
              const results = computeFieldResults(doc, {
                layout: prevLayout ?? undefined,
                now: new Date(),
                ...values,
              });
              return collabDocOp(() => documentOperationBody("updateFields", { results }));
            }
            history.checkpoint();
            const rebuilt = bibliographies > 0 && refreshBibliographies(doc);
            const rebuiltIndex = indexes > 0 && refreshIndexes(doc);
            const results = computeFieldResults(doc, {
              layout: prevLayout ?? undefined,
              now: new Date(),
              ...values,
            });
            const applied = applyFieldResults(doc, results);
            if (!rebuilt && !rebuiltIndex && !applied) return false;
            pages = rerender(doc, undefined, "global");
            if (rebuiltIndex) {
              // A rebuilt index landed fresh PAGEREF placeholders; the layout
              // the rerender above produced is what fills them.
              if (applyFieldResults(doc, computeFieldResults(doc, { layout: prevLayout ?? undefined }))) {
                pages = rerender(doc, undefined, "global");
              }
            }
            return true;
          },
          insertToc: (options) => {
            // In a room the entries land with PLACEHOLDER page numbers. The
            // real ones come from a layout, which depends on this host's font
            // metrics, so they are the value updateFields carries as data
            // rather than a value a replica may recompute for itself.
            if (
              collabRunOperation("insertToc", {
                entryCount: tocEntryCount(doc, options),
                ...(options?.levels ? { levels: options.levels } : {}),
                ...(options?.leader ? { leader: options.leader } : {}),
                ...(options?.captionLabel ? { captionLabel: options.captionLabel } : {}),
              })
            ) {
              return true;
            }
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertToc(doc, target.t, options)) return false;
            // The entries land with placeholder page numbers; the layout this
            // rerender produces is what the update pass reads the real ones from.
            pages = rerender(doc, undefined, "global");
            if (applyFieldResults(doc, computeFieldResults(doc, { layout: prevLayout ?? undefined }))) {
              pages = rerender(doc, undefined, "global");
            }
            return true;
          },
          refreshTocs: () => {
            if (collabRef.current?.submitOp) return false;
            const tocs = findTocFields(doc);
            if (tocs.length === 0) return false;
            history.checkpoint();
            // Rebuilding replaces paragraphs, so each pass invalidates the
            // element handles the previous one found: re-find between rebuilds.
            let rebuilt = false;
            for (let i = 0; i < tocs.length; i++) {
              const current = findTocFields(doc)[i];
              if (current && rebuildToc(doc, current)) rebuilt = true;
            }
            if (!rebuilt) return false;
            pages = rerender(doc, undefined, "global");
            if (applyFieldResults(doc, computeFieldResults(doc, { layout: prevLayout ?? undefined }))) {
              pages = rerender(doc, undefined, "global");
            }
            return true;
          },
          listMergeFieldNames: () => documentMergeFieldNames(doc),
          listCitationSources: () => [...(documentBibliography(doc)?.sources.values() ?? [])],
          getCitationStyle: () => documentBibliography(doc)?.styleName ?? null,
          setCitationStyle: (style) => {
            if (collabDocOp(() => documentOperationBody("setCitationStyle", { style }))) return true;
            history.checkpoint();
            if (!setCitationStyle(doc, style)) return false;
            // The switch changes every CITATION's painted text and the
            // bibliography entries; the same call refreshes both, so the user
            // never sees a half-switched document.
            refreshBibliographies(doc);
            applyFieldResults(doc, computeFieldResults(doc, { layout: prevLayout ?? undefined }));
            pages = rerender(doc, undefined, "global");
            return true;
          },
          createCitationSource: (spec) => {
            if (collabDocOp(() => documentOperationBody("createCitationSource", { source: spec }))) return true;
            history.checkpoint();
            if (!createCitationSource(doc, spec)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          editCitationSource: (tag, patch) => {
            if (collabDocOp(() => documentOperationBody("editCitationSource", { tag, patch }))) return true;
            history.checkpoint();
            if (!editCitationSource(doc, tag, patch)) return false;
            // An edited source changes its citations' display and its
            // bibliography entry — refresh both, the setCitationStyle reason.
            refreshBibliographies(doc);
            applyFieldResults(doc, computeFieldResults(doc, { layout: prevLayout ?? undefined }));
            pages = rerender(doc, undefined, "global");
            return true;
          },
          deleteCitationSource: (tag) => {
            if (collabDocOp(() => documentOperationBody("deleteCitationSource", { tag }))) return true;
            history.checkpoint();
            if (!deleteCitationSource(doc, tag)) return false;
            refreshBibliographies(doc);
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertCitation: (tag) => {
            if (collabRunOperation("insertCitation", { tag })) return true;
            const target = insertionTarget();
            if (!target) return false;
            // The same predicate the registered apply uses: the tag must name
            // a source, and the display text derives from the sources part.
            const bibliography = documentBibliography(doc);
            if (!bibliography || !bibliography.sources.has(tag)) return false;
            history.checkpoint();
            const display = citationText(`CITATION ${tag} \\l 1033`, bibliography) ?? "";
            if (!insertCitationField(doc, target.t, target.offset, tag, display)) return false;
            pages = rerender(doc);
            return true;
          },
          insertBibliography: () => {
            if (collabRunOperation("insertBibliography", { entryCount: bibliographyEntryCount(doc) })) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertBibliography(doc, target.t)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          listBuildingBlocks: () => listBuildingBlocks(doc),
          createBuildingBlock: (name, category) => {
            const segments = (editor?.getSelectionSegments() ?? []).filter((s) => s.t);
            const blocks = selectionClipboardBlocks(doc, segments);
            if (blocks.length === 0) return false;
            const blocksXml = encodeClipboardOoxml(blocks);
            if (collabDocOp(() => documentOperationBody("createBuildingBlock", { name, category, blocksXml }))) {
              return true;
            }
            history.checkpoint();
            return createBuildingBlock(doc, { name, category, blocksXml });
          },
          insertBuildingBlock: (name) => {
            if (collabRunOperation("insertBuildingBlock", { name, blockCount: buildingBlockNodeCount(doc, name) })) {
              return true;
            }
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertBuildingBlock(doc, target.t, name)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          deleteBuildingBlock: (name) => {
            if (collabDocOp(() => documentOperationBody("deleteBuildingBlock", { name }))) return true;
            history.checkpoint();
            return deleteBuildingBlock(doc, name);
          },
          addIndexEntry: (entry) => {
            const segments = (editor?.getSelectionSegments() ?? []).filter((s) => s.t);
            const text = (entry ?? segments.map((s) => s.t!.text.slice(s.start, s.end)).join("")).trim();
            if (!isValidIndexEntry(text)) return false;
            // The mark lands AFTER the selection (Word's placement), else at
            // the caret.
            const last = segments[segments.length - 1];
            const at = last
              ? { t: last.t!, offset: Math.min(last.end, last.t!.text.length) }
              : insertionTarget();
            if (!at) return false;
            if (collabOp((a, ids) => operationBody("insertIndexEntry", a.runId, { entry: text }, ids) as never, at)) {
              return true;
            }
            history.checkpoint();
            if (!insertIndexEntry(doc, at.t, at.offset, text)) return false;
            pages = rerender(doc);
            return true;
          },
          insertIndex: () => {
            // In a room the page numbers land as placeholders — they come out
            // of a layout, the insertToc reason; updateFields carries the
            // real ones as data.
            if (collabRunOperation("insertIndex", { entryCount: indexEntryCount(doc) })) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertIndex(doc, target.t)) return false;
            pages = rerender(doc, undefined, "global");
            if (applyFieldResults(doc, computeFieldResults(doc, { layout: prevLayout ?? undefined }))) {
              pages = rerender(doc, undefined, "global");
            }
            return true;
          },
          insertEquation: (linear) => {
            if (collabOp((a, ids) => ({ kind: "insertMath", runId: a.runId, mathText: linear, nodeIds: ids(24) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertMathAt(doc, target.t, target.offset, linear)) return false;
            pages = rerender(doc);
            return true;
          },
          insertSymbol: (symbol) => editor?.insertText(symbol) ?? false,
          insertShape: (preset, text, lineStyle) => {
            // Collab: the line style is not carried by the intent yet — the
            // shape lands with its default outline everywhere (consistent).
            if (collabOp((a, ids) => ({ kind: "insertShape", runId: a.runId, preset, text, nodeIds: ids(12) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            const drawing = insertShapeAt(doc, target.t, preset, text);
            if (!drawing) return false;
            if (lineStyle) setDrawingLineStyle(doc, drawing, lineStyle.color, lineStyle.width, lineStyle.dash);
            pages = rerender(doc, undefined, "global");
            editor?.reselectDrawing(drawing);
            return true;
          },
          insertWordArt: (text, preset = "plain", style) => {
            if (collabOp((a, ids) => ({ kind: "insertWordArt", runId: a.runId, text, preset, ...(style ? { style } : {}), nodeIds: ids(12) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertWordArtAt(doc, target.t, text, preset, style)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertChart: (data) => {
            if (collabOp((a, ids) => ({ kind: "insertChart", runId: a.runId, chart: data, nodeIds: ids(24) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertChartAt(doc, target.t, data)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          updateSelectedChart: (data) => {
            // Collab: selected-object updates are not intent-anchored yet;
            // no-op rather than make a local-only (diverging) edit.
            if (collabRef.current?.submitOp) return false;
            const source = editor?.getSelectedDrawingSource();
            if (!source) return false;
            history.checkpoint();
            if (!setChartData(doc, source, data)) return false;
            pages = rerender(doc, undefined, "global");
            editor?.reselectDrawing(source);
            return true;
          },
          getSelectedChart: () => editor?.getSelectedChartData() ?? null,
          insertSmartArt: (data) => {
            if (collabOp((a, ids) => ({ kind: "insertSmartArt", runId: a.runId, smartArt: data, nodeIds: ids(24) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertSmartArtAt(doc, target.t, data)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          updateSelectedSmartArt: (data) => {
            // Collab: selected-object updates are not intent-anchored yet;
            // no-op rather than make a local-only (diverging) edit.
            if (collabRef.current?.submitOp) return false;
            const source = editor?.getSelectedDrawingSource();
            if (!source) return false;
            history.checkpoint();
            if (!setSmartArtData(doc, source, data)) return false;
            pages = rerender(doc, undefined, "global");
            editor?.reselectDrawing(source);
            return true;
          },
          getSelectedSmartArt: () => editor?.getSelectedSmartArtData() ?? null,
          insertModel3D: async (file, poster) => {
            // No model3D intent exists, so there is nothing to send. Refusing
            // is the only honest answer: the toolbar hides this in collab, but
            // the API is reachable directly, and mutating here would edit this
            // replica alone and tell nobody — a silent fork.
            if (collabRef.current?.submitOp) return false;
            const target = insertionTarget();
            if (!target) return false;
            const preview = poster ?? await objectPoster("3D model", "Double-click in Word to explore", "◇");
            const data = new Uint8Array(await file.arrayBuffer());
            const posterBytes = new Uint8Array(await preview.arrayBuffer());
            history.checkpoint();
            if (!insertModel3DAt(doc, target.t, { data, poster: posterBytes })) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertOnlineVideo: async (url) => {
            // Same as insertModel3D: no intent carries a web-video reference,
            // so a local insert here would fork the room silently.
            if (collabRef.current?.submitOp) return false;
            const target = insertionTarget();
            if (!target) return false;
            const preview = await objectPoster("Online video", url, "▶");
            const poster = new Uint8Array(await preview.arrayBuffer());
            history.checkpoint();
            if (!insertWebVideoAt(doc, target.t, { url, poster })) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertEmbeddedObject: async (file, filename) => {
            // Same as insertModel3D: no intent carries an OLE part, so a local
            // insert here would fork the room silently.
            if (collabRef.current?.submitOp) return false;
            const target = insertionTarget();
            if (!target) return false;
            const name = filename ?? (file instanceof File ? file.name : "embedded-file.bin");
            const preview = await objectPoster("Embedded object", name, "▤");
            const data = new Uint8Array(await file.arrayBuffer());
            const poster = new Uint8Array(await preview.arrayBuffer());
            history.checkpoint();
            if (!insertEmbeddedObjectAt(doc, target.t, { data, filename: name, poster })) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          setDrawingTool: (tool) => editor?.setDrawingTool(tool),
          getDrawingTool: () => editor?.getDrawingTool() ?? null,
          arrangeObject: (action) => editor?.arrangeSelectedObject(action) ?? false,
          hasSelectedObject: () => editor?.hasSelectedObject() ?? false,
          getSelectedObjectContext: () => editor?.getSelectedObjectContext() ?? null,
          runSelectedObjectCommand: (command) => editor?.runSelectedObjectCommand(command) ?? false,
          getSelectedSmartArtTextFormat: () => editor?.getSelectedSmartArtTextFormat() ?? null,
          setSelectedSmartArtTextFormat: (patch) => editor?.setSelectedSmartArtTextFormat(patch) ?? false,
          addComment: (text) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const segments = segs.length > 0 ? segs : handle ? selectionToSegments(handle.bindings) : [];
            if (segments.length === 0) return false;
            {
              // Collab: whole-run comment on the first selected run, with
              // carried provenance so every replica writes identical XML.
              const seg0 = segments.find((sg) => sg.t);
              const initials0 = commentAuthor.split(/\s+/).map((part) => part[0] ?? "").join("").slice(0, 2).toUpperCase();
              if (seg0?.t && collabOp(
                (a) => ({ kind: "commentRun", runId: a.runId, text, author: commentAuthor, initials: initials0 || undefined, date: new Date().toISOString(), paraId: hex8() }),
                { t: seg0.t, offset: seg0.start },
              )) return true;
            }
            const initials = commentAuthor
              .split(/\s+/)
              .map((part) => part[0] ?? "")
              .join("")
              .slice(0, 2)
              .toUpperCase();
            history.checkpoint();
            if (addComment(doc, segments, text, commentAuthor, initials || undefined)) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
          undo: () => editor?.applyHistory("undo"),
          redo: () => editor?.applyHistory("redo"),
          canUndo: () => history.canUndo,
          canRedo: () => history.canRedo,
          insertTable: (rows, cols) => {
            if (collabRunOperation("insertTable", { rows, cols })) return;
            const caret = editor?.getCaretTarget();
            if (!caret) return;
            history.checkpoint();
            if (insertTableAfter(doc, caret.t, rows, cols)) pages = rerender(doc);
          },
          tableOp: (op) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return;
            // One draw per gesture, as for the run and paragraph commands: the
            // same author+date reach the local *PrChange and the intent every
            // replica applies. The structural ops have no tracked form and
            // core's applyTableOp ignores the meta for them, so this passes it
            // unconditionally rather than restating which kinds honour it.
            const suggest = editor?.suggestionMeta();
            if (collabOp((a, ids) => ({ kind: "tableOp", cellParagraphId: a.blockId, op, nodeIds: ids(16), ...(suggest ? { suggest } : {}) }), { t: caret.t, offset: 0 })) return;
            history.checkpoint();
            if (applyTableOp(doc, caret.t, op, suggestMeta(doc, suggest))) pages = rerender(doc);
          },
          getTableCellFill: () => {
            const caret = editor?.getCaretTarget();
            return caret ? cellShadingAt(doc, caret.t) : undefined;
          },
          getTableProperties: () => {
            const caret = editor?.getCaretTarget();
            return caret ? readTableProperties(doc, caret.t) : undefined;
          },
          setTableBorders: (scope, edges, border) =>
            runTableOperation("setTableBorders", { scope, edges, border }, (caret, meta) =>
              setTableBorders(doc, caret, scope, edges, border, meta),
            ),
          listTableStyles: () => listTableStyles(doc),
          getTableStyleId: () => {
            const tbl = caretTable();
            if (!tbl) return undefined;
            const tblPr = tbl.children.find((c) => localName(c.name) === "tblPr");
            const style = tblPr?.children.find((c) => localName(c.name) === "tblStyle");
            return style ? attr(style, "val") ?? null : null;
          },
          setTableStyle: (styleId) =>
            runTableOperation("setTableStyle", { styleId }, (_caret, meta) => {
              const tbl = caretTable();
              return tbl ? setTableStyle(doc, tbl, styleId, meta) : false;
            }),
          getTableLook: () => {
            const tbl = caretTable();
            return tbl ? tableLookOf(tbl) : undefined;
          },
          setTableLook: (patch) =>
            runTableOperation("setTableLook", { look: patch }, (_caret, meta) => {
              const tbl = caretTable();
              return tbl ? setTableLook(doc, tbl, patch, meta) : false;
            }),
          setTableWidth: (unit, value) =>
            runTableOperation("setTableWidth", { unit, value }, (_caret, meta) => {
              const tbl = caretTable();
              return tbl ? setTableWidth(doc, tbl, unit, value ?? 0, meta) : false;
            }),
          setTableColumnWidth: (colIdx, widthPt) =>
            runTableOperation("setTableColumnWidth", { colIdx, widthPt }, (_caret, meta) => {
              const tbl = caretTable();
              return tbl ? setTableColumnWidth(doc, tbl, colIdx, widthPt, meta) : false;
            }),
          setTableLayout: (layout) => {
            const tbl = caretTable();
            if (!tbl) return;
            // Measured on THIS replica, then sent as data: freezing an autofit
            // table has to keep the columns the user can see, and a replica
            // that re-measured them itself could freeze a different table.
            const renderedWidths = layout === "fixed" ? renderedColumnWidths(tbl) : undefined;
            runTableOperation("setTableLayout", { layout, renderedWidths }, (_caret, meta) =>
              setTableLayoutMode(doc, tbl, layout, renderedWidths, meta),
            );
          },
          setTableCellMargins: (scope, margins) =>
            runTableOperation("setTableCellMargins", { scope, margins }, (caret, meta) =>
              setTableCellMargins(doc, caret, scope, margins, meta),
            ),
          setTableHeaderRows: (count) =>
            runTableOperation("setTableHeaderRows", { count }, (_caret, meta) => {
              const tbl = caretTable();
              return tbl ? setTableHeaderRows(doc, tbl, count, meta) : false;
            }),
          sortTableRows: (colIdx, order, compare, hasHeader) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return;
            // STRUCTURAL like insertTable: no tracked form, so no suggest
            // payload rides along — the sort applies untracked in any mode.
            const args = { colIdx, order, compare, ...(hasHeader !== undefined ? { hasHeader } : {}) };
            if (collabOp((a) => operationBody("sortTableRows", a.blockId, args) as never, { t: caret.t, offset: 0 })) return;
            history.checkpoint();
            const tbl = caretTable();
            if (tbl && sortTableRows(doc, tbl, colIdx, order, compare, hasHeader ?? false)) pages = rerender(doc);
          },
          insertTableFormula: (formula, numFmt) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return false;
            const p = paragraphOf(doc, caret.t);
            // Refuse outside a table BEFORE the collab gate: the wire apply
            // would be an honest no-op there anyway, and a false return is
            // what lets the dialog say so.
            if (!p || !cellContextOf(doc, caret.t)) return false;
            const args = { formula, ...(numFmt !== undefined ? { numFmt } : {}) };
            if (
              collabOp(
                (a, alloc) => operationBody("insertTableFormula", a.blockId, args, alloc) as never,
                { t: caret.t, offset: 0 },
              )
            ) return true;
            history.checkpoint();
            if (!insertTableFormula(doc, p, formula, numFmt)) return false;
            pages = rerender(doc);
            return true;
          },
          convertTextToTable: (separator) => {
            const caret = editor?.getCaretTarget();
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : caret ? [caret.t] : [];
            if (targets.length === 0) return false;
            const paras: XmlElement[] = [];
            for (const t of targets) {
              const p = paragraphOf(doc, t);
              if (p && !paras.includes(p)) paras.push(p);
            }
            if (paras.length === 0) return false;
            // The id budget rides as data: rows × columns as THIS replica
            // computed them (the insertToc entryCount pattern).
            const sep = separator === "tab" ? "\t" : ",";
            const cols = Math.max(1, ...paras.map((p) => plainTextOf(p).split(sep).length));
            const cellCount = Math.min(paras.length * cols, 10_000);
            const blockIds = doc.stableIds
              ? paras.map((p) => doc.stableIds!.idOf(p)).filter((n): n is number => n !== undefined)
              : [];
            if (
              collabOp(
                (anchor, alloc) =>
                  blockIds.length === 0
                    ? null
                    : (operationBody("convertTextToTable", blockIds[0], {
                        separator,
                        cellCount,
                        ...(blockIds.length > 1 ? { moreBlockIds: blockIds.slice(1) } : {}),
                      }, alloc) as never),
                { t: targets[0], offset: 0 },
              )
            ) return true;
            history.checkpoint();
            if (!convertTextToTable(doc, paras, separator)) return false;
            pages = rerender(doc);
            return true;
          },
          convertTableToText: (separator) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return false;
            const tbl = caretTable();
            if (!tbl) return false;
            const rowCount = Math.min(
              Math.max(tbl.children.filter((c) => localName(c.name) === "tr").length, 1),
              10_000,
            );
            if (
              collabOp(
                (anchor, alloc) => operationBody("convertTableToText", anchor.blockId, { separator, rowCount }, alloc) as never,
                { t: caret.t, offset: 0 },
              )
            ) return true;
            history.checkpoint();
            if (!convertTableToText(doc, tbl, separator)) return false;
            pages = rerender(doc);
            return true;
          },
          imageAccept: () => (collabRef.current?.submitOp && doc.stableIds ? COLLAB_IMAGE_ACCEPT : LOCAL_IMAGE_ACCEPT),
          imageMaxBytes: () => (collabRef.current?.submitOp && doc.stableIds ? collabRef.current.mediaMaxBlobBytes ?? null : null),
          insertImage: async (file) => {
            // Caret, else the selection's end, else the start of the document
            // (see documentStart): picking a file is an unambiguous request to
            // insert one, so answering it with silence because no caret has
            // been placed yet is a bug, not a safeguard.
            const caret = insertionTarget() ?? documentStart();
            if (!caret) return "no-caret";
            // SIZE PRE-CHECK, before a single byte is read. The relay would
            // refuse this file anyway, but only after the browser had read it,
            // decoded it, sealed it, hashed it and pushed it over the network —
            // so the user waits out a whole upload to be told no. `file.size`
            // is known immediately and costs nothing.
            //
            // ONE BRANCH, and the null case is the important half: no published
            // limit means SKIP, never substitute a default. The server enforces
            // the real number regardless, so skipping risks nothing, while
            // guessing would either block uploads it would have accepted or
            // promise a size it will refuse.
            const maxBytes = collabRef.current?.submitOp && doc.stableIds ? collabRef.current.mediaMaxBlobBytes ?? null : null;
            if (typeof maxBytes === "number" && file.size > maxBytes) return "too-large";
            const bytes = new Uint8Array(await file.arrayBuffer());
            const isSvg = file.type === "image/svg+xml";
            const bmp = isSvg ? null : await createImageBitmap(new Blob([bytes.buffer as ArrayBuffer], { type: file.type }));
            const svgRoot = isSvg
              ? new DOMParser().parseFromString(new TextDecoder().decode(bytes), "image/svg+xml").documentElement
              : null;
            const viewBox = (svgRoot?.getAttribute("viewBox") ?? "").trim().split(/[\s,]+/).map(Number);
            const naturalWidth = (bmp?.width ?? parseFloat(svgRoot?.getAttribute("width") ?? "")) || viewBox[2] || 96;
            const naturalHeight = (bmp?.height ?? parseFloat(svgRoot?.getAttribute("height") ?? "")) || viewBox[3] || 96;
            const sp = doc.sections[0]?.props;
            const maxW = sp ? sp.pageWidth - sp.marginLeft - sp.marginRight : 624;
            const scale = Math.min(1, maxW / naturalWidth);
            const ext = file.type === "image/svg+xml" ? "svg" : (file.type.split("/")[1] ?? "png").replace("jpeg", "jpg");
            // COLLAB (plan doc 16 §5.1): bytes travel out of band and the
            // intent carries only their address, so the ORDER is load-bearing
            // — upload first, reserve only on success. Reserving first would
            // leave every other replica with a skeleton pointing at a blob
            // that was refused and will never exist.
            const current = collabRef.current;
            if (current?.submitOp && doc.stableIds) {
              bmp?.close();
              // Past the collab gate this NEVER falls through to the local
              // mutation (checkpoint A18): an app that wired submitOp but no
              // relay has no way to place an image, and inserting one locally
              // would fork the room the moment anyone else typed. Honest
              // no-op instead. Caught by INVARIANT C, which mounts exactly
              // that configuration.
              if (!current.uploadMedia) return "no-relay";
              // The wire allowlist is raster-only (validate.ts): an SVG would
              // be refused by every replica, so decline it here rather than
              // inserting something only this client can see. The picker is
              // built from this same list (imageAccept) so this is the last
              // line of defence, not the first — a user should never reach it
              // through the toolbar, only through the API.
              if (!(COLLAB_IMAGE_EXTS as readonly string[]).includes(ext)) return "unsupported-format";
              const media = await current.uploadMedia(bytes);
              if (!media) return "upload-failed"; // relay refused — nothing reserved, nothing forked
              const w = naturalWidth * scale;
              const h2 = naturalHeight * scale;
              collabOp((a, ids) => ({
                kind: "insertImage",
                runId: a.runId,
                blobSha: media.blobSha,
                bytesLen: media.bytesLen,
                ext,
                ...(media.iv ? { iv: media.iv } : {}),
                widthPx: w,
                heightPx: h2,
                nodeIds: ids(8),
              }), { t: caret.t, offset: caret.offset });
              // The placer's own pixels install with no round trip — the
              // media client kept the plaintext it just uploaded and fills the
              // part from memory the moment the reservation applies (which in
              // an encrypted room is asynchronous, after this returns).
              pages = rerender(doc);
              return "inserted";
            }
            history.checkpoint();
            const relId = doc.addImageResource(bytes, ext === "jpg" ? "jpeg" : ext);
            const h = naturalHeight * scale;
            const drawing = insertImageAt(doc, caret.t, relId, naturalWidth * scale, h);
            if (drawing) {
              // An image taller than an "exact"-spaced line would be clipped
              // (Word) or overlap neighbors — float it with square wrap.
              const exact = exactLineHeightAt(doc, caret.t);
              if (exact !== null && h > exact + 0.5) {
                setImageWrap(doc, drawing, "square", { x: 0, y: 0 });
              }
              pages = rerender(doc);
            }
            bmp?.close();
            // No drawing means the caret's paragraph could not carry one — an
            // outcome, not a non-event.
            return drawing ? "inserted" : "error";
          },
          insertScreenshot: async () => {
            if (!editor?.getCaretTarget()) return "no-caret";
            if (!navigator.mediaDevices?.getDisplayMedia) return "unsupported";
            let stream: MediaStream | null = null;
            try {
              stream = await navigator.mediaDevices.getDisplayMedia({ video: true });
              const video = document.createElement("video");
              video.muted = true;
              video.playsInline = true;
              video.srcObject = stream;
              const ready = new Promise<void>((resolve, reject) => {
                if (video.readyState >= HTMLMediaElement.HAVE_METADATA && video.videoWidth > 0 && video.videoHeight > 0) {
                  resolve();
                  return;
                }
                video.addEventListener("loadedmetadata", () => resolve(), { once: true });
                video.addEventListener("error", () => reject(new Error("Captured video is unavailable")), { once: true });
              });
              await video.play();
              await ready;
              const canvas = document.createElement("canvas");
              canvas.width = video.videoWidth;
              canvas.height = video.videoHeight;
              const context = canvas.getContext("2d");
              if (!context || canvas.width === 0 || canvas.height === 0) return "error";
              context.drawImage(video, 0, 0);
              const image = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
              if (!image) return "error";
              // Never claim a success the insert didn't have: the capture can
              // work and the PLACEMENT still fail (no relay in a shared
              // document, a paragraph that can't carry a drawing). Reporting
              // "inserted" here regardless is the same silence one layer up.
              const placed = await api.insertImage(image);
              return placed === "inserted" ? "inserted" : placed === "no-caret" ? "no-caret" : "error";
            } catch (error) {
              return error instanceof DOMException && (error.name === "NotAllowedError" || error.name === "AbortError")
                ? "cancelled"
                : "error";
            } finally {
              for (const track of stream?.getTracks() ?? []) track.stop();
            }
          },
          setAlignment: (align) => {
            // The EDITOR's selection, like every other paragraph command. The
            // editor paints its own highlight and keeps the DOM selection in
            // its hidden input sink, so window.getSelection() is empty while a
            // range is selected — reading it here aligned whatever paragraph
            // the caret was left in by the PREVIOUS gesture instead.
            const caret = editor?.getCaretTarget();
            const segs = editor?.getSelectionSegments() ?? [];
            const segTs = segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t);
            const targets = segTs.length > 0 ? segTs : caret ? [caret.t] : [];
            if (targets.length === 0) return;
            history.checkpoint();
            const suggest = editor?.suggestionMeta();
            if (setParagraphAlignment(doc, targets as Parameters<typeof setParagraphAlignment>[1], align, suggestMeta(doc, suggest))) {
              emitBlockIntents(targets, (blockId) => ({ kind: "formatParagraph", blockId, align, ...(suggest ? { suggest } : {}) }));
              pages = rerender(doc);
            }
          },
          closeHeaderFooter: () => editor?.exitHeaderFooter(),
          openHeaderFooter: (kind) => editor?.enterHeaderFooter(kind) ?? false,
          setDifferentFirstPage: (on) => {
            if (collabDocOp((ids) => documentOperationBody("setTitlePage", { enabled: on }, ids))) return true;
            history.checkpoint();
            if (!setTitlePage(doc, on)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          getDifferentFirstPage: () => titlePageEnabled(doc),
          setOddEvenHeaders: (on) => {
            if (collabDocOp((ids) => documentOperationBody("setEvenOddHeaders", { enabled: on }, ids))) return true;
            history.checkpoint();
            if (!setEvenOddHeaders(doc, on)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          getOddEvenHeaders: () => doc.evenAndOddHeaders,
          setPageNumberFormat: (patch, scope) => {
            // Collab: document-level; a "section" scope falls back to all
            // sections (consistent everywhere, demo limitation — see
            // setPageLayout).
            if (collabDocOp(() => documentOperationBody("setPageNumberFormat", patch))) return true;
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (!setPageNumberFormat(doc, patch, target)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          getPageNumberFormat: () => {
            const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t ?? documentStart()?.t;
            return t ? pageNumberFormatAt(doc, t) : { fmt: "decimal", start: null };
          },
          setHyphenation: (patch) => {
            if (collabDocOp(() => documentOperationBody("setHyphenation", patch))) return true;
            history.checkpoint();
            // Settings-only state: this engine's layout does not hyphenate,
            // so nothing painted changes and no rerender is needed.
            return doc.setHyphenation({
              ...(patch.auto !== undefined ? { auto: patch.auto } : {}),
              ...(patch.zonePt !== undefined
                ? { zoneTwips: patch.zonePt === null ? null : Math.round(patch.zonePt * 20) }
                : {}),
              ...(patch.noCaps !== undefined ? { noCaps: patch.noCaps } : {}),
            });
          },
          getHyphenation: () => ({
            auto: doc.autoHyphenation,
            zonePt: doc.hyphenationZoneTwips === null ? null : doc.hyphenationZoneTwips / 20,
            noCaps: doc.doNotHyphenateCaps,
          }),
          insertPageNumberPosition: (position, align) => {
            if (collabDocOp((ids) => documentOperationBody("insertPageNumberPosition", { position, align }, ids))) return true;
            history.checkpoint();
            if (!insertPageNumberPosition(doc, position, align)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          removePageNumbers: () => {
            if (collabDocOp(() => documentOperationBody("removePageNumbers", {}))) return true;
            history.checkpoint();
            if (!removePageNumberFields(doc)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertHeaderFooterPreset: (kind, preset) => {
            if (collabDocOp((ids) => documentOperationBody("insertHeaderFooterPreset", { hfKind: kind, preset }, ids))) return true;
            history.checkpoint();
            if (!insertHeaderFooterPreset(doc, kind, preset)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          setFootnoteOptions: (patch, scope) => {
            // Same document/section-scope shape as setPageNumberFormat.
            if (collabDocOp(() => documentOperationBody("setFootnoteOptions", patch))) return true;
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (!setFootnoteOptions(doc, patch, target)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          getFootnoteOptions: () => {
            const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t ?? documentStart()?.t;
            return t ? footnoteOptionsAt(doc, t) : { fmt: "decimal", start: null, restart: "continuous", pos: "pageBottom" };
          },
          setEndnoteOptions: (patch, scope) => {
            if (collabDocOp(() => documentOperationBody("setEndnoteOptions", patch))) return true;
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (!setEndnoteOptions(doc, patch, target)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          getEndnoteOptions: () => {
            const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t ?? documentStart()?.t;
            return t ? endnoteOptionsAt(doc, t) : { fmt: "lowerRoman", start: null, restart: "continuous", pos: "docEnd" };
          },
          resolveComment: (id, resolved) => {
            if (collabDocOp(() => ({ kind: "resolveComment", commentId: id, resolved, paraId: hex8() }))) return true;
            history.checkpoint();
            if (!setCommentResolved(doc, id, resolved)) return false;
            pages = rerender(doc);
            return true;
          },
          editComment: (id, text) => {
            if (collabDocOp(() => ({ kind: "editComment", commentId: id, text }))) return true;
            history.checkpoint();
            if (!editCommentText(doc, id, text)) return false;
            pages = rerender(doc);
            return true;
          },
          stepComment: (delta) => {
            const anchors = doc.commentAnchors();
            const threads = doc.comments.filter(
              (c) => !c.parentId && (anchors.get(c.id)?.length ?? 0) > 0,
            );
            if (threads.length === 0) return null;
            commentNav = commentNav < 0
              ? (delta === 1 ? 0 : threads.length - 1)
              : (commentNav + delta + threads.length) % threads.length;
            const comment = threads[commentNav];
            const ranges = anchors.get(comment.id)!.map((t) => ({ t, start: 0, end: t.text.length }));
            if (editor && ranges.length > 0) {
              // Bring the anchor into view (same dance as find's selectMatch:
              // a virtualized page must be mounted before it can scroll).
              const restore = handle?._virtualized ? handle.materializeAll?.() : undefined;
              editor.selectRanges(ranges);
              const el = handle?.bindingsByText.get(ranges[0].t)?.[0]?.el;
              el?.scrollIntoView({ block: "center", behavior: "instant" as ScrollBehavior });
              restore?.();
              handle?.updateViewport?.();
              editor.selectRanges(ranges);
            }
            return comment.id;
          },
          wordCount: () => ({ ...documentTextStatistics(doc), pages }),
          insertBreak: (kind) => {
            let target = editor?.getCaretTarget() ?? null;
            if (!target) {
              const segs = editor?.getSelectionSegments() ?? [];
              const last = [...segs].reverse().find((sg) => sg.t);
              if (last?.t) target = { t: last.t, offset: last.end };
            }
            if (!target) return false;
            if (kind === "page" || kind === "column"
              ? collabOp((a, ids) => ({ kind: "insertBreak", runId: a.runId, breakKind: kind, nodeIds: ids(8) }), target)
              : collabOp((a, ids) => ({ kind: "insertSectionBreak", runId: a.runId, breakType: kind === "sectionNextPage" ? "nextPage" : "continuous", nodeIds: ids(8) }), target)) return true;
            history.checkpoint();
            const done =
              kind === "page" || kind === "column"
                ? insertBreakAt(doc, target.t, target.offset, kind)
                : insertSectionBreak(doc, target.t, kind === "sectionNextPage" ? "nextPage" : "continuous");
            if (done) pages = rerender(doc);
            return !!done;
          },
          insertBlankPage: () => {
            if (collabOp((a, ids) => ({ kind: "insertBlankPage", runId: a.runId, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertBlankPageAt(doc, target.t, target.offset)) return false;
            pages = rerender(doc);
            return true;
          },
          insertCoverPage: (content) => {
            if (collabDocOp((ids) => ({ kind: "insertCoverPage", content: content as unknown as Record<string, unknown>, nodeIds: ids(24) }))) return true;
            history.checkpoint();
            if (!insertCoverPage(doc, content)) return false;
            pages = rerender(doc);
            return true;
          },
          setPageLayout: (patch, scope) => {
            // Collab: the intent is document-level; a "section" scope falls
            // back to all sections (consistent everywhere, demo limitation).
            if (collabDocOp(() => ({ kind: "setPageLayout", patch: patch as Record<string, unknown> }))) return;
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (setPageLayout(doc, patch, target)) pages = rerender(doc);
          },
          getSectionContext: () => {
            const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
            return t ? sectionContextAt(doc, t) : null;
          },
          setLineNumbering: (patch, scope) => {
            if (collabDocOp(() => ({ kind: "setLineNumbering", patch: patch as unknown as Record<string, unknown> }))) return;
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (setLineNumbering(doc, patch, target)) pages = rerender(doc);
          },
          getLineNumbering: () => {
            const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
            return t ? lineNumberingAt(doc, t) : null;
          },
          setLink: (url) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const t = segs.find((sg) => sg.t)?.t ?? editor?.getCaretTarget()?.t;
            if (t && (url === null
              ? collabOp((a) => ({ kind: "removeLink", runId: a.runId }), { t, offset: 0 })
              : collabOp((a, ids) => ({ kind: "setLink", runId: a.runId, url, nodeIds: ids(4) }), { t, offset: 0 }))) return;
            history.checkpoint();
            const changed = url === null ? (t ? removeLink(doc, t) : false) : setLink(doc, segs, url);
            if (changed) pages = rerender(doc);
          },
          getLinkAt: () => {
            const segs = editor?.getSelectionSegments() ?? [];
            const t = segs.find((sg) => sg.t)?.t ?? editor?.getCaretTarget()?.t;
            return t ? linkAt(doc, t) : null;
          },
          adjustIndent: (direction) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : editor?.getCaretTarget() ? [editor.getCaretTarget()!.t] : [];
            if (targets.length === 0) return;
            const suggest = editor?.suggestionMeta();
            if (collabBlockOp(targets, (blockId) => ({ kind: "adjustIndent", blockId, direction, ...(suggest ? { suggest } : {}) }))) return;
            history.checkpoint();
            if (adjustIndent(doc, targets as Parameters<typeof adjustIndent>[1], direction, suggestMeta(doc, suggest))) pages = rerender(doc);
          },
          setParagraphSpacing: (patch) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : editor?.getCaretTarget() ? [editor.getCaretTarget()!.t] : [];
            if (targets.length === 0) return;
            const suggest = editor?.suggestionMeta();
            if (collabBlockOp(targets, (blockId) => ({ kind: "setSpacing", blockId, patch: patch as Record<string, unknown>, ...(suggest ? { suggest } : {}) }))) return;
            history.checkpoint();
            if (setParagraphSpacing(doc, targets as Parameters<typeof setParagraphSpacing>[1], patch, suggestMeta(doc, suggest))) pages = rerender(doc);
          },
          setParagraphDivider: (divider) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const caret = editor?.getCaretTarget();
            const targets = segs.length > 0
              ? segs.map((segment) => segment.t).filter((target): target is NonNullable<typeof target> => !!target)
              : caret
                ? [caret.t]
                : [];
            if (targets.length === 0) return false;
            if (collabBlockOp(targets, (blockId) => ({ kind: "setDivider", blockId, divider: divider as Record<string, unknown> | null }))) return true;
            history.checkpoint();
            if (!setParagraphDivider(doc, targets, divider)) return false;
            pages = rerender(doc);
            return true;
          },
          getParagraphDivider: () => {
            const target = editor?.getSelectionSegments()?.find((segment) => segment.t)?.t ?? editor?.getCaretTarget()?.t;
            return target ? paragraphDividerAt(doc, target) : null;
          },
          setTabStops: (stops) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const caret = editor?.getCaretTarget();
            const targets = segs.length > 0
              ? segs.map((segment) => segment.t).filter((target): target is NonNullable<typeof target> => !!target)
              : caret
                ? [caret.t]
                : [];
            if (targets.length === 0) return false;
            const suggest = editor?.suggestionMeta();
            if (collabBlockOp(targets, (blockId) => operationBody("setTabStops", blockId, { stops, suggest }) as never)) return true;
            history.checkpoint();
            if (!setTabStops(doc, targets, stops, suggestMeta(doc, suggest))) return false;
            pages = rerender(doc);
            return true;
          },
          getTabStops: () => {
            const target = editor?.getSelectionSegments()?.find((segment) => segment.t)?.t ?? editor?.getCaretTarget()?.t;
            return target ? tabStopsAt(doc, target) : [];
          },
          setParagraphBorders: (patch) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const caret = editor?.getCaretTarget();
            const targets = segs.length > 0
              ? segs.map((segment) => segment.t).filter((target): target is NonNullable<typeof target> => !!target)
              : caret
                ? [caret.t]
                : [];
            if (targets.length === 0) return false;
            const suggest = editor?.suggestionMeta();
            if (collabBlockOp(targets, (blockId) => operationBody("setParagraphBorders", blockId, { patch, suggest }) as never)) return true;
            history.checkpoint();
            if (!setParagraphBorders(doc, targets, patch, suggestMeta(doc, suggest))) return false;
            pages = rerender(doc);
            return true;
          },
          getParagraphBorders: () => {
            const target = editor?.getSelectionSegments()?.find((segment) => segment.t)?.t ?? editor?.getCaretTarget()?.t;
            return target ? paragraphBordersAt(doc, target) : { borders: {}, shading: null };
          },
          setDropCap: (mode, lines = 3) => {
            if (collabOp((a, ids) => ({ kind: "setDropCap", blockId: a.blockId, mode, nodeIds: ids(8) }))) return true;
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!setDropCapAt(doc, target.t, mode, lines)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          clearFormatting: () => {
            api.applyFormat({ clear: true });
          },
          changeCase: (mode) => {
            const segs = editor?.getSelectionSegments() ?? [];
            if (segs.length === 0) return;
            const selectedAll = editor?.isEntireDocumentSelected() ?? false;
            history.checkpoint();
            const changed = transformCase(doc, segs, mode);
            if (changed.length > 0) {
              pages = rerender(doc);
              if (selectedAll) editor?.selectAll();
              else editor?.selectRanges(changed);
            }
          },
          find: (query, opts) => {
            findState = { matches: findAll(doc, query, opts), index: 0 };
            if (findState.matches.length > 0) selectMatch(0);
            return findState.matches.length;
          },
          findStep: (delta) => {
            if (findState.matches.length === 0) return 0;
            findState.index = (findState.index + delta + findState.matches.length) % findState.matches.length;
            selectMatch(findState.index);
            return findState.index + 1;
          },
          replaceCurrent: (replacement) => {
            const m = findState.matches[findState.index];
            if (!m) return 0;
            const current = collabRef.current;
            if (current?.submitOp && doc.stableIds) {
              // In a room the replacement rides the wire as the canonical
              // deleteText/insertText intents (strike-then-insert while
              // suggesting) — the local mutation below never replicates. An
              // unaddressable match is an honest no-op (see collabOp).
              const intents = compileReplaceMatch(doc, m, replacement, editor?.suggestionMeta());
              if (intents) {
                history.checkpoint();
                for (const intent of intents) current.submitOp(intent);
              }
            } else {
              history.checkpoint();
              replaceMatch(doc, m, replacement);
              pages = rerender(doc);
            }
            findState.matches.splice(findState.index, 1);
            if (findState.index >= findState.matches.length) findState.index = 0;
            if (findState.matches.length > 0) selectMatch(findState.index);
            return findState.matches.length;
          },
          replaceAll: (query, replacement, opts) => {
            const current = collabRef.current;
            if (current?.submitOp && doc.stableIds) {
              // One fixed find pass compiled to per-match intents, submitted
              // back-to-front (the compiled order) so every offset encoded
              // against the pre-replace tree stays valid as the optimistic
              // applies land. One checkpoint: the whole sweep is one gesture.
              const { intents, result } = compileReplaceAll(doc, query, replacement, opts, editor?.suggestionMeta());
              if (intents.length > 0) {
                history.checkpoint();
                for (const intent of intents) current.submitOp(intent);
              }
              findState = { matches: [], index: 0 };
              return result;
            }
            history.checkpoint();
            const result = replaceAll(doc, query, replacement, opts);
            if (result.total > 0) pages = rerender(doc);
            findState = { matches: [], index: 0 };
            return result;
          },
          goToPage: (page) => {
            if (!handle) return false;
            // A far page may be virtualized out: mount all pages, scroll,
            // restore the window — find navigation's pattern (selectMatch).
            const restore = handle._virtualized ? handle.materializeAll?.() : undefined;
            const el = handle.root.querySelectorAll<HTMLElement>(".dxw-page")[page - 1];
            el?.scrollIntoView({ block: "start", behavior: "instant" as ScrollBehavior });
            restore?.();
            handle.updateViewport?.();
            return !!el;
          },
          goToBookmark: (name) => {
            const target = bookmarkTextTarget(doc, name);
            if (!target || !handle) return false;
            const restore = handle._virtualized ? handle.materializeAll?.() : undefined;
            editor?.selectRanges([{ t: target.t, start: target.offset, end: target.offset }]);
            const el = handle.bindingsByText.get(target.t)?.[0]?.el;
            el?.scrollIntoView({ block: "center", behavior: "instant" as ScrollBehavior });
            restore?.();
            handle.updateViewport?.();
            return !!el;
          },
          selectRange: (range) => {
            if (!editor) return false;
            const ids = doc.stableIds ?? doc.enableStableIds();
            const resolved: { t: XmlElement; start: number; end: number }[] = [];
            for (const r of Array.isArray(range) ? range : [range]) {
              if (!Number.isInteger(r.start) || !Number.isInteger(r.end) || r.end <= r.start || r.start < 0) continue;
              const runEl = ids.elOf(r.runId);
              if (!runEl) continue;
              resolved.push(...resolveWireRange(runEl, r.start, r.end));
            }
            if (resolved.length === 0) return false;
            // Find-navigation's materialize dance (selectMatch): a far range
            // may live on a virtualized-out page. "nearest" instead of
            // "center" so a range already on screen — the spellcheck's word
            // under the pointer — never jumps the view.
            const restore = handle?._virtualized ? handle.materializeAll?.() : undefined;
            editor.selectRanges(resolved);
            const el = handle?.bindingsByText.get(resolved[0].t)?.[0]?.el;
            el?.scrollIntoView({ block: "nearest", behavior: "instant" as ScrollBehavior });
            restore?.();
            handle?.updateViewport?.();
            editor.selectRanges(resolved);
            return true;
          },
          toggleList: (kind) => {
            const caret = editor?.getCaretTarget();
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : caret ? [caret.t] : [];
            if (targets.length === 0) return;
            const current = listTypeAt(doc, targets[0]);
            const dirtyBlock = targets.length === 1 ? topLevelBlockOf(doc, targets[0]) ?? undefined : undefined;
            history.checkpoint();
            const listKind = current === kind ? null : kind;
            const suggest = editor?.suggestionMeta();
            if (setListType(doc, targets as Parameters<typeof setListType>[1], listKind, suggestMeta(doc, suggest))) {
              // ONE intent for the whole selection (moreBlockIds carries the
              // paragraphs beyond the first): per-paragraph intents each
              // minted a fresh numbering definition at apply, so the server
              // numbered a multi-paragraph toggle 1,2,3 (three restarting
              // lists) while this client's single mutation above shared one
              // definition (1,1,1) — byte divergence and wrong semantics.
              emitListIntent(targets, (blockIds) =>
                operationBody("setListType", blockIds[0], {
                  listKind, suggest,
                  ...(blockIds.length > 1 ? { moreBlockIds: blockIds.slice(1) } : {}),
                }));
              pages = rerender(doc, dirtyBlock);
              document.dispatchEvent(new CustomEvent("dxw-selection"));
            }
          },
          getListType: () => {
            const segs = editor?.getSelectionSegments() ?? [];
            const t = segs.find((sg) => sg.t)?.t ?? editor?.getCaretTarget()?.t;
            return t ? listTypeAt(doc, t) : null;
          },
          setParagraphStyle: (styleId) => {
            const caret = editor?.getCaretTarget();
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : caret ? [caret.t] : [];
            if (targets.length === 0) return;
            const suggest = editor?.suggestionMeta();
            if (collabBlockOp(targets, (blockId) => ({ kind: "formatParagraph", blockId, styleId, ...(suggest ? { suggest } : {}) }))) return;
            history.checkpoint();
            if (setParagraphStyle(doc, targets as Parameters<typeof setParagraphStyle>[1], styleId, suggestMeta(doc, suggest))) {
              pages = rerender(doc);
            }
          },
          listParagraphStyles: () => {
            const out = new Map<string, string>();
            // Word built-ins are always offered; applying one injects its
            // standard definition if the file lacks it.
            for (let n = 1; n <= 6; n++) out.set(`Heading${n}`, `Heading ${n}`);
            out.set("Title", "Title");
            for (const st of doc.styles.byId.values()) {
              if (st.type !== "paragraph" || !st.name) continue;
              if (/^(normal|title|subtitle|heading \d)$/i.test(st.name)) {
                out.set(st.id, st.name);
              }
            }
            const list = [...out.entries()].map(([id, name]) => ({ id, name }));
            list.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
            return list;
          },
          listStyles: (filter) => listStyles(doc, filter),
          createStyle: (spec) => {
            if (collabDocOp(() => documentOperationBody("createStyle", { style: spec }))) return true;
            history.checkpoint();
            if (!createStyle(doc, spec)) return false;
            // A definition change repaints every paragraph that resolves
            // through it, so the whole document relayouts rather than a block.
            pages = rerender(doc, undefined, "global");
            return true;
          },
          modifyStyle: (styleId, patch) => {
            if (collabDocOp(() => documentOperationBody("modifyStyle", { styleId, patch }))) return true;
            history.checkpoint();
            if (!modifyStyle(doc, styleId, patch)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          deleteStyle: (styleId) => {
            if (collabDocOp(() => documentOperationBody("deleteStyle", { styleId }))) return true;
            history.checkpoint();
            if (!deleteStyle(doc, styleId)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertWatermark: (spec) => {
            // The watermark goes into the header parts, and insertWatermark
            // will not create one: making a PART (rel + content-type override
            // + sectPr references) is ensureHeaderFooter's structural intent.
            // So in a room the two ride the wire in order — the header exists
            // by the time the watermark operation applies on every replica.
            if (!doc.hasHfPart("header")) {
              const created = collabDocOp((ids) => ({
                kind: "ensureHeaderFooter",
                hfKind: "header",
                nodeIds: ids(8),
              }));
              if (created) {
                // The local replica applies its own submission, so the part is
                // there; ask again with the real count.
                doc.ensureHfPart("header");
              } else {
                history.checkpoint();
                doc.ensureHfPart("header");
              }
            }
            const headerCount = doc.headerRoots().length;
            if (headerCount === 0) return false;
            const args = {
              text: spec.text,
              headerCount,
              ...(spec.diagonal !== undefined ? { diagonal: spec.diagonal } : {}),
              ...(spec.color !== undefined ? { color: spec.color } : {}),
              ...(spec.opacity !== undefined ? { opacity: spec.opacity } : {}),
            };
            if (collabDocOp((ids) => documentOperationBody("insertWatermark", args, ids))) return true;
            history.checkpoint();
            if (!insertWatermark(doc, spec)) return false;
            // A watermark paints behind every page, so nothing reflows — but
            // it lives in the header, which the page chrome repaints globally.
            pages = rerender(doc, undefined, "global");
            return true;
          },
          removeWatermark: () => {
            if (collabDocOp(() => documentOperationBody("removeWatermark", {}))) return true;
            history.checkpoint();
            if (!removeWatermark(doc)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          setCharacterStyle: (styleId) => {
            // A character style is a run property, so it rides the same path
            // bold and colour do — including the run splitting a partial
            // selection needs, and the collab intents applyFormat emits.
            api.applyFormat({ characterStyleId: styleId });
          },
          setNumberingLevel: (ilvl, patch) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return false;
            if (collabOp((anchor) => operationBody("setNumberingLevel", anchor.blockId, { ilvl, patch }), {
              t: caret.t,
              offset: 0,
            })) {
              return true;
            }
            history.checkpoint();
            if (!setNumberingLevelAt(doc, caret.t, ilvl, patch)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          applyNumberingPreset: (preset) => {
            const spec = NUMBERING_PRESETS[preset];
            if (!spec) return false;
            if (api.getListType() !== "number") api.toggleList("number");
            let applied = true;
            for (let ilvl = 0; ilvl < spec.levels.length; ilvl++) {
              applied = api.setNumberingLevel(ilvl, spec.levels[ilvl]) && applied;
            }
            return applied;
          },
          setNumberingRestart: (start) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return false;
            if (collabOp((anchor) => operationBody("setNumberingRestart", anchor.blockId, { start }), {
              t: caret.t,
              offset: 0,
            })) {
              return true;
            }
            history.checkpoint();
            const applied =
              start === null ? continueNumberingAt(doc, caret.t) : restartNumberingAt(doc, caret.t, start);
            if (!applied) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          copyFormatting: () => api.getSelectionFormat(),
          applyCopiedFormatting: (format) => {
            api.applyFormat(formatPatchFrom(format));
          },
          getParagraphStyleId: () => {
            const segs = editor?.getSelectionSegments() ?? [];
            const t = segs.find((sg) => sg.t)?.t ?? editor?.getCaretTarget()?.t;
            return t ? paragraphStyleIdOf(doc, t) : null;
          },
          print: () => {
            if (!handle) return;
            const sp = doc.sections[0]?.props;
            const restore = handle.materializeAll?.();
            printPages(handle.root, sp?.pageWidth ?? 816, sp?.pageHeight ?? 1056);
            restore?.();
          },
          exportPrintHtml: () => {
            if (!handle) return null;
            const sp = doc.sections[0]?.props;
            const restore = handle.materializeAll?.();
            const html = buildPrintHtml(
              handle.root,
              sp?.pageWidth ?? 816,
              sp?.pageHeight ?? 1056,
            );
            restore?.();
            return html;
          },
          save: () => doc.save(),
          setSuggesting: (on, author) => editor?.setSuggesting(on, author ?? commentAuthor),
          isSuggesting: () => editor?.isSuggesting() ?? false,
          // The editor re-renders through host.rerender (which updates pages).
          acceptRevisionAtCaret: () => editor?.acceptRevisionRef() ?? false,
          rejectRevisionAtCaret: () => editor?.rejectRevisionRef() ?? false,
          // Count works in any mode (read-only walk); the bulk operations
          // need the editor (they re-render + record an undo step).
          revisionCount: () => collectRevisions(doc).length,
          acceptAllRevisions: () => editor?.acceptAllRevisions() ?? 0,
          rejectAllRevisions: () => editor?.rejectAllRevisions() ?? 0,
          getEncodedCaret: () => {
            doc.enableStableIds(); // local documents encode too (see the API doc)
            return editor?.getEncodedCaret() ?? null;
          },
          setCaretFromEncoded: (pos) => editor?.setCaretFromEncoded(pos) ?? false,
          revealPresence: (participant) => {
            // Read the CURRENT presence (presenceRef), not a render's snapshot:
            // the person may have moved — or left — since the caller rendered.
            const pos = presenceRef.current?.[participant];
            if (!pos) return "no-position";
            const ids = doc.stableIds;
            if (!ids || !handle) return "unresolved";
            // decodeCaret resolves the wire-basis anchor to a concrete w:t and
            // falls back to the block's first text when the exact run is gone.
            const hit = ids.decodeCaret(pos.anchor);
            if (!hit) return "unresolved";
            // A far caret lives on an unmounted (virtualized) page: mount all
            // pages, scroll, restore the window — find navigation's pattern
            // (selectMatch above). Mounting renders DOM from the RETAINED
            // layout; no layout work happens.
            const restore = handle._virtualized ? handle.materializeAll?.() : undefined;
            const bindings = handle.bindingsByText.get(hit.t);
            // The wrap segment containing the offset, like the presence caret
            // draw — bindings[0] could be a different line (or page) for a
            // long run.
            const b =
              bindings?.find((bd) => {
                const s = bd.item.src?.offset ?? 0;
                return hit.offset >= s && hit.offset <= s + bd.item.text.length;
              }) ?? bindings?.[bindings.length - 1];
            if (!b) {
              restore?.();
              return "unresolved";
            }
            b.el.scrollIntoView({ block: "center", behavior: "instant" as ScrollBehavior });
            restore?.();
            handle.updateViewport?.();
            return "revealed";
          },
        };
        apiRef.current = api;
        onReady?.(api);
      }
      // Live zoom updates: re-paint pages at the new scale in place. Reads doc
      // and curZoom from this closure so it survives across zoom changes without
      // re-running the effect (which would reparse and reset editor state).
      applyZoomRef.current = (z: number) => {
        if (z === curZoom) return;
        curZoom = z;
        if (editorConfig) editorConfig.zoom = z;
        rerender(doc);
      };
      // Apply any zoom the ResizeObserver settled on while we were loading.
      if (effZoomRef.current !== curZoom) applyZoomRef.current(effZoomRef.current);
    })().catch((e: unknown) => {
      if (cancelled) return;
      const err = e instanceof Error ? e : new Error(String(e));
      setError(err);
      onError?.(err);
    });

    return () => {
      cancelled = true;
      layoutJob++;
      layoutAbort?.abort();
      layoutAbort = null;
      if (layoutTimer) clearTimeout(layoutTimer);
      layoutTimer = null;
      const container = containerRef.current;
      if (container) {
        container.inert = false;
        container.removeAttribute("data-dxw-layout-busy");
        container.setAttribute("aria-busy", "false");
      }
      setLayoutBusy(false);
      applyZoomRef.current = null;
      detachPresenceSender?.();
      detachPresenceSender = null;
      editor?.detach();
      editor = null;
      handle?.destroy();
      handle = null;
    };
  }, [source, editable, commentAuthor, showComments, revisions]);

  // Live-collab in-place repaint. When the replica applies a broadcast it
  // mutates the SAME doc object we already rendered (collab.doc) and bumps
  // renderSignal; repaint that instance directly rather than re-running the
  // load effect. No bytes, no parse, no editor re-attach — the remote edit
  // shows up as a single relayout of the live tree.
  //
  // Coalesced through rAF: a typing burst raises one signal per submit AND one
  // per echo (2× per keystroke), and painting each synchronously both wasted
  // work and could interleave so a mid-burst paint landed LAST, leaving stale
  // text on screen until the next event. One deferred repaint always runs
  // after the latest mutation, reading the doc's current state.
  const renderSignal = collab?.renderSignal;
  const repaintScheduledRef = useRef(false);
  const repaintRafRef = useRef(0);
  const repaintTimerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  useEffect(() => {
    if (renderSignal === undefined) return;
    if (repaintScheduledRef.current) return; // a repaint is already scheduled
    repaintScheduledRef.current = true;
    const paint = () => {
      if (!repaintScheduledRef.current) return; // the other scheduler already ran
      repaintScheduledRef.current = false;
      if (repaintRafRef.current) { cancelAnimationFrame(repaintRafRef.current); repaintRafRef.current = 0; }
      if (repaintTimerRef.current !== undefined) { clearTimeout(repaintTimerRef.current); repaintTimerRef.current = undefined; }
      const d = docCacheRef.current?.doc;
      const r = rerenderRef.current;
      // Drain the dirty scope only when actually painting — an aborted paint
      // leaves it accumulated for the paint that does run. A null take means
      // nothing is dirty (an earlier paint covered every recorded change, e.g.
      // the mount paint racing the first renderSignal) — skip: repainting
      // would queue a redundant whole-document relayout.
      if (d && r && d === collab?.doc) {
        const take = collab.takeRenderScope;
        const scope = take ? take() : undefined;
        if (scope !== null) {
          r(d, scope);
          // The toolbar reads api state (revision count, suggesting, formats)
          // on its dxw-selection refresh; a session-applied op changes that
          // state without moving the caret, so announce it here — once per
          // coalesced repaint.
          document.dispatchEvent(new CustomEvent("dxw-selection"));
        }
      }
    };
    // rAF coalesces to vsync when THIS window is the foreground one — but the
    // browser PAUSES rAF in a hidden tab and THROTTLES it in a visible-but-
    // unfocused window (e.g. two docs side by side). Relying on rAF alone froze
    // a remote collaborator's view — and their remote carets — on every edit
    // until the window regained focus (the on-screen divergence users hit; the
    // session data had already converged). A timer fallback guarantees the
    // repaint still runs off-foreground; whichever of the two fires first does
    // the (idempotent) paint and cancels the other.
    repaintRafRef.current = requestAnimationFrame(paint);
    repaintTimerRef.current = setTimeout(paint, 60);
    // Intentionally keyed only on renderSignal: collab.doc is stable between
    // reloads, and a reload re-runs the main effect above instead.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [renderSignal]);
  useEffect(() => () => {
    if (repaintRafRef.current) cancelAnimationFrame(repaintRafRef.current);
    if (repaintTimerRef.current !== undefined) clearTimeout(repaintTimerRef.current);
  }, []);

  /**
   * Step the mail-merge preview to another record.
   *
   * Keyed on the record's CONTENT: a host rebuilds the record object on every
   * render, and keying on identity would relay the whole document on every
   * keystroke. The mount paint already used the first record, so the ref starts
   * equal and the first run of this effect does nothing.
   *
   * "doc" scope forces a full relayout. The engine gates reuse on the record
   * too (LayoutOptions.mergeRecord), because stepping records changes no
   * blocks — so nothing in the document can tell the incremental path that the
   * painted text is now stale.
   */
  const mergeKey = mergeRecord ? JSON.stringify(mergeRecord) : "";
  const paintedMergeKeyRef = useRef(mergeKey);
  useEffect(() => {
    if (paintedMergeKeyRef.current === mergeKey) return;
    paintedMergeKeyRef.current = mergeKey;
    const doc = docCacheRef.current?.doc;
    const rerender = rerenderRef.current;
    if (doc && rerender) rerender(doc, { kind: "doc" });
  }, [mergeKey]);

  const hotBtn = (label: string, title: string, onClick: () => void) => (
    <button
      title={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      style={{
        border: "1px solid #dadce0",
        background: "#fff",
        color: "#3c4043",
        font: "12.5px system-ui, sans-serif",
        padding: "4px 10px",
        borderRadius: 14,
        cursor: "pointer",
        boxShadow: "0 1px 3px rgba(0,0,0,.12)",
      }}
    >
      {label}
    </button>
  );
  return (
    <div
      style={{
        position: "relative",
        display: "flex",
        flexDirection: "column",
        minHeight: 0,
        ...(style?.height ? { height: style.height } : {}),
      }}
    >
      {editable && hfMode && (
        <div
          data-dxw-editor-context-row=""
          style={{
            flex: "0 0 52px",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            background: "var(--dxw-canvas-bg, #e8eaed)",
            position: "relative",
            zIndex: 40,
            pointerEvents: "none",
          }}
        >
          <div
            data-dxw-hf-hotbar=""
            style={{
              display: "flex",
              gap: 6,
              background: "rgba(249,251,253,.96)",
              border: "1px solid #dadce0",
              borderRadius: 18,
              padding: "5px 8px",
              boxShadow: "0 2px 10px rgba(0,0,0,.15)",
              alignItems: "center",
              pointerEvents: "auto",
            }}
          >
            <span style={{ font: "600 11.5px system-ui, sans-serif", color: "#5f6368", padding: "0 4px" }}>
              Header &amp; footer · repeats on pages
            </span>
            {hotBtn("Page number", "Insert a dynamic page number at the caret", () => apiRef.current?.insertPageNumber("page"))}
            {hotBtn("Page X of Y", "Insert 'Page X of Y' at the caret", () => apiRef.current?.insertPageNumber("pageOfTotal"))}
            {hotBtn("Close", "Return to the document body", () => {
              apiRef.current?.closeHeaderFooter();
              setHfMode(false);
            })}
          </div>
        </div>
      )}
      <div
        ref={containerRef}
        className={className}
        style={{
          background: "var(--dxw-canvas-bg, #e8eaed)",
          overflow: "auto",
          height: "100%",
          ...style,
          ...(editable ? {
            height: "auto",
            flex: "1 1 auto",
            minHeight: 0,
            boxSizing: "border-box" as const,
          } : {}),
        }}
      >
        {error && (
          <div style={{ padding: 16, color: "#b00020", fontFamily: "system-ui" }}>
            Failed to render document: {error.message}
          </div>
        )}
      </div>
      {layoutBusy && (
        <div
          data-dxw-layout-status=""
          aria-live="polite"
          style={{
            position: "absolute",
            top: editable ? 62 : 10,
            right: 12,
            zIndex: 45,
            padding: "5px 10px",
            borderRadius: 14,
            background: "rgba(32,33,36,.88)",
            color: "#fff",
            font: "12px system-ui,sans-serif",
            pointerEvents: "none",
          }}
        >
          Repaginating…
        </div>
      )}
    </div>
  );
}

export { DocxDocument, layoutDocument, renderToDom, printPages } from "@wordinweb/core";
export type { CoverPageContent, DrawingTool, MergeRecord, RunFormatPatch, SelectionFormat, ParagraphAlignment, PageLayoutPatch, LineNumberingPatch, ShapePreset, WireRange, EncodedCaret, HostShortcutSection } from "@wordinweb/core";
export { DocxToolbar, ToolbarMenuSelect, INSERT_COMMANDS } from "./toolbar.js";
export type {
  DocxToolbarProps,
  InsertCommandSpec,
  ToolbarFeature,
  ToolbarMenuSelectOption,
  ToolbarMenuSelectProps,
  ToolbarMode,
} from "./toolbar.js";
