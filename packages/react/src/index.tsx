import { useCallback, useEffect, useRef, useState } from "react";
import {
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
  addComment,
  adjustIndent,
  deleteComment,
  findAll,
  linkAt,
  removeLink,
  replaceMatch,
  replaceAll,
  setLink,
  setParagraphSpacing,
  setDropCapAt,
  transformCase,
  exactLineHeightAt,
  replyToComment,
  insertImageAt,
  setImageWrap,
  insertFootnote,
  insertField,
  insertDateTimeField,
  insertPageField,
  listBookmarks,
  insertBookmarkAroundSelection,
  insertBookmarkAt,
  insertCrossReference,
  insertMathAt,
  insertShapeAt,
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
  setLineNumbering,
  lineNumberingAt,
  type XmlElement,
  type ShapePreset,
  type WordArtPreset,
  type DrawingTool,
  type CoverPageContent,
  type ObjectArrangeAction,
  type ChartData,
  type SmartArtData,
  insertTableAfter,
  createMeasurer,
  type TextMeasurer,
  layoutDocument,
  layoutDocumentAsync,
  relayoutHeadersFooters,
  detectMissingFonts,
  type MissingFont,
  type LayoutResult,
  listTypeAt,
  printPages,
  setListType,
  paragraphStyleIdOf,
  renderToDom,
  selectionToSegments,
  setPageLayout,
  setParagraphAlignment,
  setParagraphStyle,
  summarizeSelection,
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
  /** Insert an editable inline equation from WordInWeb's linear math syntax. */
  insertEquation(linear: string): boolean;
  /** Insert a Unicode symbol through the normal undo/suggestion-aware typing path. */
  insertSymbol(symbol: string): boolean;
  /** Insert a floating editable DrawingML shape at the caret. */
  insertShape(preset: ShapePreset, text?: string): boolean;
  /** Insert editable DrawingML WordArt at the caret. */
  insertWordArt(text: string, preset?: WordArtPreset): boolean;
  /** Insert a native editable ChartML chart at the caret. */
  insertChart(data: ChartData): boolean;
  /** Replace the selected native chart's type and data. */
  updateSelectedChart(data: ChartData): boolean;
  /** Insert a native editable SmartArt diagram at the caret. */
  insertSmartArt(data: SmartArtData): boolean;
  /** Replace the selected native SmartArt diagram's layout and node text. */
  updateSelectedSmartArt(data: SmartArtData): boolean;
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
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  canRedo(): boolean;
  /** Insert a rows×cols table at the caret's paragraph. */
  insertTable(rows: number, cols: number): void;
  /** Row/column/table operations on the table containing the caret. */
  tableOp(op: TableOp): void;
  /** Insert an image file at the caret (inline, natural size clamped to column). */
  insertImage(file: Blob): Promise<void>;
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
  /** Line spacing multiple and/or space before/after (points). */
  setParagraphSpacing(patch: { lineMultiple?: number; beforePt?: number | null; afterPt?: number | null }): void;
  /** Apply or remove a native Word drop cap on the caret paragraph. */
  setDropCap(mode: "drop" | "margin" | null, lines?: number): boolean;
  /** Remove direct character formatting from the selection. */
  clearFormatting(): void;
  /** Change the selection's case. */
  changeCase(mode: "upper" | "lower" | "title"): void;
  /** Find matches for a query; selects the first and returns the count. */
  find(query: string, opts?: { matchCase?: boolean }): number;
  /** Select the next/previous match; returns 1-based index or 0. */
  findStep(delta: 1 | -1): number;
  /** Replace the current match; returns remaining match count. */
  replaceCurrent(replacement: string): number;
  /** Replace every match; returns how many were replaced. */
  replaceAll(query: string, replacement: string): number;
  /** Paragraph styles for the style menu (declared + Word built-ins). */
  listParagraphStyles(): { id: string; name: string }[];
  /** pStyle id of the caret paragraph (null = Normal). */
  getParagraphStyleId(): string | null;
  /** Change margins / page size / orientation (inches). */
  setPageLayout(patch: PageLayoutPatch, scope?: "document" | "section"): void;
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
  /** Effective formatting of the current selection (toolbar state), or null. */
  getSelectionFormat(): SelectionFormat | null;
  /** Print the rendered pages (browser print dialog / save as PDF). */
  print(): void;
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
  document: DocxDocument;
}

export type ScreenshotInsertResult = "inserted" | "unsupported" | "cancelled" | "error" | "no-caret";

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
  /** Fires when the document is ready; the api is only usable while mounted. */
  onReady?: (api: DocxViewApi) => void;
  onError?: (error: Error) => void;
  /** Author name stamped on comment replies (default "You"). */
  commentAuthor?: string;
  /** Render review comments (range highlights + margin balloons). Default true. */
  showComments?: boolean;
  /** Tracked-changes display: "final" (default) or "markup". */
  revisions?: "final" | "markup";
  /** Fires after render with document-requested font faces the browser cannot
   * render (unavailable, or lacking the document's script) — the page is
   * silently substituting and may differ from Word. Empty array = all good. */
  onMissingFonts?: (missing: MissingFont[]) => void;
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
  onReady,
  onError,
  commentAuthor = "You",
  showComments = true,
  revisions = "final",
  onMissingFonts,
}: DocxViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
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
    let onDeleteComment: ((id: string) => void) | undefined;
    let onReplyComment: ((id: string, text: string) => void) | undefined;
    let applyStyleShortcut: ((styleId: string | null) => void) | undefined;
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
      // Re-rendering replaces the page DOM; keep the user's scroll position
      // (destroy-then-append clamps scrollTop to 0 otherwise). The previous
      // handle is handed to renderToDom so it can adopt the DOM of unchanged
      // pages and tear down only what actually changed.
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
        onViewportChange: () => editor?.afterRender(),
      }, prev ?? undefined);
      container.scrollTop = scrollTop;
      container.scrollLeft = scrollLeft;
      const t3 = perf ? performance.now() : 0;
      editor?.afterRender();
      if (perf) {
        perf.last = {
          layout: layoutMs,
          destroy: 0,
          render: t3 - t2,
          totalPages: layout.totalPages,
        };
      }
      pages = layout.totalPages;
      return layout.totalPages;
    };

    const queueGlobalLayout = (doc: DocxDocument, delayMs = 0, preferHeadersOnly = false): void => {
      layoutAbort?.abort();
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
          const fast = relayoutHeadersFooters(doc, prevLayout, measurer);
          if (fast) {
            if (doc.modelVersion === modelVersion) paintLayout(doc, fast, performance.now() - started);
            layoutAbort = null;
            setLayoutPending(false);
            return;
          }
        }
        setLayoutPending(true);
        void layoutDocumentAsync(doc, { measurer, signal: abort.signal }).then((layout) => {
          if (cancelled || job !== layoutJob || abort.signal.aborted || doc.modelVersion !== modelVersion) return;
          paintLayout(doc, layout, performance.now() - started);
        }).catch((cause: unknown) => {
          if (abort.signal.aborted || cancelled) return;
          const err = cause instanceof Error ? cause : new Error(String(cause));
          setError(err);
          onError?.(err);
        }).finally(() => {
          if (job !== layoutJob) return;
          layoutAbort = null;
          setLayoutPending(false);
        });
      };
      if (delayMs > 0) {
        setLayoutPending(false);
        layoutTimer = setTimeout(start, delayMs);
      } else {
        start();
      }
    };

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
      setLayoutPending(false);
      const started = performance.now();
      const layout =
        headerFooterOnly && prevLayout
          ? relayoutHeadersFooters(doc, prevLayout, measurer) ?? layoutDocument(doc, { measurer })
          : layoutDocument(doc, {
              measurer,
              prev: globalChange ? undefined : prevLayout ?? undefined,
              dirtyHint: globalChange ? undefined : dirtyBlock,
              dirtySource: globalChange ? undefined : dirtySource,
            });
      return paintLayout(doc, layout, performance.now() - started);
    };

    (async () => {
      const cached = docCacheRef.current?.source === source ? docCacheRef.current : null;
      const bytes = cached ? null : await toBytes(source);
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
            // Complex-script faces gate layout too (advances must be real
            // before line breaking): Mangal (Devanagari), Latha (Tamil),
            // Noto Sans Lao Looped (Lao — bundled OFL webfont).
            "Mangal", "Latha", "Noto Sans Lao Looped",
          ];
          for (const fam of latin) {
            for (const variant of ["", "italic ", "bold ", "bold italic "]) {
              loads.push(document.fonts.load(`${variant}16px "${fam}"`).catch(() => []));
            }
          }
          loads.push(document.fonts.load('16px "Cambria Math"').catch(() => []));
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
              loads.push(document.fonts.load(`${variant}16px "${fam}"`).catch(() => []));
            }
          }
          await Promise.all(loads);
          await document.fonts.ready;
        } catch {
          /* non-fatal */
        }
      }
      if (cancelled) return;
      const doc = cached ? cached.doc : DocxDocument.load(bytes!);
      if (!cached) docCacheRef.current = { source, doc, history: new EditHistory(doc) };
      const wantView = revisions === "markup" ? "markup" : "final";
      if (doc.revisionView !== wantView) doc.setRevisionView(wantView);
      // Feed the real page width into fit-to-width now that we know it, then
      // recompute (the first ResizeObserver pass ran against the 816px default).
      basePageWidthRef.current = doc.sections[0]?.props.pageWidth ?? 816;
      curZoom = effZoomRef.current;
      const pageCount = rerender(doc);
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
            const formatted = applyRunFormat(doc, segs, patch);
            pages = rerender(doc);
            if (selectedAll) editor?.selectAll();
            else if (formatted.length > 0) editor?.selectRanges(formatted);
            document.dispatchEvent(new CustomEvent("dxw-selection"));
          },
          onStyleShortcut: (styleId) => applyStyleShortcut?.(styleId),
        };
        editor = new DocxEditor(editorConfig);
        editor.attach();
        applyStyleShortcut = (styleId) => {
          const caret = editor?.getCaretTarget();
          const segs = editor?.getSelectionSegments() ?? [];
          const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : caret ? [caret.t] : [];
          if (targets.length === 0) return;
          history.checkpoint();
          if (setParagraphStyle(doc, targets as Parameters<typeof setParagraphStyle>[1], styleId)) {
            pages = rerender(doc);
            document.dispatchEvent(new CustomEvent("dxw-selection"));
          }
        };
        onDeleteComment = (id) => {
          history.checkpoint();
          if (deleteComment(doc, id)) pages = rerender(doc);
        };
        onReplyComment = (id, text) => {
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
            const formatted = applyRunFormat(doc, segments, patch);
            pages = rerender(doc);
            // Keep the formatted text selected so toolbar actions compose.
            if (selectedAll) editor?.selectAll();
            else if (formatted.length > 0) editor?.selectRanges(formatted);
          },
          addFootnote: (text) => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (insertFootnote(doc, target.t, target.offset, text) !== null) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
          insertPageNumber: (kind = "page") => {
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
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertField(doc, target.t, target.offset, instruction, cachedResult)) return false;
            pages = rerender(doc);
            return true;
          },
          insertDateTime: (kind, picture) => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            const format = picture ?? (kind === "time" ? "h:mm am/pm" : "M/d/yyyy");
            if (!insertDateTimeField(doc, target.t, target.offset, kind, format)) return false;
            pages = rerender(doc);
            return true;
          },
          listBookmarks: () => listBookmarks(doc),
          addBookmark: (name) => {
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
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertCrossReference(doc, target.t, target.offset, bookmark, kind)) return false;
            pages = rerender(doc);
            return true;
          },
          insertEquation: (linear) => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertMathAt(doc, target.t, target.offset, linear)) return false;
            pages = rerender(doc);
            return true;
          },
          insertSymbol: (symbol) => editor?.insertText(symbol) ?? false,
          insertShape: (preset, text) => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertShapeAt(doc, target.t, preset, text)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertWordArt: (text, preset = "plain") => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertWordArtAt(doc, target.t, text, preset)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          insertChart: (data) => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertChartAt(doc, target.t, data)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          updateSelectedChart: (data) => {
            const source = editor?.getSelectedDrawingSource();
            if (!source) return false;
            history.checkpoint();
            if (!setChartData(doc, source, data)) return false;
            pages = rerender(doc, undefined, "global");
            editor?.reselectDrawing(source);
            return true;
          },
          insertSmartArt: (data) => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertSmartArtAt(doc, target.t, data)) return false;
            pages = rerender(doc, undefined, "global");
            return true;
          },
          updateSelectedSmartArt: (data) => {
            const source = editor?.getSelectedDrawingSource();
            if (!source) return false;
            history.checkpoint();
            if (!setSmartArtData(doc, source, data)) return false;
            pages = rerender(doc, undefined, "global");
            editor?.reselectDrawing(source);
            return true;
          },
          insertModel3D: async (file, poster) => {
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
          addComment: (text) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const segments = segs.length > 0 ? segs : handle ? selectionToSegments(handle.bindings) : [];
            if (segments.length === 0) return false;
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
            const caret = editor?.getCaretTarget();
            if (!caret) return;
            history.checkpoint();
            if (insertTableAfter(doc, caret.t, rows, cols)) pages = rerender(doc);
          },
          tableOp: (op) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return;
            history.checkpoint();
            if (applyTableOp(doc, caret.t, op)) pages = rerender(doc);
          },
          insertImage: async (file) => {
            const caret = editor?.getCaretTarget();
            if (!caret) return;
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
              await api.insertImage(image);
              return "inserted";
            } catch (error) {
              return error instanceof DOMException && (error.name === "NotAllowedError" || error.name === "AbortError")
                ? "cancelled"
                : "error";
            } finally {
              for (const track of stream?.getTracks() ?? []) track.stop();
            }
          },
          setAlignment: (align) => {
            if (!handle) return;
            const caret = editor?.getCaretTarget();
            const segTs = selectionToSegments(handle.bindings)
              .map((s) => s.t)
              .filter((t): t is NonNullable<typeof t> => !!t);
            const targets = segTs.length > 0 ? segTs : caret ? [caret.t] : [];
            if (targets.length === 0) return;
            history.checkpoint();
            if (setParagraphAlignment(doc, targets as Parameters<typeof setParagraphAlignment>[1], align)) {
              pages = rerender(doc);
            }
          },
          closeHeaderFooter: () => editor?.exitHeaderFooter(),
          openHeaderFooter: (kind) => editor?.enterHeaderFooter(kind) ?? false,
          insertBreak: (kind) => {
            let target = editor?.getCaretTarget() ?? null;
            if (!target) {
              const segs = editor?.getSelectionSegments() ?? [];
              const last = [...segs].reverse().find((sg) => sg.t);
              if (last?.t) target = { t: last.t, offset: last.end };
            }
            if (!target) return false;
            history.checkpoint();
            const done =
              kind === "page" || kind === "column"
                ? insertBreakAt(doc, target.t, target.offset, kind)
                : insertSectionBreak(doc, target.t, kind === "sectionNextPage" ? "nextPage" : "continuous");
            if (done) pages = rerender(doc);
            return done;
          },
          insertBlankPage: () => {
            const target = insertionTarget();
            if (!target) return false;
            history.checkpoint();
            if (!insertBlankPageAt(doc, target.t, target.offset)) return false;
            pages = rerender(doc);
            return true;
          },
          insertCoverPage: (content) => {
            history.checkpoint();
            if (!insertCoverPage(doc, content)) return false;
            pages = rerender(doc);
            return true;
          },
          setPageLayout: (patch, scope) => {
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (setPageLayout(doc, patch, target)) pages = rerender(doc);
          },
          setLineNumbering: (patch, scope) => {
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
            history.checkpoint();
            if (adjustIndent(doc, targets as Parameters<typeof adjustIndent>[1], direction)) pages = rerender(doc);
          },
          setParagraphSpacing: (patch) => {
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : editor?.getCaretTarget() ? [editor.getCaretTarget()!.t] : [];
            if (targets.length === 0) return;
            history.checkpoint();
            if (setParagraphSpacing(doc, targets as Parameters<typeof setParagraphSpacing>[1], patch)) pages = rerender(doc);
          },
          setDropCap: (mode, lines = 3) => {
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
            history.checkpoint();
            replaceMatch(doc, m, replacement);
            pages = rerender(doc);
            findState.matches.splice(findState.index, 1);
            if (findState.index >= findState.matches.length) findState.index = 0;
            if (findState.matches.length > 0) selectMatch(findState.index);
            return findState.matches.length;
          },
          replaceAll: (query, replacement) => {
            history.checkpoint();
            const n = replaceAll(doc, query, replacement);
            if (n > 0) pages = rerender(doc);
            findState = { matches: [], index: 0 };
            return n;
          },
          toggleList: (kind) => {
            const caret = editor?.getCaretTarget();
            const segs = editor?.getSelectionSegments() ?? [];
            const targets = segs.length > 0 ? segs.map((sg) => sg.t).filter((t): t is NonNullable<typeof t> => !!t) : caret ? [caret.t] : [];
            if (targets.length === 0) return;
            const current = listTypeAt(doc, targets[0]);
            history.checkpoint();
            if (setListType(doc, targets as Parameters<typeof setListType>[1], current === kind ? null : kind)) {
              pages = rerender(doc);
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
            history.checkpoint();
            if (setParagraphStyle(doc, targets as Parameters<typeof setParagraphStyle>[1], styleId)) {
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
      editor?.detach();
      editor = null;
      handle?.destroy();
      handle = null;
    };
  }, [source, editable, commentAuthor, showComments, revisions]);

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
    <div style={{ position: "relative", ...(style?.height ? { height: style.height } : {}) }}>
      <div
        ref={containerRef}
        className={className}
        style={{ background: "var(--dxw-canvas-bg, #e8eaed)", overflow: "auto", height: "100%", ...style }}
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
            top: 10,
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
      {editable && hfMode && (
        <div
          data-dxw-hf-hotbar=""
          style={{
            position: "absolute",
            top: 10,
            left: "50%",
            transform: "translateX(-50%)",
            display: "flex",
            gap: 6,
            zIndex: 40,
            background: "rgba(249,251,253,.96)",
            border: "1px solid #dadce0",
            borderRadius: 18,
            padding: "5px 8px",
            boxShadow: "0 2px 10px rgba(0,0,0,.15)",
            alignItems: "center",
          }}
        >
          <span style={{ font: "600 11.5px system-ui, sans-serif", color: "#5f6368", padding: "0 4px" }}>
            Header &amp; footer
          </span>
          {hotBtn("Page number", "Insert a dynamic page number at the caret", () => apiRef.current?.insertPageNumber("page"))}
          {hotBtn("Page X of Y", "Insert 'Page X of Y' at the caret", () => apiRef.current?.insertPageNumber("pageOfTotal"))}
          {hotBtn("Close", "Return to the document body", () => {
            apiRef.current?.closeHeaderFooter();
            setHfMode(false);
          })}
        </div>
      )}
    </div>
  );
}

export { DocxDocument, layoutDocument, renderToDom, printPages } from "@wordinweb/core";
export type { CoverPageContent, DrawingTool, RunFormatPatch, SelectionFormat, ParagraphAlignment, PageLayoutPatch, LineNumberingPatch, ShapePreset } from "@wordinweb/core";
export { DocxToolbar } from "./toolbar.js";
export type { DocxToolbarProps, ToolbarFeature, ToolbarMode } from "./toolbar.js";
