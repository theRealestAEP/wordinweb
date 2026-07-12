import { useEffect, useRef, useState } from "react";
import {
  DocxDocument,
  DocxEditor,
  EditHistory,
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
  transformCase,
  exactLineHeightAt,
  replyToComment,
  insertImageAt,
  setImageWrap,
  insertFootnote,
  insertPageField,
  insertBreakAt,
  insertSectionBreak,
  sectPrAt,
  type XmlElement,
  insertTableAfter,
  createMeasurer,
  type TextMeasurer,
  layoutDocument,
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
} from "@docxinweb/core";

export interface DocxViewApi {
  /** Apply character formatting to the current browser selection. */
  applyFormat(patch: RunFormatPatch): void;
  /** Create a review comment on the current selection. False if no selection. */
  addComment(text: string): boolean;
  /** Insert a footnote at the caret. False without a caret. */
  addFootnote(text: string): boolean;
  /** Insert a dynamic page-number field at the caret (body, header or footer). */
  insertPageNumber(kind?: "page" | "pageOfTotal"): boolean;
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
  /** Leave header/footer editing mode. */
  closeHeaderFooter(): void;
  /** Effective formatting of the current selection (toolbar state), or null. */
  getSelectionFormat(): SelectionFormat | null;
  /** Print the rendered pages (browser print dialog / save as PDF). */
  print(): void;
  /** Serialize the (edited) document back to .docx bytes. */
  save(): Uint8Array;
  /** Page count after the latest layout. */
  pageCount(): number;
  document: DocxDocument;
}

export interface DocxViewProps {
  /** The document: raw bytes, a File/Blob, or a URL to fetch. */
  source: ArrayBuffer | Uint8Array | Blob | string;
  /** Zoom factor, 1 = 100%. */
  zoom?: number;
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
  editable = false,
  className,
  style,
  onLoad,
  onReady,
  onError,
  commentAuthor = "You",
  showComments = true,
  revisions = "final",
}: DocxViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState<Error | null>(null);
  // Contextual header/footer hotbar: the editor announces hf-mode via a
  // bubbled dxw-hfmode event; the tools that only make sense there (page
  // numbers, close) surface right where the user is editing.
  const [hfMode, setHfMode] = useState(false);
  const apiRef = useRef<DocxViewApi | null>(null);
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
    setError(null);

    // One measurer for the document's lifetime: its width/metrics caches survive
    // across keystrokes so unchanged text is not re-measured on every relayout
    // (the default path builds a fresh, cold measurer per layoutDocument call).
    // Cache hits return the exact same values, so layout output is unchanged.
    const measurer: TextMeasurer = createMeasurer();

    const rerender = (doc: DocxDocument): number => {
      const perf = (globalThis as { __dxwPerf?: { last?: Record<string, number> } }).__dxwPerf;
      const t0 = perf ? performance.now() : 0;
      const layout = layoutDocument(doc, { measurer });
      const t1 = perf ? performance.now() : 0;
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
        zoom,
        interactive: editable,
        comments: showComments,
        onDeleteComment,
        onReplyComment,
      }, prev ?? undefined);
      container.scrollTop = scrollTop;
      container.scrollLeft = scrollLeft;
      const t3 = perf ? performance.now() : 0;
      editor?.afterRender();
      if (perf) {
        perf.last = {
          layout: t1 - t0,
          destroy: t2 - t1,
          render: t3 - t2,
          totalPages: layout.totalPages,
        };
      }
      return layout.totalPages;
    };

    (async () => {
      const bytes = await toBytes(source);
      if (cancelled) return;
      if (typeof document !== "undefined" && document.fonts?.ready) {
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
            // Indic faces gate layout too (complex-script advances must be
            // real before line breaking): Mangal (Devanagari), Latha (Tamil).
            "Mangal", "Latha",
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
      const doc = DocxDocument.load(bytes);
      if (revisions !== "final") doc.setRevisionView(revisions);
      const pageCount = rerender(doc);
      let pages = pageCount;
      onLoad?.({ pageCount, document: doc });

      if (editable && containerRef.current) {
        const history = new EditHistory(doc);
        editor = new DocxEditor({
          doc,
          container: containerRef.current,
          getHandle: () => handle,
          rerender: () => {
            pages = rerender(doc);
          },
          zoom,
          history,
          onFormatShortcut: (kind) => {
            const segs = editor?.getSelectionSegments() ?? [];
            if (segs.length === 0) return;
            const fmt = summarizeSelection(segs);
            const patch =
              kind === "bold" ? { bold: !fmt?.bold } :
              kind === "italic" ? { italic: !fmt?.italic } :
              { underline: !fmt?.underline };
            history.checkpoint();
            const formatted = applyRunFormat(doc, segs, patch);
            pages = rerender(doc);
            if (formatted.length > 0) editor?.selectRanges(formatted);
            document.dispatchEvent(new CustomEvent("dxw-selection"));
          },
          onStyleShortcut: (styleId) => applyStyleShortcut?.(styleId),
        });
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
          editor.selectRanges(m.ranges);
          // Bring the hit into view.
          const t = m.ranges[0]?.t;
          const el = handle?.bindings.find((b) => b.item.src?.t === t)?.el;
          el?.scrollIntoView({ block: "center", behavior: "instant" as ScrollBehavior });
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
            history.checkpoint();
            const formatted = applyRunFormat(doc, segments, patch);
            pages = rerender(doc);
            // Keep the formatted text selected so toolbar actions compose.
            if (formatted.length > 0) editor?.selectRanges(formatted);
          },
          addFootnote: (text) => {
            // Caret first; else the end of the current selection.
            let target = editor?.getCaretTarget() ?? null;
            if (!target) {
              const segs = editor?.getSelectionSegments() ?? [];
              const last = [...segs].reverse().find((sg) => sg.t);
              if (last?.t) target = { t: last.t, offset: last.end };
            }
            if (!target) return false;
            history.checkpoint();
            if (insertFootnote(doc, target.t, target.offset, text) !== null) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
          insertPageNumber: (kind = "page") => {
            // Caret first; else the end of the current selection.
            let target = editor?.getCaretTarget() ?? null;
            if (!target) {
              const segs = editor?.getSelectionSegments() ?? [];
              const last = [...segs].reverse().find((sg) => sg.t);
              if (last?.t) target = { t: last.t, offset: last.end };
            }
            if (!target) return false;
            history.checkpoint();
            if (insertPageField(doc, target.t, target.offset, kind)) {
              pages = rerender(doc);
              return true;
            }
            return false;
          },
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
            const bmp = await createImageBitmap(new Blob([bytes.buffer as ArrayBuffer]));
            const sp = doc.sections[0]?.props;
            const maxW = sp ? sp.pageWidth - sp.marginLeft - sp.marginRight : 624;
            const scale = Math.min(1, maxW / bmp.width);
            const ext = (file.type.split("/")[1] ?? "png").replace("jpeg", "jpg");
            history.checkpoint();
            const relId = doc.addImageResource(bytes, ext === "jpg" ? "jpeg" : ext);
            const h = bmp.height * scale;
            const drawing = insertImageAt(doc, caret.t, relId, bmp.width * scale, h);
            if (drawing) {
              // An image taller than an "exact"-spaced line would be clipped
              // (Word) or overlap neighbors — float it with square wrap.
              const exact = exactLineHeightAt(doc, caret.t);
              if (exact !== null && h > exact + 0.5) {
                setImageWrap(doc, drawing, "square", { x: 0, y: 0 });
              }
              pages = rerender(doc);
            }
            bmp.close();
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
          setPageLayout: (patch, scope) => {
            history.checkpoint();
            let target: XmlElement | undefined;
            if (scope === "section") {
              const t = editor?.getCaretTarget()?.t ?? editor?.getSelectionSegments()?.[0]?.t;
              if (t) target = sectPrAt(doc, t) ?? undefined;
            }
            if (setPageLayout(doc, patch, target)) pages = rerender(doc);
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
          clearFormatting: () => {
            api.applyFormat({ clear: true });
          },
          changeCase: (mode) => {
            const segs = editor?.getSelectionSegments() ?? [];
            if (segs.length === 0) return;
            history.checkpoint();
            const changed = transformCase(doc, segs, mode);
            if (changed.length > 0) {
              pages = rerender(doc);
              editor?.selectRanges(changed);
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
            printPages(handle.root, sp?.pageWidth ?? 816, sp?.pageHeight ?? 1056);
          },
          save: () => doc.save(),
        };
        apiRef.current = api;
        onReady?.(api);
      }
    })().catch((e: unknown) => {
      if (cancelled) return;
      const err = e instanceof Error ? e : new Error(String(e));
      setError(err);
      onError?.(err);
    });

    return () => {
      cancelled = true;
      editor?.detach();
      editor = null;
      handle?.destroy();
      handle = null;
    };
  }, [source, zoom, editable, commentAuthor, showComments, revisions]);

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
        style={{ background: "#e8eaed", overflow: "auto", height: "100%", ...style }}
      >
        {error && (
          <div style={{ padding: 16, color: "#b00020", fontFamily: "system-ui" }}>
            Failed to render document: {error.message}
          </div>
        )}
      </div>
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

export { DocxDocument, layoutDocument, renderToDom, printPages } from "@docxinweb/core";
export type { RunFormatPatch, SelectionFormat, ParagraphAlignment, PageLayoutPatch } from "@docxinweb/core";
export { DocxToolbar } from "./toolbar.js";
