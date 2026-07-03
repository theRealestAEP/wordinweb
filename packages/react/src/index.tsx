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
  insertImageAt,
  insertTableAfter,
  layoutDocument,
  renderToDom,
  selectionToSegments,
  setPageLayout,
  setParagraphAlignment,
  summarizeSelection,
} from "@docxinweb/core";

export interface DocxViewApi {
  /** Apply character formatting to the current browser selection. */
  applyFormat(patch: RunFormatPatch): void;
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
  /** Change margins / page size / orientation (inches). */
  setPageLayout(patch: PageLayoutPatch): void;
  /** Effective formatting of the current selection (toolbar state), or null. */
  getSelectionFormat(): SelectionFormat | null;
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
}: DocxViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState<Error | null>(null);

  useEffect(() => {
    let cancelled = false;
    let handle: RenderHandle | null = null;
    let editor: DocxEditor | null = null;
    setError(null);

    const rerender = (doc: DocxDocument): number => {
      const layout = layoutDocument(doc);
      const container = containerRef.current;
      if (!container) return 0;
      // Re-rendering replaces the page DOM; keep the user's scroll position
      // (destroy-then-append clamps scrollTop to 0 otherwise).
      const { scrollTop, scrollLeft } = container;
      handle?.destroy();
      handle = renderToDom(doc, layout, container, { zoom, interactive: editable });
      container.scrollTop = scrollTop;
      container.scrollLeft = scrollLeft;
      return layout.totalPages;
    };

    (async () => {
      const bytes = await toBytes(source);
      if (cancelled) return;
      if (typeof document !== "undefined" && document.fonts?.ready) {
        try {
          // Canvas measurement doesn't trigger webfont loads; request the
          // metric-compatible substitutes explicitly if the host provides them.
          const loads: Promise<unknown>[] = [];
          for (const fam of ["Carlito", "Caladea"]) {
            for (const variant of ["", "italic ", "bold ", "bold italic "]) {
              loads.push(document.fonts.load(`${variant}16px ${fam}`).catch(() => []));
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
        });
        editor.attach();
        const api: DocxViewApi = {
          document: doc,
          pageCount: () => pages,
          getSelectionFormat: () =>
            handle ? summarizeSelection(selectionToSegments(handle.bindings)) : null,
          applyFormat: (patch) => {
            if (!handle) return;
            const segments = selectionToSegments(handle.bindings);
            if (segments.length === 0) return;
            history.checkpoint();
            const formatted = applyRunFormat(doc, segments, patch);
            pages = rerender(doc);
            // Restore a visible selection over the formatted text so repeated
            // toolbar actions compose (Google Docs behavior).
            if (formatted.length > 0 && handle) {
              const findPos = (t: unknown, off: number, atEnd: boolean) => {
                for (const b of handle!.bindings) {
                  const src = b.item.src;
                  if (!src || src.t !== t) continue;
                  const start = src.offset;
                  const end = src.offset + b.item.text.length;
                  if (off >= start && off <= end && (atEnd ? off > start || off === end : off < end || off === start)) {
                    return { node: b.el.firstChild, offset: off - start };
                  }
                }
                return null;
              };
              const first = formatted[0];
              const last = formatted[formatted.length - 1];
              const s = findPos(first.t, first.start, false);
              const e2 = findPos(last.t, last.end, true);
              if (s?.node && e2?.node) {
                try {
                  const range = document.createRange();
                  range.setStart(s.node, s.offset);
                  range.setEnd(e2.node, e2.offset);
                  const sel = window.getSelection();
                  sel?.removeAllRanges();
                  sel?.addRange(range);
                } catch {
                  /* selection restore is best-effort */
                }
              }
            }
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
            if (insertImageAt(doc, caret.t, relId, bmp.width * scale, bmp.height * scale)) {
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
          setPageLayout: (patch) => {
            history.checkpoint();
            if (setPageLayout(doc, patch)) pages = rerender(doc);
          },
          save: () => doc.save(),
        };
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
  }, [source, zoom, editable]);

  return (
    <div
      ref={containerRef}
      className={className}
      style={{ background: "#e8eaed", overflow: "auto", ...style }}
    >
      {error && (
        <div style={{ padding: 16, color: "#b00020", fontFamily: "system-ui" }}>
          Failed to render document: {error.message}
        </div>
      )}
    </div>
  );
}

export { DocxDocument, layoutDocument, renderToDom } from "@docxinweb/core";
export type { RunFormatPatch, SelectionFormat, ParagraphAlignment, PageLayoutPatch } from "@docxinweb/core";
export { DocxToolbar } from "./toolbar.js";
