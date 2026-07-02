import { useEffect, useRef, useState } from "react";
import {
  DocxDocument,
  DocxEditor,
  RenderHandle,
  RunFormatPatch,
  SelectionFormat,
  applyRunFormat,
  layoutDocument,
  renderToDom,
  selectionToSegments,
  summarizeSelection,
} from "@docxinweb/core";

export interface DocxViewApi {
  /** Apply character formatting to the current browser selection. */
  applyFormat(patch: RunFormatPatch): void;
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
      handle = renderToDom(doc, layout, container, { zoom });
      container.scrollTop = scrollTop;
      container.scrollLeft = scrollLeft;
      return layout.totalPages;
    };

    (async () => {
      const bytes = await toBytes(source);
      if (cancelled) return;
      if (typeof document !== "undefined" && document.fonts?.ready) {
        try {
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
        editor = new DocxEditor({
          doc,
          container: containerRef.current,
          getHandle: () => handle,
          rerender: () => {
            pages = rerender(doc);
          },
          zoom,
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
            applyRunFormat(doc, segments, patch);
            pages = rerender(doc);
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
export type { RunFormatPatch, SelectionFormat } from "@docxinweb/core";
export { DocxToolbar } from "./toolbar.js";
