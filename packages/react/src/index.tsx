import { useEffect, useRef, useState } from "react";
import {
  DocxDocument,
  LayoutResult,
  RenderHandle,
  layoutDocument,
  renderToDom,
} from "@docxinweb/core";

export interface DocxViewProps {
  /** The document: raw bytes, a File/Blob, or a URL to fetch. */
  source: ArrayBuffer | Uint8Array | Blob | string;
  /** Zoom factor, 1 = 100%. */
  zoom?: number;
  className?: string;
  style?: React.CSSProperties;
  onLoad?: (info: { pageCount: number; document: DocxDocument }) => void;
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
 * High-fidelity paginated DOCX viewer.
 *
 * ```tsx
 * <DocxView source="/report.docx" zoom={1} />
 * ```
 */
export function DocxView({ source, zoom = 1, className, style, onLoad, onError }: DocxViewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState<Error | null>(null);

  useEffect(() => {
    let cancelled = false;
    let handle: RenderHandle | null = null;
    setError(null);

    (async () => {
      const bytes = await toBytes(source);
      if (cancelled) return;
      // Wait for document fonts so canvas measurement uses the real metrics.
      if (typeof document !== "undefined" && document.fonts?.ready) {
        try {
          await document.fonts.ready;
        } catch {
          /* non-fatal */
        }
      }
      if (cancelled) return;
      const doc = DocxDocument.load(bytes);
      const layout: LayoutResult = layoutDocument(doc);
      const container = containerRef.current;
      if (!container || cancelled) return;
      handle = renderToDom(doc, layout, container, { zoom });
      onLoad?.({ pageCount: layout.totalPages, document: doc });
    })().catch((e: unknown) => {
      if (cancelled) return;
      const err = e instanceof Error ? e : new Error(String(e));
      setError(err);
      onError?.(err);
    });

    return () => {
      cancelled = true;
      handle?.destroy();
    };
  }, [source, zoom]);

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
