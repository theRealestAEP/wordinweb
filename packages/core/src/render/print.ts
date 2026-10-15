import { DocxDocument } from "../docx.js";
import { StreamLayoutResult, layoutDocument, layoutDocumentStream } from "../layout/engine.js";
import { clearBreakCache } from "../layout/inline.js";
import { createMeasurer } from "../layout/measure.js";
import { LaidOutPage } from "../layout/types.js";
import { StreamingDocument, openStreamingDocument } from "../stream/body.js";
import { ByteSource, fetchSource } from "../stream/source.js";
import { RenderHandle, mediaUrlSink, paintsSettled, printStylesheet, pruneBlankSpans, releaseMediaUrls, renderToDom } from "./dom.js";

/** Stands in the package for an image part whose bytes went into a Blob (mediaUrlSink), so the document counts the part as present. */
const STREAMED = new Uint8Array(0);

/**
 * Lay out and render a whole .docx into `container` for a headless browser
 * to print. Resolves once the pages are final: fonts loaded, every image
 * converted and decoded. Print with the CSS page size to get one PDF page
 * per document page.
 */
export async function renderForPrint(bytes: Uint8Array, container: HTMLElement): Promise<{ pageCount: number }> {
  await document.fonts.ready;
  const doc = DocxDocument.load(bytes);
  const layout = layoutDocument(doc);
  renderToDom(doc, layout, container, { zoom: 1, pageGap: 0, pageShadow: false, comments: false });
  pruneBlankSpans(container);

  const first = layout.pages[0];
  const style = document.createElement("style");
  style.textContent = printStylesheet(first?.width ?? 816, first?.height ?? 1056, container);
  document.head.append(style);

  await paintsSettled();
  // A broken image still prints as an empty frame, as it shows on screen.
  await Promise.all(Array.from(container.querySelectorAll("img"), (img) => img.decode().catch(() => undefined)));
  return { pageCount: layout.pages.length };
}

export interface PrintStreamOptions {
  /** Pages rendered into the container at a time. */
  batchPages?: number;
}

export interface PrintBatch {
  /** The batch's first page, zero-based. */
  first: number;
  count: number;
}

/**
 * A .docx printed a batch of pages at a time, so memory stays flat whatever
 * the page count. The document's body is never in memory whole
 * (stream/body.ts); the counting pass lays every page out without painting
 * it and keeps only the page total and the bookmark pages; the rendering
 * pass lays the document out again with those, and renders each batch as
 * the caller asks for it:
 *
 *   const stream = openPrintStream(url, document.body, { batchPages: 25 });
 *   const { pageCount } = await stream.start();
 *   for (let batch = await stream.nextBatch(); batch; batch = await stream.nextBatch()) {
 *     // print the page (Page.printToPDF), then
 *     stream.release();
 *   }
 */
export interface PrintStream {
  /** Open the document and count its pages. */
  start(): Promise<{ pageCount: number; pageWidth: number; pageHeight: number }>;
  /** Render the next batch into the container; resolves once it is final, as
   * renderForPrint does for the whole document. Null once every page is
   * out. The printed batch must be released first. */
  nextBatch(): Promise<PrintBatch | null>;
  /** Drop the printed batch: its DOM, images and layout. */
  release(): void;
}

/**
 * `source` is a URL whose server honours Range requests, or bytes already
 * held. A caller with the file on disk and no server serves it under a URL
 * of its own over the DevTools protocol: navigate the page to that origin,
 * so the reads are same-origin, Fetch.enable, and fulfill each paused
 * request's Range from the file.
 */
export function openPrintStream(source: string | ByteSource, container: HTMLElement, options: PrintStreamOptions = {}): PrintStream {
  const batchPages = options.batchPages ?? 25;
  const measurer = createMeasurer();
  const style = document.createElement("style");
  let opened: StreamingDocument;
  let knownTotalPages: number | undefined;
  let final: StreamLayoutResult;
  let pageWidth = 816;
  let pageHeight = 1056;

  // The rendering pass runs between nextBatch calls: it fills `pending` and,
  // once a batch is full, waits to be woken by the next call.
  let run: Promise<void> | null = null;
  let done = false;
  let failed: unknown;
  let pending: LaidOutPage[] = [];
  let full: (() => void) | null = null;
  let wake: (() => void) | null = null;
  let first = 0;
  let mounted: { handle: RenderHandle; parts: Set<string> } | null = null;

  const onPage = async (page: LaidOutPage): Promise<void> => {
    pending.push(page);
    if (pending.length < batchPages) return;
    full?.();
    await new Promise<void>((resolve) => {
      wake = resolve;
    });
  };

  const mount = async (pages: LaidOutPage[]): Promise<void> => {
    const doc = opened.doc;
    const parts = new Set<string>();
    for (const page of pages) {
      for (const item of page.items) {
        if (item.kind !== "image") continue;
        parts.add(item.part);
        if (item.model3D) parts.add(item.model3D.part);
        if (item.embeddedObject) parts.add(item.embeddedObject.part);
      }
    }
    // An image the browser draws as it is goes into a Blob a piece at a time,
    // and the part is marked present so renderItem draws it from that URL;
    // one it converts (TIFF, WMF, EMF) it needs whole.
    for (const part of parts) {
      const sink = mediaUrlSink(doc, part);
      if (!sink) {
        const bytes = await opened.media(part);
        if (bytes) doc.pkg.raw()[part] = bytes;
      } else if (await opened.mediaPieces(part, sink.push)) {
        sink.end();
        doc.pkg.raw()[part] = STREAMED;
      }
    }
    const handle = renderToDom(doc, { pages, totalPages: final.totalPages }, container, { zoom: 1, pageGap: 0, pageShadow: false, comments: false });
    mounted = { handle, parts };
    pruneBlankSpans(handle.root);
    style.textContent = printStylesheet(pageWidth, pageHeight, handle.root);
    await paintsSettled();
    await settled(handle.root);
  };

  return {
    async start() {
      await document.fonts.ready;
      opened = await openStreamingDocument(typeof source === "string" ? fetchSource(source) : source);
      const count = (total: number | undefined): Promise<StreamLayoutResult> => {
        clearBreakCache(measurer);
        return layoutDocumentStream(opened.doc, {
          measurer,
          prescan: opened.prescan,
          knownTotalPages: total,
          onPage: (page) => {
            if (page.index === 1) {
              pageWidth = page.width;
              pageHeight = page.height;
            }
            // The break cache keeps a line-metrics entry per paragraph laid;
            // dropped a batch at a time here as the rendering pass drops it.
            if (page.index % batchPages === 0) clearBreakCache(measurer);
          },
        });
      };
      // The body's NUMPAGES fields converge as in layoutDocument.
      final = await count(undefined);
      if (opened.bodyHasPageTotal) {
        for (let pass = 0; pass < 2; pass++) {
          knownTotalPages = final.totalPages;
          final = await count(knownTotalPages);
          if (final.totalPages === knownTotalPages) break;
        }
      }
      document.head.append(style);
      return { pageCount: final.totalPages, pageWidth, pageHeight };
    },

    async nextBatch() {
      if (mounted) throw new Error("release the printed batch first");
      const ready = new Promise<void>((resolve) => {
        full = resolve;
      });
      if (!run) {
        clearBreakCache(measurer);
        run = layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, knownTotalPages, final, onPage }).then(
          () => {
            done = true;
            full?.();
          },
          (error: unknown) => {
            failed = error;
            full?.();
          },
        );
      } else {
        const resume = wake;
        wake = null;
        resume?.();
      }
      if (pending.length < batchPages && !done && failed === undefined) await ready;
      full = null;
      if (failed !== undefined) throw failed;
      const pages = pending;
      pending = [];
      if (pages.length === 0) return null;
      await mount(pages);
      const batch = { first, count: pages.length };
      first += pages.length;
      return batch;
    },

    release() {
      if (!mounted) return;
      mounted.handle.destroy();
      releaseMediaUrls(opened.doc);
      for (const part of mounted.parts) delete opened.doc.pkg.raw()[part];
      mounted = null;
      clearBreakCache(measurer);
    },
  };
}

/** Resolves once the batch is final: its fonts loaded, every image decoded
 * (a broken one prints as an empty frame, as it shows on screen). */
async function settled(root: ParentNode): Promise<void> {
  await document.fonts.ready;
  await Promise.all(Array.from(root.querySelectorAll("img"), (img) => img.decode().catch(() => undefined)));
}
