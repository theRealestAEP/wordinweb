import { closeSync, openSync, readSync, statSync } from "node:fs";
import { Writable } from "node:stream";
import { StreamLayoutResult, layoutDocumentStream } from "../layout/engine.js";
import { clearBreakCache } from "../layout/inline.js";
import { LaidOutPage } from "../layout/types.js";
import { StreamingDocument, openStreamingDocument } from "../stream/body.js";
import { ByteSource } from "../stream/source.js";
import { CommentColumn, RAIL_WIDTH } from "./comments.js";
import { PageDrawer } from "./draw.js";
import { FontSet } from "./fonts.js";
import { WrittenImage, writeImage } from "./images.js";
import { NodeMeasurer } from "./measure.js";
import { Resources } from "./resources.js";
import { PdfWriter, pdfString } from "./writer.js";

export interface RenderPdfOptions {
  /** Directories scanned for fonts; the system directories and the
   * @fontsource Carlito and Caladea packages by default (see fonts.ts). */
  fontDirs?: string[];
  /** The PDF's Title entry. */
  title?: string;
  /** Print the review comments: each sheet is the page with a comment
   * column at its right (see comments.ts), every sheet the size of the
   * largest page plus the column. */
  comments?: boolean;
}

export interface RenderPdfResult {
  pages: number;
  /** Faces embedded, "family subfamily". */
  fonts: string[];
  /** Families the document asked for that no font directory holds. */
  missingFonts: string[];
  /** Picture parts drawn as empty frames: no decoder for their format. */
  unsupportedImages: string[];
  /** Items of kinds not drawn yet. */
  undrawn: string[];
}

/** A .docx on disk, read a range at a time; the file is opened per read,
 * so nothing stays open between them. */
export function fileSource(path: string): ByteSource {
  return {
    size: async () => statSync(path).size,
    async read(start, end) {
      const out = Buffer.allocUnsafe(Math.max(0, end - start));
      const fd = openSync(path, "r");
      let done = 0;
      try {
        while (done < out.length) {
          const n = readSync(fd, out, done, out.length - done, start + done);
          if (n === 0) break;
          done += n;
        }
      } finally {
        closeSync(fd);
      }
      return new Uint8Array(out.buffer, out.byteOffset, done);
    },
  };
}

/**
 * Render a .docx to PDF with no browser: the document streams in
 * (stream/body.ts), a counting pass learns the page total and bookmark
 * pages, and the rendering pass writes each page's objects to `out` as the
 * layout finishes it. Fonts embed once at the end, subset to the glyphs
 * used. `out` is ended when the PDF is complete.
 */
export async function renderDocxToPdf(source: string | ByteSource, out: Writable, options: RenderPdfOptions = {}): Promise<RenderPdfResult> {
  const fonts = new FontSet(options.fontDirs);
  const measurer = new NodeMeasurer(fonts);
  const opened = await openStreamingDocument(typeof source === "string" ? fileSource(source) : source);
  const final = await countPages(opened, measurer);

  const writer = new PdfWriter(out);
  await writer.header();
  const catalog = writer.allocate();
  const pagesNum = writer.allocate();
  const info = writer.allocate();
  const resources = new Resources(writer);
  const pageNums: number[] = [];
  const images = new Map<string, WrittenImage | null>();
  const unsupportedImages = new Set<string>();
  const undrawn = new Set<string>();

  const column = options.comments && opened.doc.comments.length > 0 ? new CommentColumn(opened.doc, measurer, resources, opened.commentIds) : null;
  const sheet = column ? { width: final.maxWidth + RAIL_WIDTH, height: final.maxHeight } : null;

  const writePage = async (drawer: PageDrawer, width: number, height: number, xobjects: string[]): Promise<void> => {
    const content = writer.allocate();
    await writer.deflatedStream(content, "", Buffer.from(drawer.ops.join("\n"), "latin1"));
    const pageNum = writer.allocate();
    pageNums.push(pageNum);
    await writer.object(
      pageNum,
      `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${pt(width)} ${pt(height)}] /Contents ${content} 0 R ` +
        `/Resources << ${resources.shared()}${xobjects.length ? ` /XObject << ${xobjects.join(" ")} >>` : ""} >> >>`,
    );
  };

  const onPage = async (page: LaidOutPage): Promise<void> => {
    const names = new Map<string, string>();
    for (const item of page.items) {
      if (item.kind === "warptext") undrawn.add(item.kind);
      if (item.kind !== "image" || names.has(item.part)) continue;
      let image = images.get(item.part);
      if (image === undefined) {
        const size = opened.mediaSize(item.part);
        image = size === undefined ? null : await writeImage(writer, item.part, size, (onPiece) => opened.mediaPieces(item.part, onPiece));
        images.set(item.part, image);
        if (!image) unsupportedImages.add(item.part);
      }
      if (image) names.set(item.part, `Im${image.objNum}`);
    }
    const drawer = new PageDrawer(fonts, resources, opened.doc.theme, page, 0, sheet?.height ?? page.height);
    drawer.drawPage((part) => names.get(part) ?? null);
    if (column && sheet) column.draw(page, drawer, sheet.height);
    const xobjects = [...images.entries()].filter(([part]) => names.has(part)).map(([part, image]) => `/${names.get(part)} ${(image as WrittenImage).objNum} 0 R`);
    await writePage(drawer, sheet?.width ?? page.width, sheet?.height ?? page.height, xobjects);
    if (page.index % 25 === 0) clearBreakCache(measurer);
  };
  clearBreakCache(measurer);
  await layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, knownTotalPages: final.knownTotalPages, final: final.result, onPage });

  // Balloons past the last page go on sheets of their own.
  while (column && sheet && column.hasCarried()) {
    const blank: LaidOutPage = { width: final.maxWidth, height: final.maxHeight, index: 0, number: 0, items: [], bodyTop: 0, bodyBottom: 0, hfStart: 0, columnBands: [] };
    const drawer = new PageDrawer(fonts, resources, opened.doc.theme, blank, 0, sheet.height);
    column.draw(blank, drawer, sheet.height);
    await writePage(drawer, sheet.width, sheet.height, []);
  }

  await resources.write();
  await writer.object(pagesNum, `<< /Type /Pages /Count ${pageNums.length} /Kids [${pageNums.map((n) => `${n} 0 R`).join(" ")}] >>`);
  await writer.object(catalog, `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`);
  await writer.object(info, `<< /Producer (wordinweb)${options.title ? ` /Title ${pdfString(options.title)}` : ""} >>`);
  await writer.finish(catalog, info);
  await new Promise<void>((resolve, reject) => out.end((error?: Error | null) => (error ? reject(error) : resolve())));
  return {
    pages: pageNums.length,
    fonts: resources.embedded(),
    missingFonts: [...fonts.missing],
    unsupportedImages: [...unsupportedImages],
    undrawn: [...undrawn],
  };
}

interface Count {
  result: StreamLayoutResult;
  knownTotalPages: number | undefined;
  /** The largest page's sides, for sheets that must all be one size. */
  maxWidth: number;
  maxHeight: number;
}

/** The counting pass, and the NUMPAGES convergence passes openPrintStream runs. */
async function countPages(opened: StreamingDocument, measurer: NodeMeasurer): Promise<Count> {
  let maxWidth = 0;
  let maxHeight = 0;
  const count = (total: number | undefined): Promise<StreamLayoutResult> => {
    clearBreakCache(measurer);
    return layoutDocumentStream(opened.doc, {
      measurer,
      prescan: opened.prescan,
      knownTotalPages: total,
      onPage: (page) => {
        maxWidth = Math.max(maxWidth, page.width);
        maxHeight = Math.max(maxHeight, page.height);
        if (page.index % 25 === 0) clearBreakCache(measurer);
      },
    });
  };
  let result = await count(undefined);
  let knownTotalPages: number | undefined;
  if (opened.bodyHasPageTotal) {
    for (let pass = 0; pass < 2; pass++) {
      knownTotalPages = result.totalPages;
      result = await count(knownTotalPages);
      if (result.totalPages === knownTotalPages) break;
    }
  }
  return { result, knownTotalPages, maxWidth, maxHeight };
}

function pt(px: number): string {
  return String(Math.round(px * 0.75 * 1000) / 1000);
}
