import { closeSync, openSync, readSync, statSync } from "node:fs";
import { Writable } from "node:stream";
import { StreamLayoutResult, layoutDocumentStream } from "../layout/engine.js";
import { clearBreakCache } from "../layout/inline.js";
import { LaidOutPage } from "../layout/types.js";
import { Block } from "../model.js";
import { StreamingDocument, openStreamingDocument } from "../stream/body.js";
import { XmlElement, localName } from "../xml.js";
import { ByteSource } from "../stream/source.js";
import { CommentColumn, RAIL_WIDTH } from "./comments.js";
import { IMAGE_PLACEHOLDER } from "./canvas.js";
import { PageDrawer, Picture } from "./draw.js";
import { PdfRenderError, asPdfRenderError, unsupportedScript } from "./errors.js";
import { FontSet } from "./fonts.js";
import { writeImage, writeSamples } from "./images.js";
import { convertMetafile } from "./metafile.js";
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
  try {
    return await render(source, out, options);
  } catch (error) {
    throw asPdfRenderError(error);
  }
}

async function render(source: string | ByteSource, out: Writable, options: RenderPdfOptions): Promise<RenderPdfResult> {
  const fonts = new FontSet(options.fontDirs);
  const measurer = new NodeMeasurer(fonts);
  // Text this renderer cannot shape is refused before any layout, as the
  // body's blocks are first parsed; the held parts (headers, footers,
  // notes) and the comments to be printed are checked once it is open.
  const opened = await openStreamingDocument(typeof source === "string" ? fileSource(source) : source, {
    onBlock: (block, index) => {
      const found = unsupportedInBlock(block);
      if (found) throw unsupported(found.script, `in paragraph ${index + 1}, "${found.sample}"`);
    },
  });
  for (const root of opened.doc.revisionRoots()) {
    if (root === opened.doc.docRoot) continue;
    const found = unsupportedInXml(root);
    if (found) throw unsupported(found.script, `in a header, footer or note, "${found.sample}"`);
  }
  if (options.comments) {
    for (const comment of opened.doc.comments) {
      const script = unsupportedScript(comment.text);
      if (script) throw unsupported(script, `in a comment by ${comment.author || "an unnamed author"}`);
    }
  }
  const final = await countPages(opened, measurer);

  const writer = new PdfWriter(out);
  await writer.header();
  const catalog = writer.allocate();
  const pagesNum = writer.allocate();
  const info = writer.allocate();
  const resources = new Resources(writer);
  const pageNums: number[] = [];
  /** Each picture part once: its picture and the XObjects it needs on a page. */
  const pictures = new Map<string, { picture: Picture; xobjects: string[] } | null>();
  const unsupportedImages = new Set<string>();

  const loadPicture = async (part: string): Promise<{ picture: Picture; xobjects: string[] } | null> => {
    const size = opened.mediaSize(part);
    if (size === undefined) return null;
    const ext = part.slice(part.lastIndexOf(".") + 1).toLowerCase();
    if (ext === "emf" || ext === "wmf") {
      const bytes = await opened.media(part);
      const canvas = bytes && (await convertMetafile(ext, bytes, fonts, resources));
      if (!canvas) return null;
      // The bitmaps the metafile drew become XObjects, named in its operators.
      const names: string[] = [];
      const xobjects: string[] = [];
      for (const image of canvas.images) {
        const written = await writeSamples(writer, image);
        names.push(`Im${written.objNum}`);
        xobjects.push(`/Im${written.objNum} ${written.objNum} 0 R`);
      }
      const ops = canvas.ops.join("\n").replace(IMAGE_PLACEHOLDER, (_, index) => `/${names[Number(index)]}`);
      return { picture: { vector: { ops, width: canvas.width, height: canvas.height } }, xobjects };
    }
    const image = await writeImage(writer, part, size, (onPiece) => opened.mediaPieces(part, onPiece));
    return image ? { picture: { name: `Im${image.objNum}` }, xobjects: [`/Im${image.objNum} ${image.objNum} 0 R`] } : null;
  };

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
    const shown = new Map<string, Picture>();
    const xobjects: string[] = [];
    for (const item of page.items) {
      if (item.kind !== "image" || shown.has(item.part)) continue;
      let loaded = pictures.get(item.part);
      if (loaded === undefined) {
        loaded = await loadPicture(item.part);
        pictures.set(item.part, loaded);
        if (!loaded) unsupportedImages.add(item.part);
      }
      if (!loaded) continue;
      shown.set(item.part, loaded.picture);
      xobjects.push(...loaded.xobjects);
    }
    const drawer = new PageDrawer(fonts, resources, opened.doc.theme, page, 0, sheet?.height ?? page.height);
    drawer.drawPage((part) => shown.get(part) ?? null);
    if (column && sheet) column.draw(page, drawer, sheet.height);
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
  };
}

function unsupported(script: string, where: string): PdfRenderError {
  return new PdfRenderError("unsupported_script", `${script} text ${where} needs a shaper this renderer does not have`);
}

function sample(text: string): string {
  const trimmed = text.trim();
  return trimmed.length > 40 ? `${trimmed.slice(0, 40)}…` : trimmed;
}

/** The first unsupported script in a block's text, with a sample of the text. */
function unsupportedInBlock(block: Block): { script: string; sample: string } | undefined {
  if (block.type !== "paragraph") {
    for (const row of block.rows) for (const cell of row.cells) for (const nested of cell.blocks) {
      const found = unsupportedInBlock(nested);
      if (found) return found;
    }
    return undefined;
  }
  for (const child of block.children) {
    for (const run of child.type === "run" ? [child] : child.runs) {
      for (const content of run.content) {
        if (content.kind !== "text") continue;
        const script = unsupportedScript(content.text);
        if (script) return { script, sample: sample(content.text) };
      }
    }
  }
  return undefined;
}

/** The first unsupported script in a part's w:t elements. */
function unsupportedInXml(el: XmlElement): { script: string; sample: string } | undefined {
  if (localName(el.name) === "t") {
    const script = unsupportedScript(el.text);
    return script ? { script, sample: sample(el.text) } : undefined;
  }
  for (const child of el.children) {
    const found = unsupportedInXml(child);
    if (found) return found;
  }
  return undefined;
}

interface Count {
  result: StreamLayoutResult;
  knownTotalPages: number | undefined;
  /** The largest page's sides, for sheets that must all be one size. */
  maxWidth: number;
  maxHeight: number;
}

/** The counting pass, and the NUMPAGES convergence passes layoutDocument runs. */
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
