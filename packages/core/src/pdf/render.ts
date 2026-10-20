import { closeSync, openSync, readSync, statSync } from "node:fs";
import { Writable } from "node:stream";
import { StreamLayoutResult, layoutDocumentStream } from "../layout/engine.js";
import { clearBreakCache } from "../layout/inline.js";
import { LaidOutPage } from "../layout/types.js";
import { Block } from "../model.js";
import { StreamingDocument, openStreamingDocument } from "../stream/body.js";
import { XmlElement, attr, localName } from "../xml.js";
import { ByteSource } from "../stream/source.js";
import { CommentColumn, RAIL_WIDTH, UndrawnComments } from "./comments.js";
import { IMAGE_PLACEHOLDER } from "./canvas.js";
import { PageDrawer, Picture } from "./draw.js";
import { PdfRenderError, asPdfRenderError, unsupportedScript } from "./errors.js";
import { FontSet, defaultFontDirs, deobfuscateFont } from "./fonts.js";
import { writeImage, writeSamples } from "./images.js";
import { bookmarkTops, linkAction, linkBoxes, pdfName } from "./links.js";
import { convertMetafile } from "./metafile.js";
import { NodeMeasurer } from "./measure.js";
import { Resources } from "./resources.js";
import { loadShaper } from "./shaper.js";
import { PdfWriter, num, pdfString } from "./writer.js";

// The package's "./pdf" entry is this module: everything a caller needs is here.
export { PdfRenderError } from "./errors.js";
export type { PdfRenderErrorCode } from "./errors.js";
export { bytesSource } from "../stream/source.js";
export type { ByteSource } from "../stream/source.js";
export { docToDocx } from "../doc/hooks.js";
export type { ConvertHooks } from "../doc/convert.js";
export { isWordDocumentFile } from "../doc/streams.js";

export interface RenderPdfOptions {
  /** Directories scanned for fonts INSTEAD of the defaults (the system
   * directories, the @fontsource Carlito and Caladea packages, and the
   * WORDINWEB_FONT_DIRS environment variable's; see fonts.ts). */
  fontDirs?: string[];
  /** Directories scanned for fonts ahead of the defaults (or of
   * `fontDirs`): a font file in one wins over the substitutes for its
   * family. WORDINWEB_FONT_DIRS adds such directories to every render. */
  extraFontDirs?: string[];
  /** The PDF's Title entry. */
  title?: string;
  /** Print the review comments: each sheet is the page with a comment
   * column at its right (see comments.ts), every sheet the size of the
   * largest page plus the column, or taller when its page's cards need it:
   * one sheet per page. */
  comments?: boolean;
  /**
   * Print the document's own tracked changes as Word's "All Markup" shows
   * them, instead of the final text: insertions underlined and deletions
   * struck through in their author's color (Word's "by author" colors, in
   * the order the authors first appear: red, blue, green, violet, dark red,
   * teal, dark yellow, gray), moved text double underlined or struck in
   * green, a changed picture framed in the color (a deleted one struck
   * through), a bar in the left margin beside every changed line, and, in
   * the margin column with the comments (which this prints too), a balloon
   * per change with its author, date and what it is ("Inserted", "Deleted",
   * "Moved from", "Inserted row", "Formatted: Font: Bold", ...).
   */
  markup?: boolean;
  /**
   * Name the authors of the comments and the tracked changes. Default true.
   * False prints no author's name or initials anywhere: each comment's and
   * change's balloon keeps its date, what it is and its text, with no name
   * line and no avatar, and every author's changes are in one color, red
   * (moves stay green). The comments' leaders are one color too.
   */
  authors?: boolean;
}

export interface RenderPdfResult {
  pages: number;
  /** Faces embedded, "family subfamily". */
  fonts: string[];
  /** Families the document asked for that no font directory holds; a
   * substitute or fallback face drew them. */
  missingFonts: string[];
  /** Characters no installed font has, as "U+4E2D 中": each drew as an empty box. */
  missingCharacters: string[];
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
 *
 * A link (w:hyperlink, a HYPERLINK field) is a /Link annotation over each
 * line of its text, written with the page: a URL one opens it, a bookmark
 * one goes to a named destination (the catalog's /Dests, written at the
 * end, when every bookmark's page object is known) at the top of the
 * bookmark's line, or at the top of its page when no page drew the text
 * it starts on.
 */
export async function renderDocxToPdf(source: string | ByteSource, out: Writable, options: RenderPdfOptions = {}): Promise<RenderPdfResult> {
  try {
    return await render(source, out, options);
  } catch (error) {
    throw asPdfRenderError(error);
  }
}

async function render(source: string | ByteSource, out: Writable, options: RenderPdfOptions): Promise<RenderPdfResult> {
  // The shaper first (once per process): the scripts it is there for are refused only without it.
  await loadShaper();
  const fonts = new FontSet([...(options.extraFontDirs ?? []), ...(options.fontDirs ?? defaultFontDirs())]);
  const measurer = new NodeMeasurer(fonts);
  // Text this renderer cannot shape is refused before any layout, naming
  // where it is: the body's blocks as they are first parsed (with markup,
  // the deleted text and the change authors the balloons name too, when
  // they name them); the
  // held parts (headers, footers, notes) and the comments to be printed
  // once the document is open. FontSet.shape refuses whatever else reaches
  // the page (a numbering label, a balloon's words).
  let hasChanges = false;
  const authors = options.authors !== false;
  const opened = await openStreamingDocument(typeof source === "string" ? fileSource(source) : source, {
    markup: options.markup,
    authors: options.authors,
    onBlock: (block, index) => {
      const found = unsupportedInBlock(block, authors);
      if (found) throw unsupported(found.script, `in paragraph ${index + 1}, "${found.sample}"`);
      hasChanges ||= blockHasChanges(block);
    },
  });
  // The fonts the document embeds, before any text is measured: the layout
  // breaks lines by the faces the PDF embeds.
  for (const font of opened.doc.embeddedFonts) {
    const bytes = await opened.media(font.part);
    if (!bytes) continue;
    try {
      fonts.embed(font.family, font.bold, font.italic, font.key ? deobfuscateFont(bytes, font.key) : bytes, font.part);
    } catch {
      // Not a font file this renderer opens (broken, or a format sfnt.ts
      // does not read): the family resolves as if it were not embedded.
    }
  }
  // The body root here is the document's outline: its own section
  // properties, whose w:sectPrChange is a change too.
  // The comments some part marks (a range start or a reference): a comment
  // comments.xml lists but nothing anchors is not shown, as in Word, and a
  // document with no other gets no column.
  const marked = new Set<string>();
  for (const root of opened.doc.revisionRoots()) {
    const found = unsupportedInXml(root, !!options.markup, authors);
    if (found) throw unsupported(found.script, `in a header, footer or note, "${found.sample}"`);
    hasChanges ||= xmlHasChanges(root);
    commentMarkers(root, marked);
  }
  const printComments =
    (options.comments || options.markup) &&
    opened.doc.comments.some((c) => !c.parentId && (marked.has(c.id) || opened.commentAnchor(c.id) !== undefined));
  if (printComments) {
    for (const comment of opened.doc.comments) {
      for (const text of authors ? [comment.text, comment.author, comment.initials ?? ""] : [comment.text]) {
        const script = unsupportedScript(text);
        if (script) throw unsupported(script, `in a comment by ${comment.author || "an unnamed author"}, "${sample(text)}"`);
      }
    }
  }
  const undrawn = printComments ? new UndrawnComments(opened.doc, opened) : null;
  const final = await countPages(opened, measurer, undrawn);

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
    if (ext === "emf" || ext === "wmf" || ext === "pict") {
      const bytes = await opened.media(part);
      const canvas = bytes && (await convertMetafile(ext, bytes, fonts, resources));
      if (!canvas) return null;
      // The bitmaps the metafile drew become XObjects, named in its operators.
      const names: string[] = [];
      const xobjects: string[] = [];
      for (const image of canvas.images) {
        const written =
          "jpeg" in image
            ? ((await writeImage(writer, "picture.jpg", image.jpeg.length, async (onPiece) => {
                await onPiece(image.jpeg);
                return true;
              })) ?? (await writeSamples(writer, { width: 1, height: 1, gray: true, samples: Uint8Array.of(255), alpha: Uint8Array.of(0) })))
            : await writeSamples(writer, image);
        names.push(`Im${written.objNum}`);
        xobjects.push(`/Im${written.objNum} ${written.objNum} 0 R`);
      }
      const ops = canvas.ops.join("\n").replace(IMAGE_PLACEHOLDER, (_, index) => `/${names[Number(index)]}`);
      return { picture: { vector: { ops, width: canvas.width, height: canvas.height } }, xobjects };
    }
    const image = await writeImage(writer, part, size, (onPiece) => opened.mediaPieces(part, onPiece));
    return image ? { picture: { name: `Im${image.objNum}` }, xobjects: [`/Im${image.objNum} ${image.objNum} 0 R`] } : null;
  };

  const column = printComments || (options.markup && hasChanges) ? new CommentColumn(opened.doc, measurer, resources, opened, undrawn?.anchors() ?? new Map(), authors) : null;
  const sheet = column ? { width: final.maxWidth + RAIL_WIDTH, height: final.maxHeight } : null;
  /** Each page's top edge in PDF units, by page index from 0. */
  const pageTops: string[] = [];
  /** Where each bookmark some page drew the text of begins: the page's index from 0 and the line's top in PDF units. */
  const bookmarkAt = new Map<string, { page: number; y: number }>();

  const writePage = async (drawer: PageDrawer, width: number, height: number, xobjects: string[], annots: number[] = []): Promise<void> => {
    const content = writer.allocate();
    await writer.deflatedStream(content, "", Buffer.from(drawer.ops.join("\n"), "latin1"));
    const pageNum = writer.allocate();
    pageNums.push(pageNum);
    pageTops.push(pt(height));
    await writer.object(
      pageNum,
      `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${pt(width)} ${pt(height)}] /Contents ${content} 0 R ` +
        `/Resources << ${resources.shared()}${xobjects.length ? ` /XObject << ${xobjects.join(" ")} >>` : ""} >>` +
        `${annots.length ? ` /Annots [${annots.map((n) => `${n} 0 R`).join(" ")}]` : ""} >>`,
    );
  };

  /** The page's link annotations, written before the page that lists them. */
  const writeLinks = async (page: LaidOutPage, drawer: PageDrawer): Promise<number[]> => {
    const annots: number[] = [];
    for (const box of linkBoxes(page)) {
      const action = linkAction(box.href, final.result.bookmarkPageIndices);
      if (!action) continue;
      const annot = writer.allocate();
      const rect = [drawer.x(box.x), drawer.y(box.y + box.height), drawer.x(box.x + box.width), drawer.y(box.y)].map(num).join(" ");
      await writer.object(annot, `<< /Type /Annot /Subtype /Link /Rect [${rect}] /Border [0 0 0] ${action} >>`);
      annots.push(annot);
    }
    return annots;
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
    // The rail first: a sheet whose cards reach past its bottom is taller.
    const rail = column && sheet ? column.layout(page, sheet.height, printComments && page.index === final.result.totalPages) : null;
    const height = rail?.height ?? page.height;
    const drawer = new PageDrawer(fonts, resources, opened.doc.theme, page, 0, height);
    drawer.drawPage((part) => shown.get(part) ?? null);
    rail?.draw(drawer);
    for (const [name, top] of bookmarkTops(page)) if (!bookmarkAt.has(name)) bookmarkAt.set(name, { page: page.index - 1, y: drawer.y(top) });
    await writePage(drawer, sheet?.width ?? page.width, height, xobjects, await writeLinks(page, drawer));
    if (page.index % 25 === 0) clearBreakCache(measurer);
  };
  clearBreakCache(measurer);
  await layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, knownTotalPages: final.knownTotalPages, final: final.result, onPage });

  await resources.write();
  await writer.object(pagesNum, `<< /Type /Pages /Count ${pageNums.length} /Kids [${pageNums.map((n) => `${n} 0 R`).join(" ")}] >>`);
  // The named destinations: every bookmark, at its line or at its page's top.
  let dests = "";
  if (final.result.bookmarkPageIndices.size > 0) {
    const destsNum = writer.allocate();
    const entries: string[] = [];
    for (const [name, index] of final.result.bookmarkPageIndices) {
      const at = bookmarkAt.get(name);
      const page = at?.page ?? index;
      if (page < 0 || page >= pageNums.length) continue;
      entries.push(`${pdfName(name)} [${pageNums[page]} 0 R /XYZ null ${at ? num(at.y) : pageTops[page]} null]`);
    }
    await writer.object(destsNum, `<< ${entries.join(" ")} >>`);
    dests = ` /Dests ${destsNum} 0 R`;
  }
  await writer.object(catalog, `<< /Type /Catalog /Pages ${pagesNum} 0 R${dests} >>`);
  await writer.object(info, `<< /Producer (wordinweb)${options.title ? ` /Title ${pdfString(options.title)}` : ""} >>`);
  await writer.finish(catalog, info);
  await new Promise<void>((resolve, reject) => out.end((error?: Error | null) => (error ? reject(error) : resolve())));
  return {
    pages: pageNums.length,
    fonts: resources.embedded(),
    missingFonts: [...fonts.missing],
    missingCharacters: [...fonts.missingCodePoints].sort((a, b) => a - b).map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, "0")} ${String.fromCodePoint(cp)}`),
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

/** The first unsupported script in a block's text, or in the authors of
 * its tracked changes when the balloons name them, with a sample of the text. */
function unsupportedInBlock(block: Block, authors: boolean): { script: string; sample: string } | undefined {
  if (block.type !== "paragraph") {
    for (const row of block.rows) for (const cell of row.cells) for (const nested of cell.blocks) {
      const found = unsupportedInBlock(nested, authors);
      if (found) return found;
    }
    return undefined;
  }
  const texts: string[] = [];
  for (const child of block.children) {
    for (const run of child.type === "run" ? [child] : child.runs) {
      for (const content of run.content) if (content.kind === "text") texts.push(content.text);
      if (authors) for (const change of run.changes ?? []) texts.push(change.author);
    }
  }
  if (authors) for (const change of block.markChanges ?? []) texts.push(change.author);
  for (const text of texts) {
    const script = unsupportedScript(text);
    if (script) return { script, sample: sample(text) };
  }
  return undefined;
}

/** Whether a block holds a tracked change parsed for print (see Revision). */
function blockHasChanges(block: Block): boolean {
  if (block.type !== "paragraph") return block.rows.some((row) => row.cells.some((cell) => cell.blocks.some(blockHasChanges)));
  return !!block.markChanges?.length || block.children.some((child) => (child.type === "run" ? [child] : child.runs).some((run) => run.changes?.length));
}

const CHANGE_ELEMENTS = new Set(["ins", "del", "moveTo", "moveFrom", "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange"]);

/** Whether a part holds a tracked change. */
/** The ids of the comments `el` marks with a range start or a reference, onto `out`. */
function commentMarkers(el: XmlElement, out: Set<string>): void {
  const name = localName(el.name);
  if (name === "commentRangeStart" || name === "commentReference") {
    const id = attr(el, "id");
    if (id !== undefined) out.add(id);
  }
  for (const child of el.children) commentMarkers(child, out);
}

function xmlHasChanges(el: XmlElement): boolean {
  return CHANGE_ELEMENTS.has(localName(el.name)) || el.children.some(xmlHasChanges);
}

/** The first unsupported script in a part's w:t elements; with markup, in
 * its w:delText elements and, when the balloons name them, in the authors
 * of its changes (w:author) too. */
function unsupportedInXml(el: XmlElement, markup: boolean, authors: boolean): { script: string; sample: string } | undefined {
  const name = localName(el.name);
  if (name === "t" || (markup && name === "delText")) {
    const script = unsupportedScript(el.text);
    return script ? { script, sample: sample(el.text) } : undefined;
  }
  const author = markup && authors ? attr(el, "author") : undefined;
  if (author) {
    const script = unsupportedScript(author);
    if (script) return { script, sample: sample(author) };
  }
  for (const child of el.children) {
    const found = unsupportedInXml(child, markup, authors);
    if (found) return found;
  }
  return undefined;
}

interface Count {
  result: StreamLayoutResult;
  knownTotalPages: number | undefined;
  /** The largest page's sides: the least size of a sheet with a comment column. */
  maxWidth: number;
  maxHeight: number;
}

/** The counting pass, and the NUMPAGES convergence passes layoutDocument
 * runs; `undrawn` sees the pages of each, so it ends with the last's. */
async function countPages(opened: StreamingDocument, measurer: NodeMeasurer, undrawn: UndrawnComments | null): Promise<Count> {
  let maxWidth = 0;
  let maxHeight = 0;
  const count = (total: number | undefined): Promise<StreamLayoutResult> => {
    clearBreakCache(measurer);
    undrawn?.reset();
    return layoutDocumentStream(opened.doc, {
      measurer,
      prescan: opened.prescan,
      knownTotalPages: total,
      onPage: (page) => {
        maxWidth = Math.max(maxWidth, page.width);
        maxHeight = Math.max(maxHeight, page.height);
        undrawn?.visit(page);
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
