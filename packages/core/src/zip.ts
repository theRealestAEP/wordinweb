import { unzipSync, strFromU8 } from "fflate";
import { DEFAULT_XML_LIMITS, countElements, parseXml, rejectXmlDeclarations } from "./xml.js";

/** Fixed mtime for every zip entry we write (2020-01-01T00:00:00Z). Without
 * it fflate stamps `new Date()` at 2-second DOS granularity, so two save()
 * calls straddling a boundary differ in bytes — breaking byte-identity
 * checks (docHash self-heal, checkpoint dedup, byte-compare tests). B11. */
export const FIXED_ZIP_MTIME = new Date(1577836800000);

/** Decompression-bomb guards (plan doc 11, security F4): a crafted .docx must
 * not exhaust memory via a huge uncompressed payload or a vast number of
 * entries. Generous for real documents. */
export const DEFAULT_ZIP_LIMITS = {
  maxTotalUncompressed: 512 * 1024 * 1024, // 512 MB across all entries
  maxEntries: 20_000,
};

/** A part that is XML: inflated whole, decoded and parsed, so parseXml's
 * DEFAULT_XML_LIMITS.maxLength holds it. */
export const XML_PART = /\.(xml|rels|vml)$/i;

/**
 * What one document may hold, so that a document at every limit opens
 * (DocxDocument.load) and renders (renderDocxToPdf) in a process with a
 * 1024 MB heap. Each limit is counted from the zip's directory, or from the
 * XML's bytes before it is parsed, or as the body streams. The largest each
 * was in the test corpora (the render-eval corpus's 1,826 .docx files, and
 * the legacy set's 860 .doc and .docx files) is in brackets; the costs are
 * what probes at each limit measured.
 */
export const DOCUMENT_LIMITS = {
  /** Every XML part but the main document part (styles, numbering,
   * headers, footers, notes, comments, settings, theme, charts),
   * together, in bytes [1.0 MB] ... */
  heldBytes: 8 << 20,
  /** ... in elements [38,000] ... */
  heldElements: 150_000,
  /** ... and in paragraphs and table cells [790]: a header's paragraph
   * costs about 3 KB in the render. */
  heldBlocks: 10_000,
  /** The main document part's elements [558,000]: about 0.5 KB each loaded,
   * and about 1 KB each rendered when they are text runs ... */
  documentElements: 1_000_000,
  /** ... and its paragraphs and table cells [23,000]: 250,000 empty ones
   * load in 470 MB and render in 465 MB. */
  documentBlocks: 250_000,
  /** One block of the body (a paragraph or table) as the streaming render
   * holds it, in bytes [2.0 MB] ... */
  blockBytes: 8 << 20,
  /** ... in bytes of text [255 KB]: a paragraph's text costs the render
   * about 500 bytes a byte as it is laid out ... */
  blockText: 1 << 20,
  /** ... in elements [69,000] ... */
  blockElements: 250_000,
  /** ... and in table rows, nested tables' included [407]: the render's
   * time grows with the square of a table's rows, 20 seconds at the limit. */
  tableRows: 20_000,
};

/**
 * Holds a package's entries to DEFAULT_ZIP_LIMITS, each XML part to
 * DEFAULT_XML_LIMITS.maxLength, and the XML parts but the main document
 * part to DOCUMENT_LIMITS.heldBytes, from the sizes the zip's directory
 * gives, before any entry is inflated. Bytes are compared with the XML
 * limit, which counts characters: UTF-8 has at least one byte a character,
 * so this check is the stricter one.
 */
export class ZipBudget {
  private entries = 0;
  private total = 0;
  /** Each XML part's size. */
  private readonly xml = new Map<string, number>();

  add(name: string, size: number): void {
    if (++this.entries > DEFAULT_ZIP_LIMITS.maxEntries) throw new Error(`zip has too many entries (> ${DEFAULT_ZIP_LIMITS.maxEntries})`);
    if (XML_PART.test(name) && size > DEFAULT_XML_LIMITS.maxLength) throw new Error(`too large: ${name} is ${size} bytes, over the ${DEFAULT_XML_LIMITS.maxLength} an XML part may be`);
    if (XML_PART.test(name)) this.xml.set(name, size);
    this.total += size;
    if (this.total > DEFAULT_ZIP_LIMITS.maxTotalUncompressed) throw new Error(`too large: the zip's parts come to over ${DEFAULT_ZIP_LIMITS.maxTotalUncompressed} bytes`);
  }

  /** Hold every XML part but `docPart` to DOCUMENT_LIMITS.heldBytes, together. */
  holdParts(docPart: string): void {
    let held = 0;
    for (const [name, size] of this.xml) if (name !== docPart) held += size;
    if (held > DOCUMENT_LIMITS.heldBytes) throw new Error(`too large: the XML parts other than ${docPart} come to ${held} bytes, over the ${DOCUMENT_LIMITS.heldBytes} they may be together`);
  }

  /** An XML part's size, as added. */
  xmlSize(name: string): number | undefined {
    return this.xml.get(name);
  }
}

/** Hold every entry of the zip `data` to ZipBudget, from its directory: an
 * unzip whose filter takes no entry inflates nothing. Only _rels/.rels is
 * inflated, held to DOCUMENT_LIMITS.heldBytes first, to find the main
 * document part. */
export function budgetZip(data: Uint8Array): void {
  const budget = new ZipBudget();
  unzipSync(data, {
    filter: (file) => {
      budget.add(file.name, file.originalSize);
      return false;
    },
  });
  const relsSize = budget.xmlSize("_rels/.rels") ?? 0;
  if (relsSize > DOCUMENT_LIMITS.heldBytes) throw new Error(`too large: _rels/.rels is ${relsSize} bytes, over the ${DOCUMENT_LIMITS.heldBytes} the XML parts other than the document may be`);
  const rels = relsSize > 0 ? unzipSync(data, { filter: (file) => file.name === "_rels/.rels" })["_rels/.rels"] : undefined;
  budget.holdParts(mainDocumentPart(rels));
}

/**
 * The main document part, as the package relationships (`rels`, the bytes
 * of _rels/.rels) name it; word/document.xml when they name none. Their
 * elements are held to DOCUMENT_LIMITS.heldElements before they are parsed.
 */
export function mainDocumentPart(rels: Uint8Array | undefined): string {
  if (rels) {
    const count = countElements(rels).elements;
    if (count > DOCUMENT_LIMITS.heldElements) throw new Error(`too large: _rels/.rels holds ${count} elements, over the ${DOCUMENT_LIMITS.heldElements} the XML parts other than the document may hold`);
    for (const rel of parseXml(strFromU8(rels)).children) {
      const target = rel.attrs.Target;
      if (rel.attrs.TargetMode !== "External" && target && (rel.attrs.Type ?? "").endsWith("/officeDocument")) return resolvePartPath("", target);
    }
  }
  return "word/document.xml";
}

/**
 * Hold a package's XML parts to DOCUMENT_LIMITS, from their bytes, before
 * any is parsed: the elements, paragraphs and table cells of every part
 * but `docPart` together, and those of `docPart`.
 */
export function holdXmlCounts(files: Record<string, Uint8Array>, docPart: string): void {
  let elements = 0;
  let blocks = 0;
  for (const [name, bytes] of Object.entries(files)) {
    if (!XML_PART.test(name)) continue;
    rejectXmlDeclarations(strFromU8(bytes));
    const count = countElements(bytes);
    if (name === docPart) {
      if (count.elements > DOCUMENT_LIMITS.documentElements) throw new Error(`too large: ${name} holds ${count.elements} elements, over the ${DOCUMENT_LIMITS.documentElements} it may hold`);
      if (count.blocks > DOCUMENT_LIMITS.documentBlocks) throw new Error(`too large: ${name} holds ${count.blocks} paragraphs and table cells, over the ${DOCUMENT_LIMITS.documentBlocks} it may hold`);
    } else {
      elements += count.elements;
      blocks += count.blocks;
    }
  }
  if (elements > DOCUMENT_LIMITS.heldElements) throw new Error(`too large: the XML parts other than ${docPart} hold ${elements} elements, over the ${DOCUMENT_LIMITS.heldElements} they may hold together`);
  if (blocks > DOCUMENT_LIMITS.heldBlocks) throw new Error(`too large: the XML parts other than ${docPart} hold ${blocks} paragraphs and table cells, over the ${DOCUMENT_LIMITS.heldBlocks} they may hold together`);
}

/** A read-only view over the OPC (zip) package inside a .docx file. */
export class Package {
  private files: Record<string, Uint8Array>;

  /** Every entry is held to ZipBudget before any is inflated. fflate
   * inflates an entry into a buffer of the size the directory gives, so an
   * entry whose data holds more stops there: memory stays within the
   * budget, and the entry fails, or comes out cut short (an XML part then
   * fails to parse). */
  constructor(data: Uint8Array) {
    budgetZip(data);
    let name = "";
    try {
      this.files = unzipSync(data, {
        filter: (file) => {
          name = file.name;
          return true;
        },
      });
    } catch (error) {
      throw name ? new Error(`${name} is corrupt`) : error;
    }
  }

  static from(data: ArrayBuffer | Uint8Array): Package {
    return new Package(data instanceof Uint8Array ? data : new Uint8Array(data));
  }

  /** A package over an already-unzipped entry map (DocxDocument.clone). */
  static fromFiles(files: Record<string, Uint8Array>): Package {
    const pkg = Object.create(Package.prototype) as Package;
    pkg.files = files;
    return pkg;
  }

  has(name: string): boolean {
    return this.normalize(name) in this.files;
  }

  binary(name: string): Uint8Array | undefined {
    return this.files[this.normalize(name)];
  }

  text(name: string): string | undefined {
    const bin = this.binary(name);
    return bin ? strFromU8(bin) : undefined;
  }

  names(): string[] {
    return Object.keys(this.files);
  }

  /** Raw entry map (shared, do not mutate) — used for write-back. */
  raw(): Record<string, Uint8Array> {
    return this.files;
  }

  /** OPC part names never start with '/' inside the zip; tolerate both. */
  private normalize(name: string): string {
    return name.startsWith("/") ? name.slice(1) : name;
  }
}

/** Resolve a relationship target relative to a source part. */
/** The relationship Target that reaches `toPart` from `fromPart`'s folder. */
export function relativePartPath(fromPart: string, toPart: string): string {
  const fromDir = fromPart.split("/").slice(0, -1);
  const to = toPart.split("/");
  let i = 0;
  while (i < fromDir.length && i < to.length - 1 && fromDir[i] === to[i]) i++;
  return [...fromDir.slice(i).map(() => ".."), ...to.slice(i)].join("/");
}

export function resolvePartPath(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const baseDir = sourcePart.includes("/")
    ? sourcePart.slice(0, sourcePart.lastIndexOf("/"))
    : "";
  const segments = (baseDir ? baseDir + "/" + target : target).split("/");
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") out.pop();
    else out.push(seg);
  }
  return out.join("/");
}

/**
 * Each entry's uncompressed size, read from the zip's central directory
 * without inflating anything. Throws when the bytes are not a zip.
 */
export function zipEntrySizes(data: Uint8Array): Map<string, number> {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  // The end-of-central-directory record closes the file, after a comment of
  // at most 65535 bytes.
  let end = data.length - 22;
  const stop = Math.max(0, data.length - 22 - 65535);
  while (end >= stop && view.getUint32(end, true) !== 0x06054b50) end--;
  if (end < stop) throw new Error("not a zip: no end-of-central-directory record");
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const sizes = new Map<string, number>();
  const decoder = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (at + 46 > data.length || view.getUint32(at, true) !== 0x02014b50) throw new Error("not a zip: bad central directory");
    const size = view.getUint32(at + 24, true);
    const nameLength = view.getUint16(at + 28, true);
    const extra = view.getUint16(at + 30, true);
    const comment = view.getUint16(at + 32, true);
    // 0xFFFFFFFF marks a ZIP64 size: over 4 GB, far past any limit.
    sizes.set(decoder.decode(data.subarray(at + 46, at + 46 + nameLength)), size === 0xffffffff ? Infinity : size);
    at += 46 + nameLength + extra + comment;
  }
  return sizes;
}
