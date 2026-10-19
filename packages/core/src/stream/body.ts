import { Inflate, deflateSync, inflateSync, strFromU8, strToU8 } from "fflate";
import { DocxDocument } from "../docx.js";
import { StrictRewriter, normalizeStrictXml } from "../strict.js";
import { Prescan, bodyHasPageTotal } from "../layout/engine.js";
import { Block, FieldContent, Paragraph, Revision, Section, SectionProps } from "../model.js";
import { DocParseContext, formatChange, inheritHeaderFooterRefs, parseParagraph, parseTable } from "../parse/document.js";
import { parseRelationships } from "../parse/rels.js";
import { defaultSectionProps, parseSectionProps } from "../parse/section.js";
import { XmlElement, attr, child, localName, parseXml } from "../xml.js";
import { DEFAULT_ZIP_LIMITS } from "../zip.js";
import { scanBody } from "./scan.js";
import { ByteSource } from "./source.js";
import { ZipIndex } from "./zip.js";

/** A document whose body stays out of memory: the engine reads its blocks
 * through `doc.sections`, which parse them from the deflated body as it goes
 * and keep only the ones around its cursor. */
export interface StreamingDocument {
  doc: DocxDocument;
  /** The document-order assignments every layout pass shares. */
  prescan: Prescan;
  /** Whether the body holds a NUMPAGES field, which layoutDocument converges on. */
  bodyHasPageTotal: boolean;
  /** A binary part's bytes, read from the file when asked. */
  media(part: string): Promise<Uint8Array | undefined>;
  /** A binary part's bytes a piece at a time, never held whole; false when
   * the part is missing. A promise `onPiece` gives back holds the reading. */
  mediaPieces(part: string, onPiece: (piece: Uint8Array) => void | Promise<void>): Promise<boolean>;
  /** A binary part's size in bytes, from the zip directory; undefined when missing. */
  mediaSize(part: string): number | undefined;
  /** The ids of the comments whose ranges cover a body w:t or w:delText (the element a
   * TextItem's src.t is) or an equation (the m:oMath or m:oMathPara a
   * TextItem's mathSrc is), as DocxDocument.commentAnchors gives them for a
   * document loaded whole; undefined when none does. */
  commentIds(t: XmlElement): string[] | undefined;
  /** Where a body w:t, w:delText or equation is: its block's index and its
   * place among the block's texts, so the nearest drawn text to one that is
   * not drawn can be found. */
  textPosition(t: XmlElement): TextPosition | undefined;
  /** Where a comment is anchored: its first marker (range start or
   * reference) in the body, halfway between the texts before and after it;
   * undefined for a comment the body never marks. */
  commentAnchor(id: string): TextPosition | undefined;
}

export interface TextPosition {
  block: number;
  ordinal: number;
}

/** Parts the shell holds whole; every other part stays in the file. */
const TEXT_PART = /\.(xml|rels|vml)$/i;

/** The body's XML is deflated in pieces of about this much, cut at block boundaries. */
const PIECE_BYTES = 1 << 20;

/** Parsed blocks kept behind the furthest one read, for the engine's lookback. */
const KEEP_BEHIND = 64;
/** The parsed-block cache is trimmed once it holds this many. */
const KEEP_MAX = 512;

/**
 * Open a .docx for layout without its body in memory. The file's small parts
 * make a shell DocxDocument, as compare-cli's readInput does; the body is
 * scanned once, block by block, into deflated pieces that the sections' lazy
 * block arrays parse from on demand, and parsed once more for the section
 * properties and the Prescan.
 */
export interface OpenStreamingOptions {
  /** Sees each body block once, parsed, in the first pass: `index` is the
   * block's position in the body. A caller refuses a document here before
   * any layout. */
  onBlock?: (block: Block, index: number) => void;
  /** Parse the tracked changes for print: every part in the markup view,
   * each author's changes in Word's color for the author (see
   * FromFilesOptions.revisionColors). */
  markup?: boolean;
}

export async function openStreamingDocument(source: ByteSource, options: OpenStreamingOptions = {}): Promise<StreamingDocument> {
  await refuseOtherFormat(source);
  const zip = await ZipIndex.open(source);
  const mimetype = zip.entries.has("mimetype") ? await zip.text("mimetype") : undefined;
  if (mimetype?.startsWith("application/vnd.oasis.opendocument")) throw new Error("unsupported format: an OpenDocument file");
  const docPart = documentPart(await zip.text("_rels/.rels"));
  const docEntry = zip.entries.get(docPart);
  if (!docEntry) throw new Error(`missing ${docPart}`);
  const body = new BodyStore();
  // A Strict Open XML body is rewritten to Transitional as it is scanned (strict.ts).
  const rewriter = new StrictRewriter();
  const outline = body.scan(rewriter.through(inflatePull(docEntry.method, await zip.raw(docEntry))));

  // The shell: every text part, and document.xml with only the body's own
  // section properties. Their sizes, from the directory, are held to what
  // Package allows a document loaded whole, before any of them is inflated.
  const held = [...zip.entries.values()].filter((entry) => entry.name !== docPart && !entry.name.endsWith("/") && TEXT_PART.test(entry.name));
  const heldBytes = held.reduce((total, entry) => total + entry.size, 0);
  if (heldBytes > DEFAULT_ZIP_LIMITS.maxTotalUncompressed) throw new Error(`the parts held in memory come to ${heldBytes} bytes, over ${DEFAULT_ZIP_LIMITS.maxTotalUncompressed}`);
  const files: Record<string, Uint8Array> = {};
  for (const entry of held) {
    const bytes = await zip.read(entry);
    files[entry.name] = rewriter.strict ? strToU8(normalizeStrictXml(strFromU8(bytes))) : bytes;
  }
  files[docPart] = outline;
  const doc = DocxDocument.fromFiles(files, options.markup ? { revisionView: "markup", revisionColors: true } : {});
  body.hydrate = (el) => doc.hydrateControls(el);

  // Every block once, in order: which candidates are blocks at all, where
  // each section ends, and what the Prescan assigns.
  const ctx = doc.bodyParseContext();
  doc.refBookmarks = ctx.refBookmarks!.byName;
  const prescan = new Prescan(doc, (field) => body.fieldKeys.get(field));
  const sections: { props: SectionProps; from: number; to: number }[] = [];
  let from = 0;
  let hasPageTotal = false;
  // Comment ranges cross blocks: this pass learns which are open when each
  // block starts, so a later parse of the block can tag its w:t elements.
  const comments = doc.comments.length > 0 ? new CommentRanges() : null;
  const endSection = (props: SectionProps): void => {
    sections.push({ props, from, to: body.count });
    from = body.count;
    prescan.endSection();
  };
  // Printed markup: the body's own section change is anchored at the last
  // paragraph's mark, as Word balloons it.
  let finalSection: { index: number; change: Revision } | null = null;
  for (let candidate = 0; candidate < body.candidates; candidate++) {
    const el = body.element(candidate);
    if (body.kind(candidate) === "sectPr") {
      endSection(parseSectionProps(el));
      const sectPrChange = ctx.revisionColor ? child(el, "sectPrChange") : undefined;
      if (sectPrChange) finalSection = { index: body.count - 1, change: formatChange(sectPrChange, "Section") };
      continue;
    }
    const block = parseBlock(el, ctx);
    if (!block) continue;
    body.keep(candidate, block);
    options.onBlock?.(block, body.count - 1);
    comments?.first(el, body.count - 1);
    if (!hasPageTotal && bodyHasPageTotal(el)) hasPageTotal = true;
    prescan.visit(block);
    if (block.type === "paragraph" && block.sectionBreak) endSection(block.sectionBreak);
  }
  if (body.count > from) endSection(defaultSectionProps());
  prescan.finish(sections.map((section) => section.props));
  body.forget();

  // Later parses must not touch the captured REF bookmark ranges, which this
  // pass completed: a block parsed again out of order would extend them.
  const later: DocParseContext = { ...ctx, refBookmarks: undefined };
  body.parser = (el, index) => {
    comments?.again(el, index);
    const block = parseBlock(el, later) as Block;
    if (finalSection && index === finalSection.index) {
      const last = lastParagraph(block);
      if (last) last.markChanges = [...(last.markChanges ?? []), finalSection.change];
    }
    return block;
  };
  doc.sections = sections.map((section): Section => ({ props: section.props, blocks: body.blocks(section.from, section.to) }));
  inheritHeaderFooterRefs(doc.sections);

  return {
    doc,
    prescan,
    bodyHasPageTotal: hasPageTotal,
    async media(part) {
      const entry = zip.entries.get(part.startsWith("/") ? part.slice(1) : part);
      return entry ? zip.read(entry) : undefined;
    },
    async mediaPieces(part, onPiece) {
      const entry = zip.entries.get(part.startsWith("/") ? part.slice(1) : part);
      if (!entry) return false;
      await zip.pieces(entry, onPiece);
      return true;
    },
    mediaSize(part) {
      return zip.entries.get(part.startsWith("/") ? part.slice(1) : part)?.size;
    },
    commentIds(t) {
      return comments?.ids.get(t);
    },
    textPosition(t) {
      return comments?.positions.get(t);
    },
    commentAnchor(id) {
      return comments?.anchors.get(id);
    },
  };
}

/**
 * Which comments cover each w:t of a streamed body. The first pass walks
 * the blocks in order and records the comments open at each block's start;
 * every later parse of a block replays its markers from that state and
 * tags the w:t and w:delText elements it made. A point comment (a reference with no
 * range) anchors to the nearest w:t before it in its block, else the first
 * after, as commentAnchors anchors it to the one before it in the document.
 *
 * A text inside an equation (m:t) tags the equation too: the layout's items
 * for an equation name the m:oMath or m:oMathPara, not its texts. Every
 * text is also given its position (block and ordinal), and each comment
 * the position of its first marker, so the card of a comment whose text is
 * never drawn can go beside the nearest text that is.
 */
class CommentRanges {
  readonly ids = new WeakMap<XmlElement, string[]>();
  readonly positions = new WeakMap<XmlElement, TextPosition>();
  readonly anchors = new Map<string, TextPosition>();
  private readonly openAtStart = new Map<number, string[]>();
  private readonly open = new Set<string>();

  first(el: XmlElement, index: number): void {
    if (this.open.size > 0) this.openAtStart.set(index, [...this.open]);
    this.walk(el, index, this.open, true);
  }

  again(el: XmlElement, index: number): void {
    this.walk(el, index, new Set(this.openAtStart.get(index)), false);
  }

  private walk(block: XmlElement, index: number, open: Set<string>, first: boolean): void {
    let lastT: XmlElement | null = null;
    let lastMath: XmlElement | null = null;
    let firstT: XmlElement | null = null;
    let firstMath: XmlElement | null = null;
    /** The equation being walked, when one is. */
    let math: XmlElement | null = null;
    let ordinal = 0;
    const pointIds: string[] = [];
    const tag = (el: XmlElement, id: string): void => {
      const list = this.ids.get(el);
      if (!list) this.ids.set(el, [id]);
      else if (!list.includes(id)) list.push(id);
    };
    const tagText = (t: XmlElement, equation: XmlElement | null, id: string): void => {
      tag(t, id);
      if (equation) tag(equation, id);
    };
    /** A comment's anchor: its first marker, between the texts before and after it. */
    const anchor = (id: string): void => {
      if (first && !this.anchors.has(id)) this.anchors.set(id, { block: index, ordinal: ordinal - 0.5 });
    };
    const visit = (el: XmlElement): void => {
      const ln = localName(el.name);
      if (ln === "commentRangeStart") {
        const id = attr(el, "id");
        if (id === undefined) return;
        open.add(id);
        anchor(id);
        return;
      }
      if (ln === "commentRangeEnd") {
        const id = attr(el, "id");
        if (id !== undefined) open.delete(id);
        return;
      }
      if (ln === "commentReference") {
        const id = attr(el, "id");
        if (id === undefined) return;
        anchor(id);
        if (lastT) tagText(lastT, lastMath, id);
        else pointIds.push(id);
        return;
      }
      // Deleted text (w:delText) prints in the markup view, its comments with it.
      if (ln === "t" || ln === "delText") {
        const position = { block: index, ordinal: ordinal++ };
        this.positions.set(el, position);
        if (math && !this.positions.has(math)) this.positions.set(math, position);
        lastT = el;
        lastMath = math;
        if (!firstT) {
          firstT = el;
          firstMath = math;
        }
        for (const id of open) tagText(el, math, id);
        return;
      }
      if ((ln === "oMath" || ln === "oMathPara") && !math) {
        math = el;
        for (const c of el.children) visit(c);
        math = null;
        return;
      }
      for (const c of el.children) visit(c);
    };
    visit(block);
    if (firstT) for (const id of pointIds) tagText(firstT, firstMath, id);
  }
}

/** The last paragraph in a block: the block itself, or the last one in a
 * table's last cell (a nested table's included); null when there is none. */
function lastParagraph(block: Block): Paragraph | null {
  if (block.type === "paragraph") return block;
  const cells = block.rows.at(-1)?.cells ?? [];
  for (const nested of [...(cells.at(-1)?.blocks ?? [])].reverse()) {
    const found = lastParagraph(nested);
    if (found) return found;
  }
  return null;
}

/** A body's block as parse/document.ts's parseBody makes it: null when the
 * element is no block (a paragraph deleted in the final view, a table whose
 * rows all are). */
function parseBlock(el: XmlElement, ctx: DocParseContext): Block | null {
  const ln = localName(el.name);
  if (ln === "p") {
    const para = parseParagraph(el, ctx);
    return para.revisionHidden ? null : para;
  }
  if (ln === "tbl") {
    const table = parseTable(el, ctx);
    return table.rows.length > 0 ? table : null;
  }
  return null;
}

/**
 * A file that is another format altogether is refused from its first bytes,
 * before the zip directory is looked for: an OLE compound file (a Word
 * 97-2003 .doc, or a password-protected .docx, whose package is a stream
 * inside one), an RTF file, a PDF.
 */
async function refuseOtherFormat(source: ByteSource): Promise<void> {
  const size = await source.size();
  const head = await source.read(0, Math.min(size, 8));
  const ascii = String.fromCharCode(...head);
  if (ascii.startsWith("{\\rtf")) throw new Error("unsupported format: an RTF file");
  if (ascii.startsWith("%PDF")) throw new Error("unsupported format: a PDF file");
  if (head.length === 8 && [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, i) => head[i] === byte)) {
    // An encrypted package names its EncryptionInfo stream in the directory,
    // which is within the first sectors of any file Word writes.
    const front = await source.read(0, Math.min(size, 64 << 10));
    const names = String.fromCharCode(...front.filter((_, i) => i % 2 === 0));
    const encrypted = names.includes("EncryptionInfo") || names.includes("EncryptedPackage");
    throw new Error(`unsupported format: ${encrypted ? "a password-protected document (its package is encrypted)" : "an OLE compound file (a Word 97-2003 .doc)"}`);
  }
}

/** The main document part, as the package relationships name it. */
function documentPart(rels: string | undefined): string {
  if (rels) {
    for (const rel of parseRelationships(parseXml(rels), "").values()) {
      if (rel.type.endsWith("/officeDocument")) return rel.target;
    }
  }
  return "word/document.xml";
}

/**
 * The body's blocks, held deflated and parsed on demand. A scan first finds
 * every candidate block element (w:p, w:tbl, the body's w:sectPr) and writes
 * their bytes, in order, into pieces deflated apart, so any block is one
 * piece's inflate away; a parse then decides which candidates are blocks.
 */
class BodyStore {
  /** Each candidate's [start, end) in the concatenated candidate bytes, as pairs. */
  private readonly ranges: number[] = [];
  /** Each candidate's local name: p, tbl, or the body's own sectPr. */
  private readonly kinds: string[] = [];
  /** Each block-level w:sdt's properties, as its start tag through its
   * sdtContent's, and the sdt it sits in; each candidate's innermost one. */
  private readonly sdts: { parent: number; xml: string }[] = [];
  private readonly sdtOf: number[] = [];
  /** The candidate bytes, deflated in pieces: piece i holds [starts[i], starts[i + 1]). */
  private readonly pieces: Uint8Array[] = [];
  private readonly starts: number[] = [];
  /** Each block's candidate index. */
  private readonly kept: number[] = [];
  /** Parsed blocks, by block index, around the engine's cursor. */
  private parsed = new Map<number, Block>();
  private furthest = -1;
  /** The piece inflated last. */
  private inflated: { index: number; bytes: Uint8Array } | null = null;
  /** The stable key of each field occurrence the parses made (see Prescan). */
  readonly fieldKeys = new WeakMap<object, string>();
  /** Parses a block's element; `index` is the block's position in the body. */
  parser: ((el: XmlElement, index: number) => Block) | null = null;
  /** Resolves a block's content controls as the shell resolves its own (DocxDocument.hydrateControls). */
  hydrate: ((el: XmlElement) => void) | null = null;

  get candidates(): number {
    return this.kinds.length;
  }

  get count(): number {
    return this.kept.length;
  }

  /** Scan the document part, pulled a piece at a time; returns the part without the body's children. */
  scan(pull: () => Uint8Array | null): Uint8Array {
    let piece: Uint8Array[] = [];
    let pieceBytes = 0;
    let total = 0;
    const flush = (): void => {
      if (piece.length === 0) return;
      const bytes = new Uint8Array(pieceBytes);
      let at = 0;
      for (const part of piece) {
        bytes.set(part, at);
        at += part.length;
      }
      this.pieces.push(deflateSync(bytes, { level: 1 }));
      piece = [];
      pieceBytes = 0;
    };
    const outline = scanBody(
      pull,
      (child) => {
        const ln = localName(child.name);
        // A sectPr counts at the body's own level only, as parseBody walks it.
        if (ln !== "p" && ln !== "tbl" && (ln !== "sectPr" || !child.top)) return;
        if (pieceBytes >= PIECE_BYTES) flush();
        if (piece.length === 0) this.starts.push(total);
        piece.push(child.bytes.slice());
        pieceBytes += child.bytes.length;
        this.ranges.push(total, total + child.bytes.length);
        this.kinds.push(ln);
        this.sdtOf.push(child.sdt);
        total += child.bytes.length;
      },
      (sdt) => {
        this.sdts.push({ parent: sdt.parent, xml: strFromU8(sdt.bytes) });
      },
    );
    flush();
    this.starts.push(total);
    const head = outline.head;
    const sectPr = this.kinds.lastIndexOf("sectPr");
    const own = sectPr === -1 ? null : this.candidateBytes(sectPr);
    const shell = new Uint8Array(head.length + (own?.length ?? 0) + outline.tail.length);
    shell.set(head);
    if (own) shell.set(own, head.length);
    shell.set(outline.tail, head.length + (own?.length ?? 0));
    return shell;
  }

  kind(candidate: number): string {
    return this.kinds[candidate];
  }

  /**
   * A candidate's element, its content controls resolved. The block-level
   * sdts around a block are not in its bytes, so a block in one is parsed
   * inside them again: a control bound to a document property is resolved
   * from the sdt's properties, as the shell resolves its own.
   */
  element(candidate: number): XmlElement {
    const xml = strFromU8(this.candidateBytes(candidate));
    const wrappers: string[] = [];
    for (let sdt = this.sdtOf[candidate]; sdt !== -1; sdt = this.sdts[sdt].parent) wrappers.unshift(this.sdts[sdt].xml);
    let el = parseXml(wrappers.join("") + xml + "</w:sdtContent></w:sdt>".repeat(wrappers.length));
    this.hydrate?.(el);
    for (let i = 0; i < wrappers.length; i++) el = (child(el, "sdtContent") as XmlElement).children[0];
    return el;
  }

  private candidateBytes(candidate: number): Uint8Array {
    const start = this.ranges[2 * candidate];
    const end = this.ranges[2 * candidate + 1];
    let index = this.inflated?.index ?? -1;
    if (index === -1 || this.starts[index] > start || this.starts[index + 1] <= start) {
      index = this.starts.findIndex((at, i) => at <= start && this.starts[i + 1] > start);
      const bytes = inflateSync(this.pieces[index], { out: new Uint8Array(this.starts[index + 1] - this.starts[index]) });
      this.inflated = { index, bytes };
    }
    const bytes = (this.inflated as { bytes: Uint8Array }).bytes;
    return bytes.subarray(start - this.starts[index], end - this.starts[index]);
  }

  /** Candidate `candidate` is a block: the next one, parsed as `block`. */
  keep(candidate: number, block: Block): void {
    this.parsed.set(this.kept.length, block);
    this.kept.push(candidate);
    keyFields(block, this.kept.length - 1, this.fieldKeys);
    this.trim(this.kept.length - 1);
  }

  /** Drop the parsed blocks, after the first pass over them. */
  forget(): void {
    this.parsed = new Map();
    this.furthest = -1;
  }

  /** Blocks [from, to) as an array that parses each block when it is read. */
  blocks(from: number, to: number): Block[] {
    const target: Block[] = new Array(to - from);
    const inRange = (prop: string | symbol): number => {
      if (typeof prop !== "string") return -1;
      const i = Number(prop);
      return Number.isInteger(i) && i >= 0 && i < target.length ? i : -1;
    };
    return new Proxy(target, {
      get: (t, prop, receiver) => {
        const i = inRange(prop);
        return i === -1 ? Reflect.get(t, prop, receiver) : this.block(from + i);
      },
      has: (t, prop) => inRange(prop) !== -1 || Reflect.has(t, prop),
    });
  }

  private block(index: number): Block {
    let block = this.parsed.get(index);
    if (block) return block;
    block = (this.parser as (el: XmlElement, index: number) => Block)(this.element(this.kept[index]), index);
    this.parsed.set(index, block);
    keyFields(block, index, this.fieldKeys);
    this.trim(index);
    return block;
  }

  /** Keep the parsed blocks near the furthest one read: the engine looks a
   * bounded number of blocks ahead and behind its cursor. */
  private trim(index: number): void {
    if (index > this.furthest) this.furthest = index;
    if (this.parsed.size < KEEP_MAX) return;
    for (const key of this.parsed.keys()) if (key < this.furthest - KEEP_BEHIND) this.parsed.delete(key);
  }
}

/** Give each field occurrence in `block` the key the Prescan's memos use for it. */
function keyFields(block: Block, index: number, keys: WeakMap<object, string>): void {
  let n = 0;
  const visit = (b: Block): void => {
    if (b.type !== "paragraph") {
      for (const row of b.rows) for (const cell of row.cells) for (const nested of cell.blocks) visit(nested);
      return;
    }
    for (const c of b.children) {
      const runs = c.type === "run" ? [c] : c.runs;
      for (const r of runs) for (const rc of r.content) if (rc.kind === "field") keys.set(rc as FieldContent, `${index}:${n++}`);
    }
  };
  visit(block);
}

/** The inflated part, a chunk per call; null once it is all out. */
function inflatePull(method: number, raw: Uint8Array): () => Uint8Array | null {
  if (method !== 8) {
    let done = false;
    return () => {
      if (done) return null;
      done = true;
      return raw;
    };
  }
  const out: Uint8Array[] = [];
  const inflater = new Inflate((chunk) => out.push(chunk));
  let at = 0;
  return () => {
    while (out.length === 0) {
      if (at >= raw.length) return null;
      const next = Math.min(at + (64 << 10), raw.length);
      inflater.push(raw.subarray(at, next), next === raw.length);
      at = next;
    }
    return out.shift() as Uint8Array;
  };
}
