import { Hash, createHash } from "node:crypto";
import { closeSync, openSync, readSync, fstatSync } from "node:fs";
import { join } from "node:path";
import { DocxDocument } from "../../core/src/docx.js";
import { DEFAULT_XML_LIMITS, XmlElement, child, cyrb53, localName, parseXml, serializeXml } from "../../core/src/xml.js";
import { acceptAllRevisions, collectRevisions } from "../../core/src/edit/suggest.js";
import { paragraphText } from "../../core/src/edit/compare/content.js";
import { ROW_LEVEL, blockFingerprint, rowFingerprint, sectionReferences } from "../../core/src/edit/compare/index.js";
import { normalizeText } from "../../core/src/edit/compare/text.js";
import { resolvePartPath } from "../../core/src/zip.js";
import { altChunkError, subDocError } from "../../core/src/parse/document.js";
import { parseRelationships, relsPathFor } from "../../core/src/parse/rels.js";
import { CompareError } from "./errors.js";
import { DiskNotes } from "./notes.js";
import { scanBody, scanChildren } from "./scan.js";
import { BlockStore } from "./store.js";
import { HEAP_MB, MAX_HELD_BYTES, ZipEntry, ZipReader, mb } from "./zip.js";

export type Side = "original" | "revised";

/**
 * A binary part's stand-in inside a shell: the shell needs to know the part
 * exists, not what it holds. The output writer copies the real bytes from
 * the input zip the stub names.
 */
const STUB = "\u0000dxw-stub\u0000";

/** The input zip and part name a stub stands for, or null for real content. */
export function stubSource(bytes: Uint8Array): { side: Side; name: string } | null {
  if (bytes.length < STUB.length || bytes[0] !== 0 || Buffer.from(bytes.subarray(0, STUB.length)).toString("latin1") !== STUB) return null;
  const [side, name] = Buffer.from(bytes.subarray(STUB.length)).toString("utf8").split("\u0000");
  return { side: side as Side, name };
}

/** Tracked-change markup. */
const TRACKED_CHANGE =
  /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|trPrChange|tcPrChange|numberingChange|cellIns|cellDel|cellMerge)[\s>/]/;

/** Blocks with tracked changes are accepted in batches of about this much XML. */
const BATCH_BYTES = 4 << 20;

/** Stands where a row sits in a table's shell. */
export const ROW_SLOT: XmlElement = { name: "w:dxw-row", attrs: {}, children: [], text: "" };

/**
 * One input document: its body on disk, block by block, and everything else
 * in a shell. A table's rows are blocks of their own: the table's block is
 * its shell, the table element with a ROW_SLOT where each row sits.
 */
export interface Input {
  side: Side;
  zip: ZipReader;
  docPart: string;
  /**
   * The document with every part but the body: the body holds only its own
   * section properties, and each binary part is a stub.
   */
  shell: DocxDocument;
  /** Each body block as XML, with tracked changes accepted. */
  blocks: BlockStore;
  /** Each block's paragraphText; a table's is "", read from its rows when wanted. */
  texts: BlockStore;
  /** Each block's blockFingerprint, hashed to a number: the anchor pass's keys. */
  fingerprints: number[];
  /** A hash of each block's XML, as a number. */
  hashes: number[];
  /** Each block's local name. */
  kinds: string[];
  /** Whether each block is a paragraph that ends a section. */
  endsSection: boolean[];
  /** The sectionReferences of each block that ends a section (pairable). */
  sectionKeys: Map<number, string>;
  /** The w:sectPr of each block that ends a section, in order. */
  sectPrs: XmlElement[];
  /** Every table's row-level children (tr, sdt, customXml) as XML, in document order. */
  rows: BlockStore;
  /** Each row's rowFingerprint, hashed to a number. */
  rowFingerprints: number[];
  /** Each row's local name. */
  rowKinds: string[];
  /** Each table block's rows: [from, to) in `rows`. */
  tables: Map<number, [number, number]>;
  /** Note reference ids in document order. */
  noteRefs: Record<"footnote" | "endnote", string[]>;
  /** Comment reference ids in document order. */
  commentRefs: string[];
  /** The footnotes and endnotes parts, on disk; the shell holds them empty. */
  notes: Record<"footnote" | "endnote", DiskNotes | null>;
  /** Whether the input held tracked changes, now accepted. */
  hadRevisions: boolean;
}

/** Read one input into `dir`. */
export async function readInput(zip: ZipReader, side: Side, dir: string): Promise<Input> {
  const docPart = documentPart(zip);
  const docEntry = zip.entries.get(docPart);
  if (!docEntry) throw new CompareError("not_docx", `${zip.path}: missing ${docPart}`);
  const docFile = join(dir, `${side}-document.xml`);
  await zip.extract(docEntry, docFile);

  // First the body's outline: where it is, and where each child is.
  const starts: number[] = [];
  const ends: number[] = [];
  const names: string[] = [];
  let bounds;
  try {
    bounds = scanBody(docFile, (c) => {
      starts.push(c.start);
      ends.push(c.end);
      names.push(c.name);
    });
  } catch (error) {
    throw withPart(docPart, error);
  }

  const fd = openSync(docFile, "r");
  try {
    // One buffer for every block read, so the reads leave no garbage outside the heap.
    let scratch = Buffer.allocUnsafe(64 * 1024);
    const range = (start: number, end: number): Buffer => {
      checkBlock(docPart, end - start);
      if (end - start > scratch.length) scratch = Buffer.allocUnsafe(Math.max(end - start, scratch.length * 2));
      const buf = scratch.subarray(0, end - start);
      let done = 0;
      while (done < buf.length) done += readSync(fd, buf, done, buf.length - done, start + done);
      return buf;
    };
    const isSectPr = (i: number): boolean => localName(names[i]) === "sectPr";

    // The notes parts go to disk; the shell holds each empty. The comments
    // part is held: struck content may need the original's comments.
    const noteParts: Partial<Record<"footnote" | "endnote", string>> = {};
    const outlines = new Map<string, Outline>();
    for (const rel of documentRels(zip, docPart)) {
      const kind = rel.type.endsWith("/footnotes") ? "footnote" : rel.type.endsWith("/endnotes") ? "endnote" : null;
      if (!kind || !zip.entries.has(rel.part) || outlines.has(rel.part)) continue;
      noteParts[kind] = rel.part;
      outlines.set(rel.part, await outline(zip, rel.part, join(dir, `${side}-${fileName(rel.part)}`)));
    }

    // The shell: document.xml with only the body's own section properties.
    // Its parts (everything but the body, the notes on disk,
    // and the binary stubs) are held whole, so their sizes are checked
    // first, from the directory, and each is inflated once, straight into
    // the document.
    const held = (entry: ZipEntry): boolean => entry.name !== docPart && !entry.name.endsWith("/") && !outlines.has(entry.name) && /\.(xml|rels|vml)$/i.test(entry.name);
    let heldBytes = 0;
    for (const entry of zip.entries.values()) if (held(entry)) heldBytes += entry.size;
    if (heldBytes > MAX_HELD_BYTES) {
      throw new CompareError("too_large", `${zip.path}: the parts held in memory (all but the body, the notes and the binary parts) come to ${mb(heldBytes)}; at most ${mb(MAX_HELD_BYTES)} is supported, so that a compare fits in a ${HEAP_MB} MB heap`);
    }
    const files: Record<string, Uint8Array> = {};
    for (const entry of zip.entries.values()) {
      if (entry.name === docPart || entry.name.endsWith("/")) continue;
      const empty = outlines.get(entry.name);
      if (empty) files[entry.name] = Buffer.from(empty.head + empty.tail, "utf8");
      else files[entry.name] = held(entry) ? zip.read(entry) : Buffer.from(`${STUB}${side}\u0000${entry.name}`, "utf8");
    }
    const skeleton: Buffer[] = [Buffer.from(range(0, bounds.contentStart))];
    for (let i = 0; i < names.length; i++) if (isSectPr(i)) skeleton.push(Buffer.from(range(starts[i], ends[i])));
    skeleton.push(Buffer.from(range(bounds.contentEnd, fstatSync(fd).size)));
    files[docPart] = Buffer.concat(skeleton);
    /** Refuse a w:altChunk or w:subDoc in `el`, from `part`: its content is
     * in another part or file, which Word reads and the compare does not. */
    const refuseUnsupported = (el: XmlElement, part: string): void => {
      const ln = localName(el.name);
      if (ln === "altChunk" || ln === "subDoc") {
        const rels = files[relsPathFor(part)];
        const partRels = parseRelationships(rels ? parseXml(Buffer.from(rels).toString("utf8")) : undefined, part);
        throw ln === "altChunk" ? altChunkError(el, partRels) : subDocError(el, partRels);
      }
      for (const c of el.children) refuseUnsupported(c, part);
    };
    const unsupported = (xml: string): boolean => xml.includes("altChunk") || xml.includes("subDoc");
    for (const rel of documentRels(zip, docPart)) {
      const text = /\/(header|footer)$/.test(rel.type) && files[rel.part] ? Buffer.from(files[rel.part]).toString("utf8") : "";
      if (unsupported(text)) refuseUnsupported(parseXml(text), rel.part);
    }
    // The compare reads the headers' and footers' XML; their block models would cost about 1 KB a paragraph.
    const shell = DocxDocument.fromFiles(files, { headerFooterBlocks: false });
    let hadRevisions = false;
    if (collectRevisions(shell).length > 0) {
      acceptAllRevisions(shell);
      hadRevisions = true;
    }

    const input: Input = {
      side,
      zip,
      docPart,
      shell,
      blocks: new BlockStore(join(dir, `${side}-blocks`)),
      texts: new BlockStore(join(dir, `${side}-texts`)),
      fingerprints: [],
      hashes: [],
      kinds: [],
      endsSection: [],
      sectionKeys: new Map(),
      sectPrs: [],
      rows: new BlockStore(join(dir, `${side}-rows`)),
      rowFingerprints: [],
      rowKinds: [],
      tables: new Map(),
      noteRefs: { footnote: [], endnote: [] },
      commentRefs: [],
      notes: { footnote: null, endnote: null },
      hadRevisions,
    };

    const noteRefsIn = (e: XmlElement): void => {
      const ln = localName(e.name);
      if (ln === "footnoteReference") input.noteRefs.footnote.push(e.attrs["w:id"]);
      else if (ln === "endnoteReference") input.noteRefs.endnote.push(e.attrs["w:id"]);
      else if (ln === "commentReference") input.commentRefs.push(e.attrs["w:id"]);
      for (const c of e.children) noteRefsIn(c);
    };

    /** cyrb53 as a number: eight bytes in an array of them, where its string is forty. */
    const hash = (s: string): number => parseInt(cyrb53(s), 36);

    /** One block's index entry: its XML, the hash of the XML it stands for, its text and its fingerprint. */
    const index = (el: XmlElement, xml: string, xmlHash: number, text: string, fingerprint: number): void => {
      const i = input.blocks.count;
      input.blocks.add(xml);
      input.texts.add(text);
      input.fingerprints.push(fingerprint);
      input.hashes.push(xmlHash);
      input.kinds.push(localName(el.name));
      const sectPr = localName(el.name) === "p" ? child(child(el, "pPr"), "sectPr") : undefined;
      input.endsSection.push(!!sectPr);
      if (sectPr) {
        input.sectionKeys.set(i, sectionReferences(el, shell));
        input.sectPrs.push(sectPr);
      }
      noteRefsIn(el);
    };

    /**
     * Index one block. `xml` is the block as read when nothing changed it: a
     * block with no tracked changes is stored as its own bytes, not parsed
     * and written again (that was a fifth of a large compare's time).
     */
    const add = (el: XmlElement, xml = serializeXml(el)): void => {
      index(el, xml, hash(xml), paragraphText(el), hash(blockFingerprint(el, shell)));
    };

    /**
     * A table, one child at a time: its rows go to `rows`, and its block is
     * the shell. The table's fingerprint and hash are those of the whole
     * element, taken as the children pass: a table's text runs to megabytes,
     * and is never held whole.
     */
    const beginTable = (open: XmlElement) => {
      const from = input.rows.count;
      const shell: XmlElement = { ...open, children: [] };
      const text = new TextHasher();
      const rowXml = createHash("sha1");
      return {
        open,
        /** How many w:tr rows it holds. */
        trs: 0,
        /** Take the table's next child; `raw` is its XML as read, when nothing changed it. */
        take(el: XmlElement, raw?: string): void {
          text.add(paragraphText(el));
          const rowKind = ROW_LEVEL.indexOf(localName(el.name));
          if (rowKind === -1) {
            shell.children.push(el);
            return;
          }
          if (ROW_LEVEL[rowKind] === "tr") this.trs++;
          const xml = raw ?? serializeXml(el);
          input.rows.add(xml);
          input.rowFingerprints.push(hash(rowFingerprint(el)));
          input.rowKinds.push(ROW_LEVEL[rowKind]);
          rowXml.update(xml);
          shell.children.push(ROW_SLOT);
          noteRefsIn(el);
        },
        end(): void {
          input.tables.set(input.blocks.count, [from, input.rows.count]);
          const xml = serializeXml(shell);
          index(shell, xml, digest(rowXml.update(xml)), "", text.digest());
        },
      };
    };

    // Blocks with tracked changes go through the shell's body, a batch at a
    // time, to be accepted as acceptAllRevisions accepts a whole document.
    // `wrap` is the body child the batch sat in: a table around its rows,
    // or null for body blocks.
    const body = child(shell.docRoot, "body");
    if (!body) throw new CompareError("not_docx", `${zip.path}: document.xml has no w:body`);
    const tail = body.children;
    const accept = (els: XmlElement[], wrap: XmlElement | null): XmlElement[] => {
      body.children = [...(wrap ? [{ ...wrap, children: els }] : els), ...tail];
      acceptAllRevisions(shell);
      const kept = body.children.filter((c) => !tail.includes(c));
      body.children = tail;
      return wrap ? kept.flatMap((t) => t.children) : kept;
    };

    /**
     * Read a table's children. Rows with tracked changes are accepted a batch
     * at a time, any other child on its own: a batch whose rows are all
     * deleted takes its stand-in table with it, and must take nothing else.
     * Null when accepting removed the table, as it does when the last w:tr
     * row is deleted.
     */
    const readTable = (start: number, name: string) => {
      let table!: ReturnType<typeof beginTable>;
      let rows: XmlElement[] = [];
      let rowBytes = 0;
      let hadTr = false;
      const flushRows = (): void => {
        if (rows.length > 0) for (const el of accept(rows, table.open)) table.take(el);
        rows = [];
        rowBytes = 0;
      };
      // Each child is taken as the scan finds it: a table's children are too many to list first.
      scanChildren(
        fd,
        start,
        (contentStart) => {
          const openTag = range(start, contentStart).toString("utf8");
          table = beginTable(parseXml(openTag.endsWith("/>") ? openTag : `${openTag}</${name}>`));
        },
        (c) => {
          const xml = range(c.start, c.end).toString("utf8");
          const el = parseXml(xml);
          if (unsupported(xml)) refuseUnsupported(el, docPart);
          const tracked = TRACKED_CHANGE.test(xml);
          if (tracked) hadRevisions = true;
          if (!ROW_LEVEL.includes(localName(el.name))) {
            flushRows();
            if (tracked) for (const e of accept([el], table.open)) table.take(e);
            else table.take(el);
            return;
          }
          hadTr ||= localName(el.name) === "tr";
          if (rows.length > 0 || tracked) {
            rows.push(el);
            rowBytes += xml.length;
            if (rowBytes >= BATCH_BYTES) flushRows();
          } else table.take(el, xml);
        },
      );
      flushRows();
      return hadTr && table.trs === 0 ? null : table;
    };

    // A deleted paragraph mark joins its paragraph to the next block, so a
    // batch never ends on one. A table is accepted apart: a batch open when
    // one comes is flushed once the table stands, as nothing joins across a
    // table, and goes on when accepting removed the table, as the join then
    // crosses where it stood.
    let batch: XmlElement[] = [];
    let batchBytes = 0;
    const flush = (): void => {
      if (batch.length > 0) for (const el of accept(batch, null)) add(el);
      batch = [];
      batchBytes = 0;
    };
    for (let i = 0; i < names.length; i++) {
      if (isSectPr(i)) continue;
      if (localName(names[i]) === "tbl") {
        const table = readTable(starts[i], names[i]);
        if (table) {
          flush();
          table.end();
        }
        continue;
      }
      const xml = range(starts[i], ends[i]).toString("utf8");
      const el = parseXml(xml);
      if (unsupported(xml)) refuseUnsupported(el, docPart);
      if (batch.length > 0 || TRACKED_CHANGE.test(xml)) {
        hadRevisions = true;
        batch.push(el);
        batchBytes += xml.length;
        if (batchBytes >= BATCH_BYTES && !joinsNext(el)) flush();
      } else {
        add(el, xml);
      }
    }
    flush();

    // The notes, a part at a time, accepted as body blocks are.
    for (const kind of ["footnote", "endnote"] as const) {
      const part = noteParts[kind];
      if (!part) continue;
      const { file, head, tail } = outlines.get(part) as Outline;
      const notes = new DiskNotes(kind, part, join(dir, `${side}-${fileName(part)}-notes`));
      notes.head = head;
      notes.tail = tail;
      input.notes[kind] = notes;
      const notesFd = openSync(file, "r");
      try {
        let batch: XmlElement[] = [];
        let batchBytes = 0;
        const flushNotes = (): void => {
          if (batch.length > 0) for (const el of accept(batch, null)) notes.add(el);
          batch = [];
          batchBytes = 0;
        };
        scanChildren(
          notesFd,
          0,
          () => {},
          (c) => {
            checkBlock(part, c.end - c.start);
            const xml = readRange(notesFd, c.start, c.end);
            const el = parseXml(xml);
            if (unsupported(xml)) refuseUnsupported(el, part);
            if (batch.length > 0 || TRACKED_CHANGE.test(xml)) {
              hadRevisions = true;
              batch.push(el);
              batchBytes += xml.length;
              if (batchBytes >= BATCH_BYTES) flushNotes();
            } else {
              notes.add(el, xml);
            }
          },
        );
        flushNotes();
      } finally {
        closeSync(notesFd);
      }
    }
    input.hadRevisions = hadRevisions;
    return input;
  } finally {
    closeSync(fd);
  }
}

/** A part extracted to disk, and its bytes around its root's children: the part with none. */
interface Outline {
  file: string;
  head: string;
  tail: string;
}

/** Extract `part` to `file`, and find its root's children. */
async function outline(zip: ZipReader, part: string, file: string): Promise<Outline> {
  await zip.extract(zip.entries.get(part) as ZipEntry, file);
  const fd = openSync(file, "r");
  try {
    let bounds;
    try {
      bounds = scanChildren(fd, 0, () => {}, () => {});
    } catch (error) {
      throw withPart(part, error);
    }
    return { file, head: readRange(fd, 0, bounds.contentStart), tail: readRange(fd, bounds.contentEnd, fstatSync(fd).size) };
  } finally {
    closeSync(fd);
  }
}

/** `error` with the part it was found in named first, its code kept. */
function withPart(part: string, error: unknown): CompareError {
  const message = `${part}: ${error instanceof Error ? error.message : String(error)}`;
  return new CompareError(error instanceof CompareError ? error.code : "damaged", message, undefined, { cause: error });
}

/** A part name as a file name. */
function fileName(part: string): string {
  return part.replace(/\W/g, "-");
}

/** Refuse a block (a paragraph, a table's row, a note) of more bytes than
 * parseXml takes, from its size in the file, before it is read. */
function checkBlock(part: string, bytes: number): void {
  if (bytes > DEFAULT_XML_LIMITS.maxLength) throw new CompareError("too_large", `${part}: a block is ${mb(bytes)}; one block is at most ${mb(DEFAULT_XML_LIMITS.maxLength)}`);
}

function readRange(fd: number, start: number, end: number): string {
  const buf = Buffer.allocUnsafe(end - start);
  let done = 0;
  while (done < buf.length) done += readSync(fd, buf, done, buf.length - done, start + done);
  return buf.toString("utf8");
}

/** The main document part's relationships: each one's type and the part it names. */
function documentRels(zip: ZipReader, docPart: string): { type: string; part: string }[] {
  const at = docPart.lastIndexOf("/");
  const entry = zip.entries.get(`${docPart.slice(0, at + 1)}_rels/${docPart.slice(at + 1)}.rels`);
  if (!entry) return [];
  return parseXml(zip.read(entry).toString("utf8"))
    .children.filter((rel) => rel.attrs.TargetMode !== "External" && rel.attrs.Target)
    .map((rel) => ({ type: rel.attrs.Type ?? "", part: resolvePartPath(docPart, rel.attrs.Target) }));
}

/**
 * tableFingerprint of a text fed a piece at a time, hashed as it goes: the
 * whole text's normalizeText, with a whitespace run that crosses a piece
 * boundary collapsed as one.
 */
class TextHasher {
  private readonly hash = createHash("sha1");
  /** Whether any text is in: whitespace before the first is trimmed. */
  private started = false;
  /** Whether whitespace ended the text so far: one space, once more text comes. */
  private space = false;

  add(piece: string): void {
    if (/^\s/.test(piece)) this.space = true;
    const text = normalizeText(piece);
    if (text) {
      if (this.started && this.space) this.hash.update(" ");
      this.hash.update(text);
      this.started = true;
      this.space = false;
    }
    if (/\s$/.test(piece)) this.space = true;
  }

  digest(): number {
    return digest(this.hash);
  }
}

/** A streamed hash as a number, like `hash`. */
function digest(hash: Hash): number {
  return hash.digest().readUIntBE(0, 6);
}

/** A paragraph whose mark is deleted or moved away: accepting it joins the next block. */
function joinsNext(el: XmlElement): boolean {
  const rPr = child(child(el, "pPr"), "rPr");
  return localName(el.name) === "p" && !!(child(rPr, "del") || child(rPr, "moveFrom"));
}

/** The main document part, as the package relationships name it. */
function documentPart(zip: ZipReader): string {
  const rels = zip.entries.get("_rels/.rels");
  if (rels) {
    const root = parseXml(zip.read(rels).toString("utf8"));
    for (const rel of root.children) {
      if ((rel.attrs.Type ?? "").endsWith("/officeDocument")) return (rel.attrs.Target ?? "").replace(/^\//, "");
    }
  }
  return "word/document.xml";
}
