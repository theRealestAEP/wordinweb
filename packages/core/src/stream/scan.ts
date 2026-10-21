import { DEFAULT_XML_LIMITS, XmlParseError } from "../xml.js";
import { DOCUMENT_LIMITS } from "../zip.js";

/** One block-level child of the body: its qualified name, and its bytes,
 * valid only until the scanner reads on. */
export interface ScannedChild {
  name: string;
  /** Whether it is the body's own child, not one of a w:sdt's. */
  top: boolean;
  /** The innermost block-level w:sdt it sits in (see ScannedSdt), or -1. */
  sdt: number;
  bytes: Uint8Array;
}

/** A block-level w:sdt, reported once its w:sdtContent opens: its bytes from
 * its start tag through the sdtContent's, which hold its properties. */
export interface ScannedSdt {
  id: number;
  /** The block-level w:sdt this one sits in, or -1. */
  parent: number;
  bytes: Uint8Array;
}

/** The body's bytes around its children: the part with none of them. */
export interface BodyOutline {
  head: Uint8Array;
  tail: Uint8Array;
}

const CHUNK = 1 << 20;
const FIRST_READ = 64 * 1024;
const LT = 0x3c;
const GT = 0x3e;
const SLASH = 0x2f;

/**
 * Find each block-level child of a document part's w:body, in order, reading
 * the part a chunk at a time: the body's own children and, through any w:sdt
 * among them, the children of its w:sdtContent (nested as deep as they go, as
 * parse/document.ts walks them). A w:customXml is walked into too, and so is
 * an mc:AlternateContent, through its first branch only (parse/document.ts's
 * wrappedContent). The browser twin of compare-cli's scanner,
 * over a pull of inflated chunks instead of a file. Comments, processing
 * instructions and CDATA are skipped; a document type declaration, an
 * unterminated construct or a missing body is an error. `onAltChunk` gets
 * the start tag of every w:altChunk in the body, at any depth.
 *
 * The body is held to DOCUMENT_LIMITS as it streams: its elements, and its
 * paragraphs and table cells, are counted at each start tag, and a block
 * (or any construct the scanner must hold whole) over blockBytes bytes,
 * blockText bytes of text, blockElements elements or tableRows rows is
 * refused before more of it is read.
 */
export function scanBody(
  pull: () => Uint8Array | null,
  onChild: (child: ScannedChild) => void,
  onSdt: (sdt: ScannedSdt) => void,
  onAltChunk: (tag: string) => void,
): BodyOutline {
  return new Scanner(pull).run(onChild, onSdt, onAltChunk);
}

class Scanner {
  private store = new Uint8Array(2 * FIRST_READ);
  /** The part's offset of store[0]. */
  private base = 0;
  /** The read part of `store`. */
  private length = 0;
  private eof = false;
  /** The part's offset of the first byte that must stay held: the open block's start, or where the scan is. */
  private keepFrom = 0;

  constructor(private readonly pull: () => Uint8Array | null) {}

  run(onChild: (child: ScannedChild) => void, onSdt: (sdt: ScannedSdt) => void, onAltChunk: (tag: string) => void): BodyOutline {
    // The body sits at depth 1, under the document root.
    const cd = 1;
    let depth = 0;
    let inside = false;
    let head: Uint8Array | null = null;
    // Depths whose children are blocks: the body's, then each open sdtContent's.
    const containers: number[] = [];
    // Each open mc:AlternateContent container, by depth: whether a branch
    // has been walked into. The branches after it are children, which the
    // caller passes over.
    const alternates = new Map<number, boolean>();
    // Each open block-level w:sdt: its level, where it starts, and its id
    // once its sdtContent has opened.
    const sdts: { level: number; at: number; id: number }[] = [];
    let nextSdt = 0;
    let childStart = -1;
    let childName = "";
    let pos = 0;
    // The body's elements, its paragraphs and table cells, and the elements before the open block's.
    let elements = 0;
    let blocks = 0;
    let childFirst = 0;
    // The open block's text (its bytes between tags), and its table rows.
    let childText = 0;
    let childRows = 0;
    for (;;) {
      // An open child's bytes stay, as its reader wants them, and so do an
      // open sdt's properties.
      this.keepFrom = childStart >= 0 ? Math.min(this.base + pos, childStart) : this.base + pos;
      const sdt = sdts[sdts.length - 1];
      if (sdt && sdt.id === -1) this.keepFrom = Math.min(this.keepFrom, sdt.at);
      if (pos > CHUNK) {
        // Move what is left to the front, so the buffer never grows past a
        // chunk or two, besides the bytes that stay.
        const keep = this.keepFrom - this.base;
        this.store.copyWithin(0, keep, this.length);
        this.length -= keep;
        this.base += keep;
        pos -= keep;
      }
      let lt = this.indexOf(LT, pos);
      while (lt === -1) {
        if (!this.more()) throw new XmlParseError(inside ? "an element is not closed" : "no w:body: it is not a Word document");
        lt = this.indexOf(LT, pos);
      }
      const at = this.base + lt;
      if (childStart >= 0 && (childText += lt - pos) > DOCUMENT_LIMITS.blockText) throw new Error(`too large: the body holds a block of over ${DOCUMENT_LIMITS.blockText} bytes of text`);
      const next = this.byteAt(lt + 1);
      let end: number;
      if (next === 0x3f /* ? */) {
        end = this.find("?>", lt + 2, "processing instruction");
      } else if (next === 0x21 /* ! */) {
        if (this.startsWith("<!--", lt)) end = this.find("-->", lt + 4, "comment");
        else if (this.startsWith("<![CDATA[", lt)) end = this.find("]]>", lt + 9, "CDATA section");
        else throw new XmlParseError("a document type declaration is not allowed");
      } else if (next === SLASH) {
        end = this.find(">", lt + 2, "end tag");
        depth--;
        if (inside) {
          const top = containers[containers.length - 1];
          if (depth === cd) {
            const tail = this.rest(at);
            return { head: head as Uint8Array, tail };
          }
          // A block's end tag and an sdt's both come at the container's depth
          // plus one: only a block is open as a child.
          if (depth === top + 1 && childStart >= 0) {
            if (end - childStart > DOCUMENT_LIMITS.blockBytes) throw new Error(`too large: the body holds a block over ${DOCUMENT_LIMITS.blockBytes} bytes`);
            onChild({ name: childName, top: containers.length === 1, sdt: sdts[sdts.length - 1]?.id ?? -1, bytes: this.store.subarray(childStart - this.base, end - this.base) });
            childStart = -1;
          } else if (sdts.length > 0 && depth === sdts[sdts.length - 1].level - 1) sdts.pop();
          else if (depth === top && containers.length > 1) {
            containers.pop();
            alternates.delete(top);
          }
        }
      } else {
        end = this.tagEnd(lt + 1);
        const selfClosing = this.store[end - this.base - 2] === SLASH;
        const name = this.nameAt(lt + 1);
        if (inside) {
          const ln = localName(name);
          if (ln === "altChunk") onAltChunk(new TextDecoder().decode(this.store.subarray(lt, end - this.base)));
          if (++elements > DOCUMENT_LIMITS.documentElements) throw new Error(`too large: the body holds over ${DOCUMENT_LIMITS.documentElements} elements`);
          if ((ln === "p" || ln === "tc") && ++blocks > DOCUMENT_LIMITS.documentBlocks) throw new Error(`too large: the body holds over ${DOCUMENT_LIMITS.documentBlocks} paragraphs and table cells`);
          if (childStart >= 0 && elements - childFirst > DOCUMENT_LIMITS.blockElements) throw new Error(`too large: the body holds a block of over ${DOCUMENT_LIMITS.blockElements} elements`);
          if (childStart >= 0 && ln === "tr" && ++childRows > DOCUMENT_LIMITS.tableRows) throw new Error(`too large: the body holds a table of over ${DOCUMENT_LIMITS.tableRows} rows`);
        }
        if (!inside) {
          if (depth === cd && localName(name) === "body") {
            if (this.base !== 0) throw new XmlParseError("the document's head is longer than a chunk");
            head = this.store.slice(0, end);
            inside = true;
            containers.push(cd);
            if (selfClosing) return { head, tail: this.rest(end) };
          }
        } else {
          const top = containers[containers.length - 1];
          if (depth === top + 1) {
            const ln = localName(name);
            if (ln === "body") throw new XmlParseError("a w:body inside the body");
            if (ln === "sdt" && !selfClosing) sdts.push({ level: depth + 1, at, id: -1 });
            else if ((ln === "customXml" || ln === "AlternateContent") && !selfClosing) {
              containers.push(depth);
              if (ln === "AlternateContent") alternates.set(depth, false);
            } else if ((ln === "Choice" || ln === "Fallback") && !selfClosing && alternates.get(top) === false) {
              alternates.set(top, true);
              containers.push(depth);
            } else if (selfClosing) {
              if (end - at > DOCUMENT_LIMITS.blockBytes) throw new Error(`too large: the body holds a block over ${DOCUMENT_LIMITS.blockBytes} bytes`);
              onChild({ name, top: containers.length === 1, sdt: sdts[sdts.length - 1]?.id ?? -1, bytes: this.store.subarray(lt, end - this.base) });
            } else {
              childStart = at;
              childName = name;
              childFirst = elements - 1;
              childText = 0;
              childRows = 0;
            }
          } else if (sdts.length > 0 && depth === sdts[sdts.length - 1].level && localName(name) === "sdtContent" && !selfClosing) {
            const sdt = sdts[sdts.length - 1];
            sdt.id = nextSdt++;
            onSdt({ id: sdt.id, parent: sdts[sdts.length - 2]?.id ?? -1, bytes: this.store.subarray(sdt.at - this.base, end - this.base) });
            containers.push(depth);
          }
        }
        if (!selfClosing && ++depth >= DEFAULT_XML_LIMITS.maxDepth) throw new XmlParseError(`XML parse error: nesting too deep (> ${DEFAULT_XML_LIMITS.maxDepth})`);
      }
      pos = end - this.base;
    }
  }

  /** Everything from part offset `from` to the end of the part. */
  private rest(from: number): Uint8Array {
    this.keepFrom = from;
    while (this.more());
    return this.store.slice(from - this.base, this.length);
  }

  /** Read the next chunk; false at the end of the part. What is held is
   * about one block, or one construct, and a chunk: once the bytes that
   * must stay are over DOCUMENT_LIMITS.blockBytes, the block is refused
   * before more of it is read. */
  private more(): boolean {
    if (this.eof) return false;
    const limit = DOCUMENT_LIMITS.blockBytes;
    if (this.base + this.length - this.keepFrom > limit) throw new Error(`too large: the body holds a block over ${limit} bytes`);
    const chunk = this.pull();
    if (chunk === null) {
      this.eof = true;
      return false;
    }
    if (this.length + chunk.length > this.store.length) {
      const bigger = new Uint8Array(2 * (this.length + chunk.length));
      bigger.set(this.store.subarray(0, this.length));
      this.store = bigger;
    }
    this.store.set(chunk, this.length);
    this.length += chunk.length;
    return true;
  }

  private indexOf(byte: number, from: number): number {
    const i = this.store.subarray(0, this.length).indexOf(byte, from);
    return i;
  }

  private byteAt(i: number): number {
    while (i >= this.length) if (!this.more()) throw new XmlParseError("unexpected end of file");
    return this.store[i];
  }

  private startsWith(text: string, i: number): boolean {
    for (let k = 0; k < text.length; k++) if (this.byteAt(i + k) !== text.charCodeAt(k)) return false;
    return true;
  }

  /** The part offset just past `text`, searched from buffer index `from`. */
  private find(text: string, from: number, what: string): number {
    const first = text.charCodeAt(0);
    for (;;) {
      let i = this.indexOf(first, from);
      while (i !== -1 && i + text.length <= this.length) {
        let k = 1;
        while (k < text.length && this.store[i + k] === text.charCodeAt(k)) k++;
        if (k === text.length) return this.base + i + text.length;
        i = this.indexOf(first, i + 1);
      }
      // Keep looking from where a match could still start.
      from = Math.max(from, this.length - text.length + 1);
      if (!this.more()) throw new XmlParseError(`unterminated ${what}`);
    }
  }

  /** The part offset just past the `>` that ends the start tag at buffer index `from`. */
  private tagEnd(from: number): number {
    let quote = 0;
    for (let i = from; ; i++) {
      if (i >= this.length && !this.more()) throw new XmlParseError("unterminated start tag");
      const c = this.store[i];
      if (quote) {
        if (c === quote) quote = 0;
      } else if (c === 0x22 || c === 0x27) quote = c;
      else if (c === GT) return this.base + i + 1;
      else if (c === LT) throw new XmlParseError("'<' inside a start tag");
    }
  }

  private nameAt(from: number): string {
    let i = from;
    for (;;) {
      const c = this.byteAt(i);
      if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === SLASH || c === GT) break;
      i++;
    }
    return String.fromCharCode(...this.store.subarray(from, i));
  }
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}
