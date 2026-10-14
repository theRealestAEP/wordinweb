import { closeSync, openSync, readSync } from "node:fs";

/** One child of an element: its bytes are [start, end) of the file. */
export interface Child {
  start: number;
  end: number;
  /** Its qualified name, such as w:p. */
  name: string;
}

/** Where an element's content sits in the file. */
export interface Bounds {
  /** Just after the start tag. */
  contentStart: number;
  /** At the end tag. */
  contentEnd: number;
}

const CHUNK = 1 << 20;
const LT = 0x3c;
const GT = 0x3e;
const SLASH = 0x2f;

/**
 * Find each child element of the document's w:body, in order, reading the
 * file a chunk at a time. Comments, processing instructions and CDATA are
 * skipped; a document type declaration, an unterminated construct or a
 * missing body is an error.
 */
export function scanBody(path: string, onChild: (child: Child) => void): Bounds {
  const fd = openSync(path, "r");
  try {
    return new Scanner(fd, 0).run({ depth: 1, name: "body" }, onChild);
  } finally {
    closeSync(fd);
  }
}

/** Find each child element of the element whose start tag is at byte `start`, as scanBody finds the body's. */
export function scanChildren(fd: number, start: number, onChild: (child: Child) => void): Bounds {
  return new Scanner(fd, start).run({ depth: 0 }, onChild);
}

/** The element whose children are wanted: its depth below the scan's start, and its local name if it must be found by name. */
interface Container {
  depth: number;
  name?: string;
}

class Scanner {
  /** Holds the file from byte `base` on, as far as it has been read. */
  private store = Buffer.allocUnsafe(2 * CHUNK);
  /** The read part of `store`. */
  private buf = this.store.subarray(0, 0);
  private eof = false;

  constructor(
    private readonly fd: number,
    private base: number,
  ) {}

  run(container: Container, onChild: (child: Child) => void): Bounds {
    // Open elements: the container's children are one deeper than it.
    const cd = container.depth;
    let depth = 0;
    let inside = false;
    let contentStart = -1;
    let childStart = -1;
    let childName = "";
    let pos = 0;
    for (;;) {
      if (pos > CHUNK) {
        // Move what is left to the front, so the buffer never grows past a chunk or two.
        this.store.copyWithin(0, pos, this.buf.length);
        this.buf = this.store.subarray(0, this.buf.length - pos);
        this.base += pos;
        pos = 0;
      }
      let lt = this.buf.indexOf(LT, pos);
      while (lt === -1) {
        if (!this.more()) throw new Error(inside ? "document.xml: an element is not closed" : "document.xml: no w:body");
        lt = this.buf.indexOf(LT, pos);
      }
      const at = this.base + lt;
      const next = this.byteAt(lt + 1);
      let end: number;
      if (next === 0x3f /* ? */) {
        end = this.find("?>", lt + 2, "processing instruction");
      } else if (next === 0x21 /* ! */) {
        if (this.startsWith("<!--", lt)) end = this.find("-->", lt + 4, "comment");
        else if (this.startsWith("<![CDATA[", lt)) end = this.find("]]>", lt + 9, "CDATA section");
        else throw new Error("document.xml: a document type declaration is not allowed");
      } else if (next === SLASH) {
        end = this.find(">", lt + 2, "end tag");
        depth--;
        if (inside && depth === cd) return { contentStart, contentEnd: at };
        if (inside && depth === cd + 1) onChild({ start: childStart, end: end, name: childName });
      } else {
        end = this.tagEnd(lt + 1);
        const selfClosing = this.buf[end - this.base - 2] === SLASH;
        const name = this.nameAt(lt + 1);
        if (!inside && depth === cd && (!container.name || localName(name) === container.name)) {
          if (selfClosing) return { contentStart: end, contentEnd: end };
          inside = true;
          contentStart = end;
        } else if (inside && depth === cd + 1) {
          if (container.name && localName(name) === container.name) throw new Error(`document.xml: a w:${container.name} inside the ${container.name}`);
          if (selfClosing) {
            onChild({ start: at, end: end, name });
          } else {
            childStart = at;
            childName = name;
          }
        }
        if (!selfClosing) depth++;
      }
      pos = end - this.base;
    }
  }

  /** Read the next chunk; false at the end of the file. */
  private more(): boolean {
    if (this.eof) return false;
    const length = this.buf.length;
    if (length + CHUNK > this.store.length) {
      // One token longer than the buffer: grow it.
      const bigger = Buffer.allocUnsafe(2 * (length + CHUNK));
      this.store.copy(bigger, 0, 0, length);
      this.store = bigger;
    }
    const n = readSync(this.fd, this.store, length, CHUNK, this.base + length);
    if (n === 0) {
      this.eof = true;
      return false;
    }
    this.buf = this.store.subarray(0, length + n);
    return true;
  }

  private byteAt(i: number): number {
    while (i >= this.buf.length) if (!this.more()) throw new Error("document.xml: unexpected end of file");
    return this.buf[i];
  }

  private startsWith(text: string, i: number): boolean {
    for (let k = 0; k < text.length; k++) if (this.byteAt(i + k) !== text.charCodeAt(k)) return false;
    return true;
  }

  /** The file offset just past `text`, searched from buffer index `from`. */
  private find(text: string, from: number, what: string): number {
    for (;;) {
      const i = this.buf.indexOf(text, from, "latin1");
      if (i !== -1) return this.base + i + text.length;
      // Keep looking from where a match could still start.
      from = Math.max(from, this.buf.length - text.length + 1);
      if (!this.more()) throw new Error(`document.xml: unterminated ${what}`);
    }
  }

  /** The file offset just past the `>` that ends the start tag at buffer index `from`. */
  private tagEnd(from: number): number {
    let quote = 0;
    for (let i = from; ; i++) {
      if (i >= this.buf.length && !this.more()) throw new Error("document.xml: unterminated start tag");
      const c = this.buf[i];
      if (quote) {
        if (c === quote) quote = 0;
      } else if (c === 0x22 || c === 0x27) quote = c;
      else if (c === GT) return this.base + i + 1;
      else if (c === LT) throw new Error("document.xml: '<' inside a start tag");
    }
  }

  private nameAt(from: number): string {
    let i = from;
    for (;;) {
      const c = this.byteAt(i);
      if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d || c === SLASH || c === GT) break;
      i++;
    }
    return this.buf.toString("latin1", from, i);
  }
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}
