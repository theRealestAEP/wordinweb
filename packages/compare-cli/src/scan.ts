import { closeSync, openSync, readSync } from "node:fs";

/** One child of w:body: its bytes are [start, end) of the file. */
export interface BodyChild {
  start: number;
  end: number;
  /** Its qualified name, such as w:p. */
  name: string;
}

/** Where the body's content sits in the file. */
export interface BodyBounds {
  /** Just after the w:body start tag. */
  contentStart: number;
  /** At the w:body end tag. */
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
export function scanBody(path: string, onChild: (child: BodyChild) => void): BodyBounds {
  const fd = openSync(path, "r");
  try {
    return new Scanner(fd).run(onChild);
  } finally {
    closeSync(fd);
  }
}

class Scanner {
  /** Holds the file from byte `base` on, as far as it has been read. */
  private store = Buffer.allocUnsafe(2 * CHUNK);
  /** The read part of `store`. */
  private buf = this.store.subarray(0, 0);
  private base = 0;
  private eof = false;

  constructor(private readonly fd: number) {}

  run(onChild: (child: BodyChild) => void): BodyBounds {
    // Open elements: the root is depth 1, the body depth 2, its children 3.
    let depth = 0;
    let inBody = false;
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
        if (!this.more()) throw new Error(inBody ? "the body is not closed" : "no w:body: it is not a Word document");
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
        else throw new Error("a document type declaration is not allowed");
      } else if (next === SLASH) {
        end = this.find(">", lt + 2, "end tag");
        depth--;
        if (inBody && depth === 1) return { contentStart, contentEnd: at };
        if (inBody && depth === 2) onChild({ start: childStart, end: end, name: childName });
      } else {
        end = this.tagEnd(lt + 1);
        const selfClosing = this.buf[end - this.base - 2] === SLASH;
        const name = this.nameAt(lt + 1);
        if (!inBody && depth === 1 && localName(name) === "body") {
          if (selfClosing) return { contentStart: end, contentEnd: end };
          inBody = true;
          contentStart = end;
        } else if (inBody && depth === 2) {
          if (localName(name) === "body") throw new Error("a w:body inside the body");
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
    while (i >= this.buf.length) if (!this.more()) throw new Error("unexpected end of file");
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
      if (!this.more()) throw new Error(`unterminated ${what}`);
    }
  }

  /** The file offset just past the `>` that ends the start tag at buffer index `from`. */
  private tagEnd(from: number): number {
    let quote = 0;
    for (let i = from; ; i++) {
      if (i >= this.buf.length && !this.more()) throw new Error("unterminated start tag");
      const c = this.buf[i];
      if (quote) {
        if (c === quote) quote = 0;
      } else if (c === 0x22 || c === 0x27) quote = c;
      else if (c === GT) return this.base + i + 1;
      else if (c === LT) throw new Error("'<' inside a start tag");
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
