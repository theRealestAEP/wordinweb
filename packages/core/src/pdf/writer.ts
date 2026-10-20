import { Writable } from "node:stream";
import { zlibSync } from "fflate";

/**
 * A PDF written front to back to a Writable, one object at a time: the
 * header first, each object as it is complete, the cross-reference table
 * and trailer last. Objects are numbered when allocated, so a page can
 * point at a font or the page tree before either is written. Writes wait
 * for the stream to drain, so a slow consumer holds the producer back
 * instead of buffering the document.
 */
export class PdfWriter {
  private position = 0;
  /** Byte offset of each object, by number - 1; -1 until it is written. */
  private readonly offsets: number[] = [];
  private failed: Error | null = null;
  private readonly onError = (error: Error): void => {
    this.failed = error;
  };

  constructor(private readonly out: Writable) {
    out.on("error", this.onError);
  }

  /** Reserve an object number. */
  allocate(): number {
    this.offsets.push(-1);
    return this.offsets.length;
  }

  async header(): Promise<void> {
    // The comment line of high bytes marks the file binary for transfer tools.
    await this.write("%PDF-1.5\n%\xe2\xe3\xcf\xd3\n");
  }

  /** Write a non-stream object. */
  async object(num: number, body: string): Promise<void> {
    await this.beginObject(num);
    await this.write(`${body}\nendobj\n`);
  }

  /** Write a stream object with the data deflated (zlib format, as
   * FlateDecode reads it), `dict` the entries besides Length and Filter. */
  async deflatedStream(num: number, dict: string, data: Uint8Array): Promise<void> {
    await this.stream(num, `${dict} /Filter /FlateDecode`, zlibSync(data));
  }

  /** Write a stream object with `data` as it is. */
  async stream(num: number, dict: string, data: Uint8Array): Promise<void> {
    await this.beginStream(num, dict, data.length);
    await this.write(data);
    await this.endStream();
  }

  /** Open a stream object whose bytes follow through write(): `length` is
   * their count, or the number of an object (allocated, written after the
   * stream) that will hold it. */
  async beginStream(num: number, dict: string, length: number | { ref: number }): Promise<void> {
    await this.beginObject(num);
    await this.write(`<< ${dict} /Length ${typeof length === "number" ? length : `${length.ref} 0 R`} >>\nstream\n`);
  }

  /** Bytes written so far: the difference across a stream's data is its length. */
  get offset(): number {
    return this.position;
  }

  async endStream(): Promise<void> {
    await this.write("\nendstream\nendobj\n");
  }

  /** The cross-reference table and trailer; the stream stays open. */
  async finish(root: number, info: number): Promise<void> {
    const missing = this.offsets.findIndex((offset) => offset === -1);
    if (missing !== -1) throw new Error(`object ${missing + 1} was allocated but never written`);
    const xref = this.position;
    const lines = [`xref\n0 ${this.offsets.length + 1}\n0000000000 65535 f \n`];
    for (const offset of this.offsets) lines.push(`${String(offset).padStart(10, "0")} 00000 n \n`);
    lines.push(`trailer\n<< /Size ${this.offsets.length + 1} /Root ${root} 0 R /Info ${info} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    await this.write(lines.join(""));
    this.out.off("error", this.onError);
  }

  private async beginObject(num: number): Promise<void> {
    if (this.offsets[num - 1] !== -1) throw new Error(`object ${num} written twice, or never allocated`);
    this.offsets[num - 1] = this.position;
    await this.write(`${num} 0 obj\n`);
  }

  /** Bytes, or a string of code points below 256 (PDF syntax is ASCII; text is hex). */
  async write(chunk: string | Uint8Array): Promise<void> {
    if (this.failed) throw this.failed;
    const bytes = typeof chunk === "string" ? Buffer.from(chunk, "latin1") : chunk;
    this.position += bytes.length;
    if (!this.out.write(bytes)) {
      await new Promise<void>((resolve, reject) => {
        const done = (): void => {
          this.out.off("drain", done);
          this.out.off("error", fail);
          resolve();
        };
        const fail = (error: Error): void => {
          this.out.off("drain", done);
          this.out.off("error", fail);
          reject(error);
        };
        this.out.on("drain", done);
        this.out.on("error", fail);
      });
    }
  }
}

/** A NaN or an infinity as num() writes it. */
const NON_FINITE = /(?:^|[\s[\]])-?(?:NaN|Infinity)(?=[\s[\]]|$)/;
/** A content stream's strings: hex and literal. */
const STRINGS = /<[0-9A-Fa-f\s]*>|\((?:\\.|[^\\)])*\)/g;
/** The operators that open and close a graphics state or a text object. */
const NESTING = /(?<![^\s[\]])(?:q|Q|BT|ET)(?![^\s[\]])/g;

/**
 * Content-stream operators with each entry that holds a number num() could
 * not write (NaN or an infinity) left out, all but its q, Q, BT and ET, so
 * the nesting holds. An entry is one drawing: a path and its paint, a run
 * of text, a picture.
 */
export function finiteOps(ops: string[]): string[] {
  return ops.map((op) => {
    const bare = op.replace(STRINGS, " ");
    return NON_FINITE.test(bare) ? (bare.match(NESTING) ?? []).join(" ") : op;
  });
}

/** A PDF literal string, escaped. */
export function pdfString(text: string): string {
  // Text outside Latin-1 goes as UTF-16BE with a byte order mark.
  if (/[^\x00-\xff]/.test(text)) {
    let hex = "FEFF";
    for (let i = 0; i < text.length; i++) hex += text.charCodeAt(i).toString(16).padStart(4, "0");
    return `<${hex}>`;
  }
  return `(${text.replace(/[\\()\r\n]/g, (c) => (c === "\r" ? "\\r" : c === "\n" ? "\\n" : `\\${c}`))})`;
}

/**
 * A number for a content stream or dictionary: short, in plain decimal
 * (PDF has no exponent form), to three places. String() and toFixed() write
 * 1e21 and above in exponent form; every number that large is an integer,
 * which BigInt writes in full. NaN and the infinities have no PDF form:
 * they are written as NaN, Infinity and -Infinity, for finiteOps to leave out.
 */
export function num(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value)) return Math.abs(value) < 1e21 ? String(value) : BigInt(value).toString();
  const fixed = value.toFixed(3).replace(/\.?0+$/, "");
  return fixed === "-0" ? "0" : fixed;
}
