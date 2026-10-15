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

  /** Open a stream object whose `length` bytes follow through write(). */
  async beginStream(num: number, dict: string, length: number): Promise<void> {
    await this.beginObject(num);
    await this.write(`<< ${dict} /Length ${length} >>\nstream\n`);
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

/** A number for a content stream or dictionary: short, never exponent form. */
export function num(value: number): string {
  if (Number.isInteger(value)) return String(value);
  const fixed = value.toFixed(3);
  return fixed.replace(/\.?0+$/, "") || "0";
}
