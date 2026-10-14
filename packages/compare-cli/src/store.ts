import { closeSync, openSync, readSync, writeSync } from "node:fs";

const PENDING = 1 << 20;

/**
 * Strings appended to a file and read back by index. Memory holds two
 * numbers per string, a megabyte of unwritten bytes and a read buffer the
 * size of the longest string read: reads and writes reuse their buffers, so
 * a long run of them leaves no garbage outside the heap.
 */
export class BlockStore {
  private readonly fd: number;
  private readonly offsets: number[] = [];
  private readonly lengths: number[] = [];
  private written = 0;
  private readonly pending = Buffer.allocUnsafe(PENDING);
  private pendingBytes = 0;
  private scratch = Buffer.allocUnsafe(64 * 1024);
  count = 0;

  constructor(readonly path: string) {
    this.fd = openSync(path, "w+");
  }

  add(text: string): number {
    const length = Buffer.byteLength(text, "utf8");
    this.offsets.push(this.written + this.pendingBytes);
    this.lengths.push(length);
    if (this.pendingBytes + length > PENDING) this.flush();
    if (length > PENDING) {
      writeAll(this.fd, Buffer.from(text, "utf8"), this.written);
      this.written += length;
    } else {
      this.pending.write(text, this.pendingBytes, "utf8");
      this.pendingBytes += length;
    }
    return this.count++;
  }

  read(i: number): string {
    return this.load(i).toString("utf8");
  }

  /** Pass string i's bytes to `use`. They are valid only until the next read. */
  withBytes(i: number, use: (bytes: Buffer) => void): void {
    use(this.load(i));
  }

  close(): void {
    closeSync(this.fd);
  }

  private load(i: number): Buffer {
    this.flush();
    const length = this.lengths[i];
    if (length > this.scratch.length) this.scratch = Buffer.allocUnsafe(Math.max(length, this.scratch.length * 2));
    const buf = this.scratch.subarray(0, length);
    let done = 0;
    while (done < length) done += readSync(this.fd, buf, done, length - done, this.offsets[i] + done);
    return buf;
  }

  private flush(): void {
    if (this.pendingBytes === 0) return;
    writeAll(this.fd, this.pending.subarray(0, this.pendingBytes), this.written);
    this.written += this.pendingBytes;
    this.pendingBytes = 0;
  }
}

function writeAll(fd: number, bytes: Buffer, position: number): void {
  let done = 0;
  while (done < bytes.length) done += writeSync(fd, bytes, done, bytes.length - done, position + done);
}
