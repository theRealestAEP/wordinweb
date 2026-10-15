import { closeSync, createReadStream, createWriteStream, fstatSync, openSync, readSync, writeSync } from "node:fs";
import { PassThrough, Transform } from "node:stream";
import { pipeline } from "node:stream/promises";
import { crc32, createDeflateRaw, createInflateRaw, deflateRawSync, inflateRawSync } from "node:zlib";
import { CompareError } from "./errors.js";

/** One entry of a zip's central directory. */
export interface ZipEntry {
  name: string;
  /** 0 stored, 8 deflated. */
  method: number;
  crc: number;
  compressedSize: number;
  size: number;
  headerOffset: number;
}

const MAX32 = 0xffffffff;

/**
 * The most the compare holds in memory, all the parts the shell keeps whole
 * (styles, numbering, headers, footers, charts, custom XML, rels) together
 * per document; read() holds any one of them to the same. The limit is set
 * by the heap, not the disk: a part of tiny elements (a header of empty
 * paragraphs, custom XML of empty elements) costs about 25 MB of heap per
 * MB of XML, so two documents at the limit take about 160 MB and a compare
 * fits in a 256 MB heap with margin. The documents seen so far hold under
 * 1 MB; the body, the notes and the binary parts never come through here.
 */
export const MAX_HELD_BYTES = 3 << 20;

/** The heap MAX_HELD_BYTES was chosen for, as the messages state it. */
export const HEAP_MB = 256;

/** A size for a message, in whole MB. */
export function mb(bytes: number): string {
  return `${Math.ceil(bytes / (1024 * 1024))} MB`;
}

/** A zip file on disk, read through its central directory. */
export class ZipReader {
  readonly entries = new Map<string, ZipEntry>();

  private constructor(
    readonly path: string,
    private readonly fd: number,
  ) {}

  static open(path: string): ZipReader {
    const fd = openSync(path, "r");
    try {
      const zip = new ZipReader(path, fd);
      zip.readDirectory();
      return zip;
    } catch (error) {
      closeSync(fd);
      throw error;
    }
  }

  close(): void {
    closeSync(this.fd);
  }

  /** Where an entry's compressed bytes start. */
  dataOffset(entry: ZipEntry): number {
    const header = this.bytes(entry.headerOffset, 30);
    if (header.readUInt32LE(0) !== 0x04034b50) throw new CompareError("damaged", `${this.path}: bad local header for ${entry.name}`);
    return entry.headerOffset + 30 + header.readUInt16LE(26) + header.readUInt16LE(28);
  }

  /** An entry's content, inflated in memory: at most MAX_HELD_BYTES, checked against the directory's size before anything is read. */
  read(entry: ZipEntry): Buffer {
    if (entry.size > MAX_HELD_BYTES) {
      throw new CompareError("too_large", `${this.path}: ${entry.name} is ${mb(entry.size)}; the parts held in memory are at most ${mb(MAX_HELD_BYTES)} together, so that a compare fits in a ${HEAP_MB} MB heap`);
    }
    const raw = this.bytes(this.dataOffset(entry), entry.compressedSize);
    let data: Buffer;
    try {
      // Never inflate past the size the directory gives: a zip bomb stops here.
      data = entry.method === 8 ? inflateRawSync(raw, { maxOutputLength: Math.max(1, entry.size) }) : raw;
    } catch {
      throw new CompareError("damaged", `${this.path}: ${entry.name} is corrupt`);
    }
    if (data.length !== entry.size || crc32(data) !== entry.crc) throw new CompareError("damaged", `${this.path}: ${entry.name} is corrupt`);
    return data;
  }

  /** Inflate an entry into a file, checking its size and CRC. */
  async extract(entry: ZipEntry, dest: string): Promise<void> {
    const start = this.dataOffset(entry);
    let size = 0;
    let crc = 0;
    const path = this.path;
    const check = new Transform({
      transform(chunk: Buffer, _encoding, done) {
        size += chunk.length;
        crc = crc32(chunk, crc);
        // Stop as soon as the entry outgrows the size the directory gives.
        if (size > entry.size) done(new CompareError("damaged", `${path}: ${entry.name} is corrupt`));
        else done(null, chunk);
      },
    });
    // createReadStream's `end` is inclusive, and a zero-length range must not read byte `start`.
    const source = entry.compressedSize > 0 ? createReadStream(this.path, { start, end: start + entry.compressedSize - 1 }) : emptyStream();
    await pipeline(source, entry.method === 8 ? createInflateRaw() : new PassThrough(), check, createWriteStream(dest));
    if (size !== entry.size || crc !== entry.crc) throw new CompareError("damaged", `${this.path}: ${entry.name} is corrupt`);
  }

  /** Copy an entry's compressed bytes, unchanged, into `out` under `name`. */
  copyInto(out: ZipWriter, entry: ZipEntry, name = entry.name): void {
    let at = this.dataOffset(entry);
    const end = at + entry.compressedSize;
    out.addEntry(name, entry.method, entry.crc, entry.compressedSize, entry.size, (write) => {
      const chunk = Buffer.allocUnsafe(Math.min(1 << 20, entry.compressedSize));
      while (at < end) {
        const n = readSync(this.fd, chunk, 0, Math.min(chunk.length, end - at), at);
        if (n === 0) throw new CompareError("damaged", `${this.path}: ${entry.name} is truncated`);
        write(chunk.subarray(0, n));
        at += n;
      }
    });
  }

  private bytes(position: number, length: number): Buffer {
    const buf = Buffer.allocUnsafe(length);
    let done = 0;
    while (done < length) {
      const n = readSync(this.fd, buf, done, length - done, position + done);
      if (n === 0) throw new CompareError("damaged", `${this.path}: unexpected end of file`);
      done += n;
    }
    return buf;
  }

  private readDirectory(): void {
    const fileSize = fstatSync(this.fd).size;
    // The end-of-central-directory record closes the file, after a comment of
    // at most 65535 bytes.
    const tailLength = Math.min(fileSize, 22 + 65535);
    const tail = this.bytes(fileSize - tailLength, tailLength);
    let eocd = tail.length - 22;
    while (eocd >= 0 && tail.readUInt32LE(eocd) !== 0x06054b50) eocd--;
    if (eocd < 0) {
      // A .doc file, and a password-protected .docx, are OLE compound files.
      const ole = fileSize >= 8 && this.bytes(0, 8).equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
      throw new CompareError("not_docx", `${this.path}: not a .docx file${ole ? " (it is a .doc file or a password-protected document)" : ""}`);
    }
    let count = tail.readUInt16LE(eocd + 10);
    let cdSize = tail.readUInt32LE(eocd + 12);
    let cdOffset = tail.readUInt32LE(eocd + 16);
    if (count === 0xffff || cdSize === MAX32 || cdOffset === MAX32) {
      // ZIP64: a locator just before the record points at the ZIP64 record.
      const locatorAt = fileSize - tailLength + eocd - 20;
      const locator = this.bytes(locatorAt, 20);
      if (locator.readUInt32LE(0) !== 0x07064b50) throw new CompareError("damaged", `${this.path}: missing ZIP64 locator`);
      const record = this.bytes(Number(locator.readBigUInt64LE(8)), 56);
      if (record.readUInt32LE(0) !== 0x06064b50) throw new CompareError("damaged", `${this.path}: bad ZIP64 record`);
      count = Number(record.readBigUInt64LE(32));
      cdSize = Number(record.readBigUInt64LE(40));
      cdOffset = Number(record.readBigUInt64LE(48));
    }
    const cd = this.bytes(cdOffset, cdSize);
    let at = 0;
    for (let i = 0; i < count; i++) {
      if (cd.readUInt32LE(at) !== 0x02014b50) throw new CompareError("damaged", `${this.path}: bad central directory`);
      if (cd.readUInt16LE(at + 8) & 1) throw new CompareError("not_docx", `${this.path}: encrypted entries are not supported`);
      const nameLength = cd.readUInt16LE(at + 28);
      const extraLength = cd.readUInt16LE(at + 30);
      const commentLength = cd.readUInt16LE(at + 32);
      const entry: ZipEntry = {
        name: cd.toString("utf8", at + 46, at + 46 + nameLength),
        method: cd.readUInt16LE(at + 10),
        crc: cd.readUInt32LE(at + 16),
        compressedSize: cd.readUInt32LE(at + 20),
        size: cd.readUInt32LE(at + 24),
        headerOffset: cd.readUInt32LE(at + 42),
      };
      if (entry.method !== 0 && entry.method !== 8) throw new CompareError("not_docx", `${this.path}: ${entry.name} uses compression method ${entry.method}`);
      // A ZIP64 extra field holds, in order, each of these that is 0xFFFFFFFF.
      let extra = at + 46 + nameLength;
      const extraEnd = extra + extraLength;
      while (extra + 4 <= extraEnd) {
        const id = cd.readUInt16LE(extra);
        const length = cd.readUInt16LE(extra + 2);
        if (id === 0x0001) {
          let field = extra + 4;
          const next = (): number => {
            const value = Number(cd.readBigUInt64LE(field));
            field += 8;
            return value;
          };
          if (entry.size === MAX32) entry.size = next();
          if (entry.compressedSize === MAX32) entry.compressedSize = next();
          if (entry.headerOffset === MAX32) entry.headerOffset = next();
        }
        extra += 4 + length;
      }
      if (!this.entries.has(entry.name)) this.entries.set(entry.name, entry);
      at = extraEnd + commentLength;
    }
  }
}

function emptyStream(): PassThrough {
  const stream = new PassThrough();
  stream.end();
  return stream;
}

/** 2020-01-01 00:00, as DOS date and time: what core's save() stamps on every entry. */
const DOS_DATE = ((2020 - 1980) << 9) | (1 << 5) | 1;
const DOS_TIME = 0;

/** A zip written to disk one entry at a time, with ZIP64 records where sizes need them. */
export class ZipWriter {
  private readonly fd: number;
  private offset = 0;
  private readonly central: Buffer[] = [];
  private count = 0;

  constructor(path: string) {
    this.fd = openSync(path, "w");
  }

  /** Add a part held in memory, deflated. */
  addBuffer(name: string, data: Uint8Array): void {
    const compressed = deflateRawSync(data);
    this.addEntry(name, 8, crc32(data), compressed.length, data.length, (write) => write(compressed));
  }

  /** Add a file, deflated through a temporary file next to it. */
  async addFile(name: string, path: string): Promise<void> {
    let size = 0;
    let crc = 0;
    const check = new Transform({
      transform(chunk: Buffer, _encoding, done) {
        size += chunk.length;
        crc = crc32(chunk, crc);
        done(null, chunk);
      },
    });
    const deflated = `${path}.deflated`;
    await pipeline(createReadStream(path), check, createDeflateRaw(), createWriteStream(deflated));
    const fd = openSync(deflated, "r");
    try {
      const compressedSize = fstatSync(fd).size;
      this.addEntry(name, 8, crc, compressedSize, size, (write) => {
        const chunk = Buffer.allocUnsafe(1 << 20);
        for (let at = 0; at < compressedSize; ) {
          const n = readSync(fd, chunk, 0, chunk.length, at);
          write(chunk.subarray(0, n));
          at += n;
        }
      });
    } finally {
      closeSync(fd);
    }
  }

  /** Add an entry whose compressed bytes `writeData` writes. */
  addEntry(name: string, method: number, crc: number, compressedSize: number, size: number, writeData: (write: (chunk: Buffer) => void) => void): void {
    const nameBytes = Buffer.from(name, "utf8");
    const headerOffset = this.offset;
    const big = compressedSize >= MAX32 || size >= MAX32;
    const localExtra = big ? zip64Extra([size, compressedSize]) : Buffer.alloc(0);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(big ? 45 : 20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(DOS_TIME, 10);
    local.writeUInt16LE(DOS_DATE, 12);
    local.writeUInt32LE(crc >>> 0, 14);
    local.writeUInt32LE(big ? MAX32 : compressedSize, 18);
    local.writeUInt32LE(big ? MAX32 : size, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(localExtra.length, 28);
    this.write(local);
    this.write(nameBytes);
    this.write(localExtra);
    const dataStart = this.offset;
    writeData((chunk) => this.write(chunk));
    if (this.offset - dataStart !== compressedSize) throw new Error(`zip: ${name} wrote ${this.offset - dataStart} bytes, expected ${compressedSize}`);

    const wide = [size >= MAX32 ? size : null, compressedSize >= MAX32 ? compressedSize : null, headerOffset >= MAX32 ? headerOffset : null];
    const centralExtra = wide.some((v) => v !== null) ? zip64Extra(wide.filter((v): v is number => v !== null)) : Buffer.alloc(0);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0);
    entry.writeUInt16LE(45, 4);
    entry.writeUInt16LE(centralExtra.length > 0 ? 45 : 20, 6);
    entry.writeUInt16LE(0x0800, 8);
    entry.writeUInt16LE(method, 10);
    entry.writeUInt16LE(DOS_TIME, 12);
    entry.writeUInt16LE(DOS_DATE, 14);
    entry.writeUInt32LE(crc >>> 0, 16);
    entry.writeUInt32LE(Math.min(compressedSize, MAX32), 20);
    entry.writeUInt32LE(Math.min(size, MAX32), 24);
    entry.writeUInt16LE(nameBytes.length, 28);
    entry.writeUInt16LE(centralExtra.length, 30);
    entry.writeUInt32LE(Math.min(headerOffset, MAX32), 42);
    this.central.push(entry, nameBytes, centralExtra);
    this.count++;
  }

  /** Write the central directory and close the file. */
  close(): void {
    const cdOffset = this.offset;
    for (const part of this.central) this.write(part);
    const cdSize = this.offset - cdOffset;
    const zip64 = this.count >= 0xffff || cdOffset >= MAX32 || cdSize >= MAX32;
    if (zip64) {
      const recordAt = this.offset;
      const record = Buffer.alloc(56);
      record.writeUInt32LE(0x06064b50, 0);
      record.writeBigUInt64LE(44n, 4);
      record.writeUInt16LE(45, 12);
      record.writeUInt16LE(45, 14);
      record.writeBigUInt64LE(BigInt(this.count), 24);
      record.writeBigUInt64LE(BigInt(this.count), 32);
      record.writeBigUInt64LE(BigInt(cdSize), 40);
      record.writeBigUInt64LE(BigInt(cdOffset), 48);
      this.write(record);
      const locator = Buffer.alloc(20);
      locator.writeUInt32LE(0x07064b50, 0);
      locator.writeBigUInt64LE(BigInt(recordAt), 8);
      locator.writeUInt32LE(1, 16);
      this.write(locator);
    }
    const end = Buffer.alloc(22);
    end.writeUInt32LE(0x06054b50, 0);
    end.writeUInt16LE(zip64 ? 0xffff : this.count, 8);
    end.writeUInt16LE(zip64 ? 0xffff : this.count, 10);
    end.writeUInt32LE(zip64 ? MAX32 : cdSize, 12);
    end.writeUInt32LE(zip64 ? MAX32 : cdOffset, 16);
    this.write(end);
    closeSync(this.fd);
  }

  private write(chunk: Buffer): void {
    let done = 0;
    while (done < chunk.length) done += writeSync(this.fd, chunk, done, chunk.length - done, this.offset + done);
    this.offset += chunk.length;
  }
}

function zip64Extra(values: number[]): Buffer {
  const extra = Buffer.alloc(4 + 8 * values.length);
  extra.writeUInt16LE(0x0001, 0);
  extra.writeUInt16LE(8 * values.length, 2);
  values.forEach((value, i) => extra.writeBigUInt64LE(BigInt(value), 4 + 8 * i));
  return extra;
}
