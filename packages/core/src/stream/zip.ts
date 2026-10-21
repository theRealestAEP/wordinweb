import { Inflate, strFromU8 } from "fflate";
import { ByteSource } from "./source.js";
import { DEFAULT_XML_LIMITS } from "../xml.js";
import { DEFAULT_ZIP_LIMITS } from "../zip.js";

/** One entry of a zip's central directory. */
export interface ZipEntry {
  name: string;
  /** 0 stored, 8 deflated. */
  method: number;
  compressedSize: number;
  size: number;
  headerOffset: number;
}

const MAX32 = 0xffffffff;

/**
 * An entry's compressed bytes are fetched this much at a time: the most one
 * request asks for, and the most the server of a file on disk holds for it.
 */
export const FETCH_BYTES = 1 << 20;

/** The inflater takes this much input at a time, so each piece it gives is bounded too: deflate inflates at most 1032:1, so 16 MB at the very most (what an entry that outgrows its declared size can take before it is stopped), and about as much as the input for an image. */
const INFLATE_INPUT = 16 << 10;

/** Converted packages keep binary parts as views or readers, so media can
 * stay compressed in its original BLIP until the renderer asks for it. */
export type PackagePart = Uint8Array | { size: number; read(): Uint8Array };

export class PartsIndex {
  readonly entries = new Map<string, ZipEntry>();

  constructor(private readonly files: Record<string, PackagePart>) {
    for (const [name, part] of Object.entries(files)) {
      const size = part instanceof Uint8Array ? part.length : part.size;
      this.entries.set(name, { name, size, compressedSize: size, method: 0, headerOffset: 0 });
    }
  }

  async raw(entry: ZipEntry): Promise<Uint8Array> {
    return this.read(entry);
  }

  async read(entry: ZipEntry, limit = DEFAULT_ZIP_LIMITS.maxTotalUncompressed): Promise<Uint8Array> {
    if (entry.size > limit) throw new Error(`too large: ${entry.name} is ${entry.size} bytes, over the ${limit} read whole`);
    const part = this.files[entry.name];
    return part instanceof Uint8Array ? part : part.read();
  }

  async pieces(entry: ZipEntry, onPiece: (piece: Uint8Array) => void | Promise<void>): Promise<void> {
    const bytes = await this.read(entry);
    for (let at = 0; at < bytes.length; at += FETCH_BYTES) await onPiece(bytes.subarray(at, at + FETCH_BYTES));
  }

  async text(name: string): Promise<string | undefined> {
    const entry = this.entries.get(name);
    return entry ? strFromU8(await this.read(entry, DEFAULT_XML_LIMITS.maxLength)) : undefined;
  }
}

/**
 * A zip read through its central directory, an entry at a time, so a
 * document's small parts come in without its body: the browser twin of
 * compare-cli's ZipReader.
 */
export class ZipIndex {
  readonly entries = new Map<string, ZipEntry>();

  private constructor(private readonly source: ByteSource) {}

  static async open(source: ByteSource): Promise<ZipIndex> {
    const zip = new ZipIndex(source);
    await zip.readDirectory();
    return zip;
  }

  /** An entry's compressed bytes, as the file holds them. */
  async raw(entry: ZipEntry): Promise<Uint8Array> {
    const out = new Uint8Array(entry.compressedSize);
    let at = 0;
    await this.compressed(entry, (piece) => {
      out.set(piece, at);
      at += piece.length;
    });
    return out;
  }

  /** An entry's content, inflated in memory: at most `limit` bytes, checked
   * against the directory's size before anything is allocated. */
  async read(entry: ZipEntry, limit = DEFAULT_ZIP_LIMITS.maxTotalUncompressed): Promise<Uint8Array> {
    if (entry.size > limit) throw new Error(`too large: ${entry.name} is ${entry.size} bytes, over the ${limit} read whole`);
    const out = new Uint8Array(entry.size);
    let at = 0;
    await this.pieces(entry, (piece) => {
      out.set(piece, at);
      at += piece.length;
    });
    return out;
  }

  /**
   * An entry's content a piece at a time, inflated as its compressed bytes
   * are fetched: nothing holds the entry whole. Stops as soon as the content
   * outgrows the size the directory gives, so a zip bomb stops here. A
   * promise `onPiece` gives back is waited for before the next fetch, so a
   * consumer that writes the pieces on holds the reading back.
   */
  async pieces(entry: ZipEntry, onPiece: (piece: Uint8Array) => void | Promise<void>): Promise<void> {
    let size = 0;
    let pending: Promise<void>[] = [];
    const take = (piece: Uint8Array): void => {
      size += piece.length;
      if (size > entry.size) throw new Error(`${entry.name} is corrupt`);
      const wait = onPiece(piece);
      if (wait) pending.push(wait);
    };
    const settle = async (): Promise<void> => {
      const waits = pending;
      pending = [];
      await Promise.all(waits);
    };
    if (entry.method !== 8) {
      await this.compressed(entry, (piece) => {
        take(piece);
        return settle();
      });
    } else {
      const inflater = new Inflate(take);
      let done = 0;
      await this.compressed(entry, (piece) => {
        for (let at = 0; at < piece.length; at += INFLATE_INPUT) {
          const next = Math.min(at + INFLATE_INPUT, piece.length);
          done += next - at;
          inflater.push(piece.subarray(at, next), done === entry.compressedSize);
        }
        return settle();
      });
    }
    if (size !== entry.size) throw new Error(`${entry.name} is corrupt`);
  }

  /** An entry's content as text, at most the XML limit. */
  async text(name: string): Promise<string | undefined> {
    const entry = this.entries.get(name);
    return entry ? strFromU8(await this.read(entry, DEFAULT_XML_LIMITS.maxLength)) : undefined;
  }

  /** An entry's compressed bytes, FETCH_BYTES at a time. */
  private async compressed(entry: ZipEntry, onPiece: (piece: Uint8Array) => void | Promise<void>): Promise<void> {
    const header = view(await this.source.read(entry.headerOffset, entry.headerOffset + 30));
    if (header.getUint32(0, true) !== 0x04034b50) throw new Error(`bad local header for ${entry.name}`);
    const start = entry.headerOffset + 30 + header.getUint16(26, true) + header.getUint16(28, true);
    const end = start + entry.compressedSize;
    for (let at = start; at < end; at += FETCH_BYTES) await onPiece(await this.source.read(at, Math.min(at + FETCH_BYTES, end)));
  }

  private async readDirectory(): Promise<void> {
    const fileSize = await this.source.size();
    // The end-of-central-directory record closes the file, after a comment of
    // at most 65535 bytes.
    const tailLength = Math.min(fileSize, 22 + 65535);
    const tail = await this.source.read(fileSize - tailLength, fileSize);
    const tailView = view(tail);
    let eocd = tail.length - 22;
    while (eocd >= 0 && tailView.getUint32(eocd, true) !== 0x06054b50) eocd--;
    if (eocd < 0) throw new Error("not a .docx file");
    let count = tailView.getUint16(eocd + 10, true);
    let cdSize = tailView.getUint32(eocd + 12, true);
    let cdOffset = tailView.getUint32(eocd + 16, true);
    if (count === 0xffff || cdSize === MAX32 || cdOffset === MAX32) {
      // ZIP64: a locator just before the record points at the ZIP64 record.
      const locatorAt = fileSize - tailLength + eocd - 20;
      const locator = view(await this.source.read(locatorAt, locatorAt + 20));
      if (locator.getUint32(0, true) !== 0x07064b50) throw new Error("missing ZIP64 locator");
      const recordAt = Number(locator.getBigUint64(8, true));
      const record = view(await this.source.read(recordAt, recordAt + 56));
      if (record.getUint32(0, true) !== 0x06064b50) throw new Error("bad ZIP64 record");
      count = Number(record.getBigUint64(32, true));
      cdSize = Number(record.getBigUint64(40, true));
      cdOffset = Number(record.getBigUint64(48, true));
    }
    const cd = await this.source.read(cdOffset, cdOffset + cdSize);
    const cdView = view(cd);
    let at = 0;
    for (let i = 0; i < count; i++) {
      if (cdView.getUint32(at, true) !== 0x02014b50) throw new Error("bad central directory");
      if (cdView.getUint16(at + 8, true) & 1) throw new Error("encrypted entries are not supported");
      const nameLength = cdView.getUint16(at + 28, true);
      const extraLength = cdView.getUint16(at + 30, true);
      const commentLength = cdView.getUint16(at + 32, true);
      const entry: ZipEntry = {
        name: strFromU8(cd.subarray(at + 46, at + 46 + nameLength)),
        method: cdView.getUint16(at + 10, true),
        compressedSize: cdView.getUint32(at + 20, true),
        size: cdView.getUint32(at + 24, true),
        headerOffset: cdView.getUint32(at + 42, true),
      };
      if (entry.method !== 0 && entry.method !== 8) throw new Error(`${entry.name} uses compression method ${entry.method}`);
      // A ZIP64 extra field holds, in order, each of these that is 0xFFFFFFFF.
      let extra = at + 46 + nameLength;
      const extraEnd = extra + extraLength;
      while (extra + 4 <= extraEnd) {
        const id = cdView.getUint16(extra, true);
        const length = cdView.getUint16(extra + 2, true);
        if (id === 0x0001) {
          let field = extra + 4;
          const next = (): number => {
            const value = Number(cdView.getBigUint64(field, true));
            field += 8;
            return value;
          };
          if (entry.size === MAX32) entry.size = next();
          if (entry.compressedSize === MAX32) entry.compressedSize = next();
          if (entry.headerOffset === MAX32) entry.headerOffset = next();
        }
        extra += 4 + length;
      }
      // The compressed bytes are allocated from this size (raw()): they must be in the file.
      if (entry.headerOffset + entry.compressedSize > fileSize) throw new Error(`bad central directory: ${entry.name} runs past the end of the file`);
      if (!this.entries.has(entry.name)) this.entries.set(entry.name, entry);
      at = extraEnd + commentLength;
    }
  }
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}
