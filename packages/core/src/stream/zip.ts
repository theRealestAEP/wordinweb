import { inflateSync, strFromU8 } from "fflate";
import { ByteSource } from "./source.js";

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
    const header = view(await this.source.read(entry.headerOffset, entry.headerOffset + 30));
    if (header.getUint32(0, true) !== 0x04034b50) throw new Error(`bad local header for ${entry.name}`);
    const start = entry.headerOffset + 30 + header.getUint16(26, true) + header.getUint16(28, true);
    return this.source.read(start, start + entry.compressedSize);
  }

  /** An entry's content, inflated in memory. */
  async read(entry: ZipEntry): Promise<Uint8Array> {
    return inflateEntry(entry, await this.raw(entry));
  }

  async text(name: string): Promise<string | undefined> {
    const entry = this.entries.get(name);
    return entry ? strFromU8(await this.read(entry)) : undefined;
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
      if (!this.entries.has(entry.name)) this.entries.set(entry.name, entry);
      at = extraEnd + commentLength;
    }
  }
}

/** An entry's content from its compressed bytes. */
export function inflateEntry(entry: ZipEntry, raw: Uint8Array): Uint8Array {
  // Never inflate past the size the directory gives: a zip bomb stops here.
  const data = entry.method === 8 ? inflateSync(raw, { out: new Uint8Array(entry.size) }) : raw;
  if (data.length !== entry.size) throw new Error(`${entry.name} is corrupt`);
  return data;
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}
