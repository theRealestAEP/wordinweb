import { unzipSync, strFromU8 } from "fflate";

/** Fixed mtime for every zip entry we write (2020-01-01T00:00:00Z). Without
 * it fflate stamps `new Date()` at 2-second DOS granularity, so two save()
 * calls straddling a boundary differ in bytes — breaking byte-identity
 * checks (docHash self-heal, checkpoint dedup, byte-compare tests). B11. */
export const FIXED_ZIP_MTIME = new Date(1577836800000);

/** Decompression-bomb guards (plan doc 11, security F4): a crafted .docx must
 * not exhaust memory via a huge uncompressed payload or a vast number of
 * entries. Generous for real documents. */
export const DEFAULT_ZIP_LIMITS = {
  maxTotalUncompressed: 512 * 1024 * 1024, // 512 MB across all entries
  maxEntries: 20_000,
};

/** A read-only view over the OPC (zip) package inside a .docx file. */
export class Package {
  private files: Record<string, Uint8Array>;

  constructor(data: Uint8Array) {
    this.files = unzipSync(data);
    let total = 0;
    let entries = 0;
    for (const name in this.files) {
      entries++;
      total += this.files[name].length;
      if (entries > DEFAULT_ZIP_LIMITS.maxEntries) {
        throw new Error(`zip has too many entries (> ${DEFAULT_ZIP_LIMITS.maxEntries})`);
      }
      if (total > DEFAULT_ZIP_LIMITS.maxTotalUncompressed) {
        throw new Error(`zip uncompressed size exceeds ${DEFAULT_ZIP_LIMITS.maxTotalUncompressed} bytes`);
      }
    }
  }

  static from(data: ArrayBuffer | Uint8Array): Package {
    return new Package(data instanceof Uint8Array ? data : new Uint8Array(data));
  }

  /** A package over an already-unzipped entry map (DocxDocument.clone). */
  static fromFiles(files: Record<string, Uint8Array>): Package {
    const pkg = Object.create(Package.prototype) as Package;
    pkg.files = files;
    return pkg;
  }

  has(name: string): boolean {
    return this.normalize(name) in this.files;
  }

  binary(name: string): Uint8Array | undefined {
    return this.files[this.normalize(name)];
  }

  text(name: string): string | undefined {
    const bin = this.binary(name);
    return bin ? strFromU8(bin) : undefined;
  }

  names(): string[] {
    return Object.keys(this.files);
  }

  /** Raw entry map (shared, do not mutate) — used for write-back. */
  raw(): Record<string, Uint8Array> {
    return this.files;
  }

  /** OPC part names never start with '/' inside the zip; tolerate both. */
  private normalize(name: string): string {
    return name.startsWith("/") ? name.slice(1) : name;
  }
}

/** Resolve a relationship target relative to a source part. */
/** The relationship Target that reaches `toPart` from `fromPart`'s folder. */
export function relativePartPath(fromPart: string, toPart: string): string {
  const fromDir = fromPart.split("/").slice(0, -1);
  const to = toPart.split("/");
  let i = 0;
  while (i < fromDir.length && i < to.length - 1 && fromDir[i] === to[i]) i++;
  return [...fromDir.slice(i).map(() => ".."), ...to.slice(i)].join("/");
}

export function resolvePartPath(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const baseDir = sourcePart.includes("/")
    ? sourcePart.slice(0, sourcePart.lastIndexOf("/"))
    : "";
  const segments = (baseDir ? baseDir + "/" + target : target).split("/");
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") out.pop();
    else out.push(seg);
  }
  return out.join("/");
}

/**
 * Each entry's uncompressed size, read from the zip's central directory
 * without inflating anything. Throws when the bytes are not a zip.
 */
export function zipEntrySizes(data: Uint8Array): Map<string, number> {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  // The end-of-central-directory record closes the file, after a comment of
  // at most 65535 bytes.
  let end = data.length - 22;
  const stop = Math.max(0, data.length - 22 - 65535);
  while (end >= stop && view.getUint32(end, true) !== 0x06054b50) end--;
  if (end < stop) throw new Error("not a zip: no end-of-central-directory record");
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const sizes = new Map<string, number>();
  const decoder = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (at + 46 > data.length || view.getUint32(at, true) !== 0x02014b50) throw new Error("not a zip: bad central directory");
    const size = view.getUint32(at + 24, true);
    const nameLength = view.getUint16(at + 28, true);
    const extra = view.getUint16(at + 30, true);
    const comment = view.getUint16(at + 32, true);
    // 0xFFFFFFFF marks a ZIP64 size: over 4 GB, far past any limit.
    sizes.set(decoder.decode(data.subarray(at + 46, at + 46 + nameLength)), size === 0xffffffff ? Infinity : size);
    at += 46 + nameLength + extra + comment;
  }
  return sizes;
}
