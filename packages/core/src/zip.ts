import { unzipSync, strFromU8 } from "fflate";

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
