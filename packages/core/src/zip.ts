import { unzipSync, strFromU8 } from "fflate";

/** A read-only view over the OPC (zip) package inside a .docx file. */
export class Package {
  private files: Record<string, Uint8Array>;

  constructor(data: Uint8Array) {
    this.files = unzipSync(data);
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
