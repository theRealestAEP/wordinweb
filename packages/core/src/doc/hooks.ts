// The one place a .doc becomes a .docx for the renderer and the compare:
// the converter with every reader's hooks (review marks, drawings).
import { Zip, ZipDeflate } from "fflate";
import { FIXED_ZIP_MTIME } from "../zip.js";
import { ConvertHooks, DocReader, ExtraParts, convertDoc, openDoc } from "./convert.js";
import { drawingHooks } from "./drawing-hooks.js";
import { reviewHooks } from "./review-xml.js";
import { ArtImage } from "./officeart.js";

/** Several readers' hooks as one: `special` from the first that writes the character, `parts` merged, every other hook from the first that has it. */
export function combineHooks(list: ConvertHooks[]): ConvertHooks {
  const first = <K extends keyof ConvertHooks>(key: K): ConvertHooks[K] => list.find((h) => h[key])?.[key];
  const specials = list.flatMap((h) => (h.special ? [h.special] : []));
  const parts = list.flatMap((h) => (h.parts ? [h.parts] : []));
  return {
    markers: first("markers"),
    runExtra: first("runExtra"),
    wrapRun: first("wrapRun"),
    markExtra: first("markExtra"),
    rowExtra: first("rowExtra"),
    paragraph: first("paragraph"),
    special: specials.length ? (...args) => specials.reduce<string | undefined>((xml, special) => xml ?? special(...args), undefined) : undefined,
    parts: parts.length
      ? (paragraphs) => {
          const all: ExtraParts = { files: {}, documentRels: [], contentTypes: [] };
          for (const p of parts) {
            const extra = p(paragraphs);
            Object.assign(all.files, extra.files);
            all.documentRels.push(...extra.documentRels);
            all.contentTypes.push(...extra.contentTypes);
          }
          return all;
        }
      : undefined,
  };
}

/** Every reader's hooks for one .doc. */
export function docHooks(reader: DocReader, writeMedia?: (name: string, image: ArtImage, size: number) => void): ConvertHooks {
  return combineHooks([reviewHooks(reader), drawingHooks(reader, writeMedia)]);
}

/** A Word 97-2003 .doc as a .docx package's parts, with everything every reader adds. */
export function docPackageFiles(bytes: Uint8Array, writeMedia?: (name: string, image: ArtImage, size: number) => void): Record<string, Uint8Array> {
  const reader = openDoc(bytes);
  return convertDoc(reader, docHooks(reader, writeMedia));
}

/** A Word 97-2003 .doc as a .docx. */
export function docToDocx(bytes: Uint8Array): Uint8Array {
  const chunks: Uint8Array[] = [];
  let size = 0;
  const zip = new Zip((error, chunk) => {
    if (error) throw error;
    chunks.push(chunk);
    size += chunk.length;
  });
  const writePart = (name: string, data: Uint8Array | ArtImage): void => {
    const entry = new ZipDeflate(name);
    entry.mtime = FIXED_ZIP_MTIME;
    zip.add(entry);
    // Bound the compressor's input and output scratch buffers too.
    const push = (bytes: Uint8Array): void => {
      const pieceBytes = 64 << 10;
      for (let at = 0; at < bytes.length; at += pieceBytes) entry.push(bytes.subarray(at, at + pieceBytes));
    };
    push(data instanceof Uint8Array ? data : data.borrowData?.() ?? data.data);
    entry.push(new Uint8Array(0), true);
  };
  const reader = openDoc(bytes);
  // Compressed metafile BLIPs can expand to gigabytes together. Zip each
  // media part as conversion emits it, so only one expansion stays live.
  const files = convertDoc(reader, docHooks(reader, writePart));
  for (const [name, data] of Object.entries(files)) writePart(name, data);
  zip.end();
  const out = new Uint8Array(size);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}
