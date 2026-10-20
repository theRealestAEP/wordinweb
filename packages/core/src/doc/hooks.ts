// The one place a .doc becomes a .docx for the renderer and the compare:
// the converter with every reader's hooks (review marks, drawings).
import { zipSync } from "fflate";
import { FIXED_ZIP_MTIME } from "../zip.js";
import { ConvertHooks, DocReader, ExtraParts, convertDoc, openDoc } from "./convert.js";
import { drawingHooks } from "./drawing-hooks.js";
import { reviewHooks } from "./review-xml.js";

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
export function docHooks(reader: DocReader): ConvertHooks {
  return combineHooks([reviewHooks(reader), drawingHooks(reader)]);
}

/** A Word 97-2003 .doc as a .docx package's parts, with everything every reader adds. */
export function docPackageFiles(bytes: Uint8Array): Record<string, Uint8Array> {
  const reader = openDoc(bytes);
  return convertDoc(reader, docHooks(reader));
}

/** A Word 97-2003 .doc as a .docx. */
export function docToDocx(bytes: Uint8Array): Uint8Array {
  return zipSync(docPackageFiles(bytes), { mtime: FIXED_ZIP_MTIME });
}
