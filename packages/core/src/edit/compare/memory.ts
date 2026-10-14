import { unzipSync, strFromU8 } from "fflate";
import { zipEntrySizes } from "../../zip.js";
import { TRACKED_CHANGE, WHOLE_TREE } from "./index.js";

/** What comparing two .docx files will need, estimated before loading either. */
export interface CompareMemoryEstimate {
  /**
   * JavaScript heap the compare needs, in bytes: load both documents,
   * compare, and save the result twice. Run the compare where the heap limit
   * (`--max-old-space-size`, or a worker's `maxOldGenerationSizeMb`) is at
   * least this.
   */
  heapBytes: number;
  /**
   * Heap plus what the process holds outside it: both packages' unzipped
   * parts, the saved result, and Node itself. Keep this under the memory of
   * the machine or container the compare runs in.
   */
  totalBytes: number;
}

const MB = 1024 * 1024;

/**
 * The heap a compare needs grows with the XML in the two packages: the parsed
 * trees, the model and the compare's own data are all made from it. Measured
 * as the smallest heap limit that finishes (bench/memory, pairs from 0.3 to
 * 63 MB of XML, light to heavy edits), on a base of 36 MB, in MB per MB of XML:
 *
 * | Path                                   | reuseRevised | copy |
 * | -------------------------------------- | ------------ | ---- |
 * | compareDocuments                       | 9.6          | 13.8 |
 * | compareDocuments, input tracked        | 14.5         | 18.0 |
 * | compareDocx                            | 3.9          |      |
 *
 * An input with tracked changes costs more because its changes are accepted
 * on a copy first. compareDocx goes through compareDocuments with
 * reuseRevised when an input needs the whole tree. These constants sit above
 * every measurement, by at least 7% on the large pairs.
 */
const HEAP_BASE = 40 * MB;
const HEAP_PER_XML = {
  reuseRevised: { clean: 10.5, tracked: 16 },
  copy: { clean: 15, tracked: 20 },
  compareDocx: 4.5,
};
/** Node and V8 themselves, outside the heap (measured 466-477 MB peak RSS at a 320 MB heap limit). */
const PROCESS_BASE = 96 * MB;

/**
 * Estimate the memory a compare of two .docx files needs: compareDocuments,
 * or compareDocx with `compareDocx: true`. No tree is built: the estimate
 * reads the sizes the zip directories list, and searches the XML parts for
 * tracked changes and for what makes compareDocx compare the whole tree.
 * Throws when either file is not a zip.
 */
export function estimateCompareMemory(
  original: Uint8Array,
  revised: Uint8Array,
  options: { reuseRevised?: boolean; compareDocx?: boolean } = {},
): CompareMemoryEstimate {
  let xml = 0;
  let all = 0;
  let tracked = false;
  let wholeTree = false;
  for (const data of [original, revised]) {
    for (const [name, size] of zipEntrySizes(data)) {
      all += size;
      if (/\.(xml|rels)$/i.test(name)) xml += size;
    }
    const parts = unzipSync(data, { filter: (file) => /\.(xml|rels)$/i.test(file.name) });
    let hasDocument = false;
    for (const name in parts) {
      // Latin-1 keeps the markup searchable at one byte per character.
      const text = strFromU8(parts[name], true);
      if (TRACKED_CHANGE.test(text)) tracked = true;
      if (WHOLE_TREE.test(text) || /\/relationships\/chart"/.test(text)) wholeTree = true;
      if (/<w:document[\s>]/.test(text)) hasDocument = true;
    }
    if (!hasDocument) wholeTree = true;
  }
  let perXml: number;
  if (options.compareDocx && !wholeTree) perXml = HEAP_PER_XML.compareDocx;
  else {
    const mode = options.reuseRevised || options.compareDocx ? HEAP_PER_XML.reuseRevised : HEAP_PER_XML.copy;
    perXml = tracked ? mode.tracked : mode.clean;
  }
  const heapBytes = Math.ceil(HEAP_BASE + perXml * xml);
  // Unzipped parts of both packages, about as much again while the result is
  // saved, and the process itself.
  return { heapBytes, totalBytes: heapBytes + 2 * all + PROCESS_BASE };
}
