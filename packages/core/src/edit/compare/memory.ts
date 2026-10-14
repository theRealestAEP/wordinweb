import { zipEntrySizes } from "../../zip.js";

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
 * as the smallest heap limit that finishes (bench/memory, 27 pairs from 0.3 to
 * 63 MB of XML, light to heavy edits): 5.8 to 9.6 MB per MB of XML with
 * reuseRevised and 9.9 to 13.8 without, on a base of 36 MB. These constants
 * sit above every measurement, by at least 7% on the large pairs.
 */
const HEAP_BASE = 40 * MB;
const HEAP_PER_XML = { reuseRevised: 10.5, copy: 15 };
/** Node and V8 themselves, outside the heap (measured 466-477 MB peak RSS at a 320 MB heap limit). */
const PROCESS_BASE = 96 * MB;

/**
 * Estimate the memory compareDocuments needs for two .docx files, from the
 * sizes their zip directories list. Nothing is unzipped, so it is cheap to
 * run before deciding whether to compare at all. Throws when either file is
 * not a zip.
 */
export function estimateCompareMemory(
  original: Uint8Array,
  revised: Uint8Array,
  options: { reuseRevised?: boolean } = {},
): CompareMemoryEstimate {
  let xml = 0;
  let all = 0;
  for (const data of [original, revised]) {
    for (const [name, size] of zipEntrySizes(data)) {
      all += size;
      if (/\.(xml|rels)$/i.test(name)) xml += size;
    }
  }
  const perXml = options.reuseRevised ? HEAP_PER_XML.reuseRevised : HEAP_PER_XML.copy;
  const heapBytes = Math.ceil(HEAP_BASE + perXml * xml);
  // Unzipped parts of both packages, about as much again while the result is
  // saved, and the process itself.
  return { heapBytes, totalBytes: heapBytes + 2 * all + PROCESS_BASE };
}
