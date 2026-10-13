import { diffArrays } from "diff";

/**
 * Histogram alignment over two sequences of opaque keys.
 *
 * This is JGit's histogram diff, which git later took as `--histogram`. It
 * builds an occurrence histogram of the left side and, among the common
 * elements, anchors on the one with the LOWEST occurrence count, then recurses
 * into the region on either side.
 *
 * Why not the two obvious alternatives:
 *
 *  - **Myers / plain LCS** has no notion of which matches are meaningful, so it
 *    anchors happily on a blank paragraph or a repeated heading and produces a
 *    shifted alignment that reads as "everything changed".
 *  - **Patience** anchors only on elements that are unique on BOTH sides. A
 *    document whose three chapters each open with "Introduction" has no unique
 *    anchor in that region at all, so patience degrades to plain diff exactly
 *    where the help is needed. Histogram still prefers the least-repeated
 *    candidate and keeps aligning. (JGit: histogram "behaves exactly like Bram
 *    Cohen's patience diff whenever there is a unique common element
 *    available… When no unique elements exist, the lowest occurrence element is
 *    chosen instead.")
 *
 * The LCS fallback for a region whose elements are all too common is delegated
 * to the `diff` package (BSD-3-Clause), which is the only thing the dependency
 * is used for: sequence alignment over arrays of keys. None of its document
 * model is adopted.
 */

export interface KeyMatch {
  /** Index into the left sequence. */
  a: number;
  /** Index into the right sequence. */
  b: number;
}

/**
 * An element appearing more times than this in a region is not worth chasing
 * as an anchor; the region falls back to the library LCS. JGit uses 64.
 */
const MAX_CHAIN = 64;

/**
 * Matched index pairs, strictly increasing on both sides. Everything not
 * returned is a deletion (left) or an insertion (right).
 */
export function histogramMatch(a: readonly string[], b: readonly string[]): KeyMatch[] {
  const out: KeyMatch[] = [];
  region(a, b, 0, a.length, 0, b.length, out);
  return out;
}

function region(
  a: readonly string[],
  b: readonly string[],
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number,
  out: KeyMatch[],
): void {
  // Common prefix and suffix are unambiguous matches; peeling them first is
  // both cheaper and more stable than letting the anchor search find them.
  while (aLo < aHi && bLo < bHi && a[aLo] === b[bLo]) {
    out.push({ a: aLo++, b: bLo++ });
  }
  const suffix: KeyMatch[] = [];
  while (aLo < aHi && bLo < bHi && a[aHi - 1] === b[bHi - 1]) {
    suffix.push({ a: --aHi, b: --bHi });
  }
  if (aLo < aHi && bLo < bHi) {
    const best = lowestOccurrenceRegion(a, b, aLo, aHi, bLo, bHi);
    if (best) {
      region(a, b, aLo, best.aStart, bLo, best.bStart, out);
      for (let k = 0; k < best.len; k++) out.push({ a: best.aStart + k, b: best.bStart + k });
      region(a, b, best.aStart + best.len, aHi, best.bStart + best.len, bHi, out);
    } else if (best === null) {
      // Every common element is over MAX_CHAIN: hand the region to the library
      // LCS rather than chasing hundreds of equally bad anchors.
      lcsFallback(a, b, aLo, aHi, bLo, bHi, out);
    }
    // `undefined` means the region has no common element at all — nothing to
    // match, so it is a pure delete plus a pure insert.
  }
  for (let i = suffix.length - 1; i >= 0; i--) out.push(suffix[i]);
}

interface Anchor {
  aStart: number;
  bStart: number;
  len: number;
}

/**
 * JGit's anchor choice: among common regions seeded by elements no more
 * repeated than the best found so far, keep the LONGER one, or the one whose
 * least-repeated element is rarer. Equal regions break on relative position,
 * so an anchor near where it "should" be wins over an identical one far away.
 *
 * A region is an exact run of equal elements grown out from its seed, so a
 * single rare element in it already proves where it belongs. Charging a region
 * for its MOST-repeated element instead (as this once did) penalised exactly
 * those runs: a true match through a section that crossed one boilerplate
 * paragraph lost to a short run of twice-repeated paragraphs paired with the
 * wrong section, and the true blocks came out deleted and reinserted (the NIH
 * contract's "The ... for this ... is $ ." boilerplate, compared against a
 * copy with scattered edits).
 *
 * Returns `undefined` when the two regions share no element at all, and `null`
 * when they share only elements above MAX_CHAIN (the caller's LCS fallback).
 */
function lowestOccurrenceRegion(
  a: readonly string[],
  b: readonly string[],
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number,
): Anchor | null | undefined {
  const positions = new Map<string, number[]>();
  for (let i = aLo; i < aHi; i++) {
    const list = positions.get(a[i]);
    if (list) list.push(i);
    else positions.set(a[i], [i]);
  }

  let best: Anchor | undefined;
  let bestCost = Infinity;
  let bestOffset = Infinity;
  let sawCommon = false;
  let sawUsable = false;
  const aSpan = aHi - aLo;
  const bSpan = bHi - bLo;

  for (let j = bLo; j < bHi; j++) {
    const candidates = positions.get(b[j]);
    if (!candidates || candidates.length === 0) continue;
    sawCommon = true;
    if (candidates.length > MAX_CHAIN) continue;
    sawUsable = true;
    // A seed more common than the best region's rarest element cannot lead to
    // a better anchor.
    if (best !== undefined && candidates.length > bestCost) continue;
    for (const i of candidates) {
      // Grow the match outward from (i, j) as far as both sides agree.
      let aStart = i;
      let bStart = j;
      while (aStart > aLo && bStart > bLo && a[aStart - 1] === b[bStart - 1]) {
        aStart--;
        bStart--;
      }
      let aEnd = i + 1;
      while (aEnd < aHi && bStart + (aEnd - aStart) < bHi && a[aEnd] === b[bStart + (aEnd - aStart)]) {
        aEnd++;
      }
      const len = aEnd - aStart;
      // The region's cost is its least-repeated element.
      let cost = candidates.length;
      for (let k = aStart; k < aEnd; k++) {
        cost = Math.min(cost, positions.get(a[k])?.length ?? cost);
      }
      const offset = Math.abs((aStart - aLo) / (aSpan || 1) - (bStart - bLo) / (bSpan || 1));
      const better =
        best !== undefined &&
        (len > best.len || cost < bestCost || (len === best.len && cost === bestCost && offset < bestOffset));
      if (best === undefined || better) {
        best = { aStart, bStart, len };
        bestCost = cost;
        bestOffset = offset;
      }
    }
  }
  if (best) return best;
  if (sawCommon && !sawUsable) return null;
  return undefined;
}

function lcsFallback(
  a: readonly string[],
  b: readonly string[],
  aLo: number,
  aHi: number,
  bLo: number,
  bHi: number,
  out: KeyMatch[],
): void {
  let ai = aLo;
  let bi = bLo;
  for (const part of diffArrays(a.slice(aLo, aHi) as string[], b.slice(bLo, bHi) as string[])) {
    const n = part.value.length;
    if (part.added) bi += n;
    else if (part.removed) ai += n;
    else {
      for (let k = 0; k < n; k++) out.push({ a: ai + k, b: bi + k });
      ai += n;
      bi += n;
    }
  }
}
