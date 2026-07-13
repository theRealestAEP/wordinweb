/**
 * Grapheme-cluster boundaries for caret movement and deletion. Complex scripts
 * store a "user-perceived character" as several code units — a Devanagari
 * conjunct (क + ् + ष), an Arabic base + harakāt, a Thai consonant + vowel/tone
 * marks, an emoji ZWJ sequence, or a surrogate pair (astral CJK, emoji). Word
 * moves the caret and Backspace by whole clusters, never landing inside one, so
 * we step on grapheme boundaries rather than UTF-16 code units.
 */

let segmenter: Intl.Segmenter | null = null;
let segmenterTried = false;

function getSegmenter(): Intl.Segmenter | null {
  if (segmenterTried) return segmenter;
  segmenterTried = true;
  try {
    if (typeof Intl !== "undefined" && typeof Intl.Segmenter === "function") {
      segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
    }
  } catch {
    segmenter = null;
  }
  return segmenter;
}

/**
 * Ascending list of grapheme-cluster boundaries in `text`, always including 0
 * and text.length. Falls back to surrogate-pair-aware code-point boundaries
 * where Intl.Segmenter is unavailable (older engines).
 */
export function graphemeBoundaries(text: string): number[] {
  if (!text) return [0];
  const seg = getSegmenter();
  const bounds: number[] = [0];
  if (seg) {
    for (const { index, segment } of seg.segment(text)) bounds.push(index + segment.length);
    return bounds;
  }
  let i = 0;
  while (i < text.length) {
    const cp = text.codePointAt(i)!;
    i += cp > 0xffff ? 2 : 1;
    bounds.push(i);
  }
  return bounds;
}

/**
 * The next grapheme boundary in `text` moving `delta` (+1 forward / -1 back)
 * from `offset`, or null when there is no boundary in that direction (offset is
 * at the start moving back, or the end moving forward) — the caller then steps
 * to the neighbouring text run. Offsets not on a boundary snap to the nearest
 * boundary in the direction of travel, so a mid-cluster caret can't get stuck.
 */
export function graphemeStep(text: string, offset: number, delta: -1 | 1): number | null {
  const bounds = graphemeBoundaries(text);
  if (delta > 0) {
    for (const b of bounds) if (b > offset) return b;
    return null;
  }
  for (let i = bounds.length - 1; i >= 0; i--) if (bounds[i] < offset) return bounds[i];
  return null;
}
