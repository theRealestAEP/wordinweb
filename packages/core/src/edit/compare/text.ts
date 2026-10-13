/**
 * Text primitives for document comparison: normalization, word tokenization
 * and the two similarity measures the gap pass uses.
 *
 * Every threshold here is a tuning knob, not a law. The values are recorded in
 * COMPARE_TUNING with the reasoning that settled them; see compare/index.ts for
 * how each one is used.
 */

/**
 * Tuning constants, gathered so they can be read (and argued with) in one
 * place. Each was chosen from the literature and then settled against the
 * fixture corpus in test/compare-corpus.test.ts.
 */
export const COMPARE_TUNING = {
  /**
   * Dice-over-word-bigrams score at or above which two paragraphs in the same
   * gap are "the same paragraph, edited". GumTree's `min-dice` default
   * (Falleri et al., ASE'14), which was tuned on syntax trees; prose tolerates
   * it well because the gap bound already stops distant look-alikes.
   */
  minDice: 0.5,
  /**
   * Below this token count a bigram set is too small to mean anything (a
   * three-word heading has two bigrams), so short paragraphs fall back to a
   * normalized Levenshtein ratio instead. Levenshtein is affordable exactly
   * because these strings are short.
   */
  shortParagraphTokens: 5,
  /** Levenshtein ratio a short paragraph pair must reach instead of minDice. */
  minShortRatio: 0.8,
  /**
   * Or the share of words a short pair has in common (Dice over words), which
   * a one-word edit in a short heading or list item keeps at half or more
   * ("Bullet one" / "Bullet updated", "Alignment and flow" / "Alignment and
   * flow as amended") while its character ratio falls under minShortRatio.
   * Below it, those came out struck and reinserted whole: 99% of the compare
   * benchmark's whole-paragraph misses were paragraphs under five words.
   */
  minShortWordDice: 0.5,
  /**
   * ABSOLUTE floor for the coalescing pass: a matched run shorter than this,
   * sitting between two changes in the inner word diff, is absorbed into the
   * surrounding change however small those changes are. This is the stopword
   * rule — one or two matched words between two edits are noise.
   */
  coalesceWindow: 3,
  /**
   * RELATIVE rule for the same pass: a matched run is also absorbed when it is
   * less than 1/`coalesceRatio` of the SMALLER change beside it. Without this a
   * rewritten sentence still comes out as confetti — the old and new wording
   * share "shortage of test", the absolute floor lets three matched words
   * through, and the reviewer gets two strike/insert pairs where they made one
   * edit. With it, the run has to be small RELATIVE to what changed around it,
   * so two one-word edits three words apart stay two edits (3 is not small
   * next to 1) while a rewritten clause becomes one (3 is small next to 7).
   *
   * Both numbers were settled against test/compare-corpus.test.ts; see the
   * calibration note there for what each fixture pins down. The ratio was 2
   * while spaces rode on words: "hardware," and "hardware " differed, so the
   * rewritten sentence of fixture 3 shared only three words. Matched exactly,
   * it shares four ("shortage of test hardware") beside seven changed, which
   * 1.5 absorbs (6 < 7) and 2 does not.
   */
  coalesceRatio: 1.5,
} as const;

/** Collapse whitespace runs and trim, for fingerprinting and similarity. */
export function normalizeText(s: string): string {
  return s.replace(/\s+/g, " ").trim();
}

/**
 * Split into tokens: a word is letters, marks and digits, with inner
 * apostrophes ("don't"); a run of whitespace is one token; every other
 * character is a token of its own. The split is lossless.
 *
 * Whitespace once rode on the word before it ("word "), so a word at the end
 * of a paragraph or before punctuation ("noqu", "hardware,") never matched the
 * same word followed by a space: appending " as replaced" to a paragraph
 * struck its last word and reinserted it, the commonest over-marking in the
 * compare benchmark. Word marks only the words that changed and leaves the
 * spaces between them alone, and so does this. The coalescing pass counts
 * words, not spaces (see changeSize), so a rewrite still reads as one change.
 * Keyed with its punctuation
 * attached, "writing." against "writing," struck and reinserted the whole
 * word ("writing.writing," to a reviewer), and a bold "Purpose:" whose colon
 * moved into the next run never matched "Purpose" at all.
 */
export function wordTokens(s: string): string[] {
  return s.match(/[\p{L}\p{M}\p{N}]+(?:['\u2019][\p{L}\p{M}\p{N}]+)*|[^\s\p{L}\p{M}\p{N}]|\s+/gu) ?? [];
}

/** Lower-cased word tokens of the normalized string, for similarity only. */
function similarityTokens(s: string): string[] {
  const n = normalizeText(s).toLowerCase();
  return n.length === 0 ? [] : n.split(" ");
}

function bigrams(tokens: string[]): Set<string> {
  const out = new Set<string>();
  for (let i = 0; i + 1 < tokens.length; i++) out.add(tokens[i] + "\u0000" + tokens[i + 1]);
  return out;
}

/** Dice coefficient 2·|A∩B| / (|A|+|B|) over two sets. */
function dice(a: Set<string>, b: Set<string>): number {
  if (a.size === 0 && b.size === 0) return 1;
  if (a.size === 0 || b.size === 0) return 0;
  let shared = 0;
  for (const x of a) if (b.has(x)) shared++;
  return (2 * shared) / (a.size + b.size);
}

/** Levenshtein distance, two-row DP. Only called on short strings. */
function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  let prev = new Array<number>(b.length + 1);
  let cur = new Array<number>(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    cur[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
    }
    const swap = prev;
    prev = cur;
    cur = swap;
  }
  return prev[b.length];
}

/**
 * How alike two paragraphs are, in [0, 1]. Short paragraphs (under
 * COMPARE_TUNING.shortParagraphTokens words on EITHER side) are scored by
 * the words they share or, when it is high, their normalized Levenshtein
 * ratio, because a bigram set that small is noise; everything else is scored
 * by Dice over word bigrams.
 *
 * The caller compares the result against `similarityFloor` for the same pair,
 * which is the threshold matching whichever measure was used.
 */
export function similarity(a: string, b: string): number {
  const ta = similarityTokens(a);
  const tb = similarityTokens(b);
  if (ta.length === 0 && tb.length === 0) return 1;
  if (isShortPair(ta.length, tb.length)) {
    const na = normalizeText(a).toLowerCase();
    const nb = normalizeText(b).toLowerCase();
    const longest = Math.max(na.length, nb.length);
    if (longest === 0) return 1;
    const ratio = 1 - levenshtein(na, nb) / longest;
    // Either measure may carry the pair; scored on one scale whose floor is
    // minShortWordDice, a ratio counts only once it reaches its own floor.
    const words = dice(new Set(ta), new Set(tb));
    return Math.max(words, ratio >= COMPARE_TUNING.minShortRatio ? ratio : 0);
  }
  return dice(bigrams(ta), bigrams(tb));
}

/** The score `similarity(a, b)` must reach for the pair to count as a match. */
export function similarityFloor(a: string, b: string): number {
  const la = similarityTokens(a).length;
  const lb = similarityTokens(b).length;
  return isShortPair(la, lb) ? COMPARE_TUNING.minShortWordDice : COMPARE_TUNING.minDice;
}

function isShortPair(la: number, lb: number): boolean {
  return Math.min(la, lb) < COMPARE_TUNING.shortParagraphTokens;
}
