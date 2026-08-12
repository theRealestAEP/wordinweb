/**
 * Ribbon layout: which toolbar controls stay on the bar's first line, and how
 * the rest are laid out once the user opens them up.
 *
 * The bar used to guess this from window-width tiers. A tier is a guess about
 * how wide the controls are, and the guess was wrong: at some widths the last
 * control on the line was cut in half by the window edge, and at others a
 * group folded away while 200px of bar sat empty. This module answers the same
 * question by arithmetic on measured widths instead, so a control is never
 * placed where it does not fit.
 *
 * Two states, one affordance (the expand chevron):
 *
 * - Collapsed — exactly one line. Controls that do not fit are hidden, lowest
 *   fold priority first, and the chevron's own width is reserved before the
 *   first fit test so the answer cannot change when the chevron appears.
 * - Expanded — nothing is hidden; the controls wrap onto as many lines as they
 *   need. The lines are balanced (see `balancedWidth`) so the last one is never
 *   left holding one stray control.
 *
 * Expanding costs the bar a whole line, so this module also decides whether
 * the chevron is offered at all — see `MIN_REVEAL`. Both rules exist because
 * the affordance was reported for advertising tools and delivering a gap.
 */

/** One measured control on the bar, in document order. */
export interface RibbonItem {
  /** Natural width in px, measured with nothing hidden and no shrinking. */
  width: number;
  /**
   * Fold priority. Higher folds away first; equal priorities fold from the
   * right. Negative values are the controls that stay longest — undo/redo,
   * bold/italic/underline, font and size — so a narrow window keeps a bar
   * you can still write with rather than whichever controls happen to be
   * leftmost.
   */
  fold: number;
  /** A group separator: never allowed to start or end the visible line. */
  separator: boolean;
}

export interface RibbonPlan {
  /** Per item, in the order given: hidden from the collapsed line. */
  hidden: boolean[];
  /** Expanded: the controls take full-width lines below the tab strip. */
  stacked: boolean;
  /** Offer the expand chevron: there is a line's worth of tools behind it. */
  offerExpand: boolean;
  /** How many controls (dividers do not count) the collapsed line hides. */
  hiddenCount: number;
  /**
   * Width to constrain the wrapping container to when expanded, or 0 to leave
   * it at its natural width.
   */
  rowMaxWidth: number;
  /** A single control is wider than the line, so that line must scroll. */
  needsScroll: boolean;
}

export interface RibbonSpace {
  /** Gap between two controls, in px. */
  gap: number;
  /** Width the controls may use while they share a line with the tab strip. */
  inlineAvailable: number;
  /** Width of a line that spans the whole bar (the expanded state). */
  fullAvailable: number;
  /** Width kept free for the expand chevron. */
  reserve: number;
}

/** Width of a line holding every item that is not hidden. */
function lineWidth(items: RibbonItem[], hidden: boolean[], gap: number): number {
  let total = 0;
  let count = 0;
  for (let i = 0; i < items.length; i++) {
    if (hidden[i]) continue;
    total += items[i].width;
    count++;
  }
  return count === 0 ? 0 : total + gap * (count - 1);
}

/** How many lines a greedy wrap needs for `widths` at `width`. */
export function lineCount(widths: number[], gap: number, width: number): number {
  if (widths.length === 0) return 0;
  let lines = 1;
  let used = 0;
  for (const w of widths) {
    const need = used === 0 ? w : used + gap + w;
    if (need <= width || used === 0) used = need;
    else {
      lines++;
      used = w;
    }
  }
  return lines;
}

/**
 * The narrowest width that still wraps into the same number of lines.
 *
 * A greedy wrap fills each line to the brim and leaves the remainder on the
 * last one, which is how a 13-control tab ends up as a full line and then a
 * line holding one lonely icon. Squeezing the container to the narrowest width
 * that keeps the same line count spreads the same controls evenly across those
 * lines instead — same number of lines, no stray tail.
 */
export function balancedWidth(widths: number[], gap: number, width: number): number {
  const lines = lineCount(widths, gap, width);
  if (lines <= 1) return width;
  let low = Math.max(...widths);
  let high = width;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (lineCount(widths, gap, mid) <= lines) high = mid;
    else low = mid + 1;
  }
  return low;
}

/** Hide separators that would start the line, end it, or double up. */
function trimSeparators(items: RibbonItem[], hidden: boolean[]): void {
  const visible = items.map((_, i) => i).filter((i) => !hidden[i]);
  let previousWasSeparator = true; // line start: a leading separator is stray
  for (const i of visible) {
    if (!items[i].separator) {
      previousWasSeparator = false;
      continue;
    }
    if (previousWasSeparator) hidden[i] = true;
    previousWasSeparator = true;
  }
  for (let k = visible.length - 1; k >= 0; k--) {
    const i = visible[k];
    if (hidden[i]) continue;
    if (items[i].separator) hidden[i] = true;
    else break;
  }
}

/**
 * The fewest folded controls worth offering the chevron for.
 *
 * Expanding costs the bar a whole line whatever it reveals, so the trade is
 * only worth making when the line comes back full of tools. Measured in the
 * app across a 1600→700px sweep: the Home tab folds two controls at 1600px
 * (the borders menu and the multilevel list gallery, the last two on the
 * line) and five by 1500px; the Insert tab folds fifteen even at 1600px.
 * Spending a line of the window on those two was the "it does nothing except
 * add a weird space" the affordance was reported for, so the band where a
 * handful fold is the band where the bar says nothing.
 */
export const MIN_REVEAL = 4;

/** Hide controls, lowest fold priority first, until the line fits `budget`. */
function foldToFit(items: RibbonItem[], gap: number, budget: number): boolean[] {
  const hidden = new Array<boolean>(items.length).fill(false);
  const order = items
    .map((_, i) => i)
    .sort((a, b) => items[b].fold - items[a].fold || b - a);
  let next = 0;
  while (next < order.length && lineWidth(items, hidden, gap) > budget) {
    hidden[order[next++]] = true;
  }
  // Folding stops at the first control that fits, which can leave a wide
  // control's worth of empty bar with narrow controls folded behind it — the
  // "too much and not enough space at the same time" the old bar was reported
  // for. Anything still folded that fits the slack comes back, most important
  // first, so the line is used up to its last few pixels.
  for (let k = order.length - 1; k >= 0; k--) {
    const index = order[k];
    if (!hidden[index] || items[index].separator) continue;
    if (lineWidth(items, hidden, gap) + gap + items[index].width <= budget) {
      hidden[index] = false;
    }
  }
  trimSeparators(items, hidden);
  return hidden;
}

/** Folded controls a user would call tools: dividers are not tools. */
function foldedControls(items: RibbonItem[], hidden: boolean[]): number {
  let count = 0;
  for (let i = 0; i < items.length; i++) if (hidden[i] && !items[i].separator) count++;
  return count;
}

/**
 * Decide what the bar shows.
 *
 * `inlineAvailable <= 0` means nothing has been measured yet (a server render,
 * or jsdom, which has no layout): the honest answer there is to hide nothing,
 * because hiding a control the code cannot measure is how controls disappear
 * on a screen that had room for them.
 */
export function planRibbon(items: RibbonItem[], space: RibbonSpace, expanded: boolean): RibbonPlan {
  const none = new Array<boolean>(items.length).fill(false);
  const idle: RibbonPlan = { hidden: none, stacked: false, offerExpand: false, hiddenCount: 0, rowMaxWidth: 0, needsScroll: false };
  if (items.length === 0 || space.inlineAvailable <= 0) return idle;
  if (lineWidth(items, none, space.gap) <= space.inlineAvailable) return idle;

  // The one-line bar, with room kept at its end for the chevron.
  const folded = foldToFit(items, space.gap, space.inlineAvailable - space.reserve);
  const reveal = foldedControls(items, folded);
  if (reveal < MIN_REVEAL) {
    // Too few to spend a line on: no chevron, and therefore no room kept for
    // one, which puts some of those controls back on the bar. Widening the
    // budget can only un-fold, never fold, so `reveal` cannot climb back over
    // the threshold and the answer cannot flip between the two fits.
    const tight = foldToFit(items, space.gap, space.inlineAvailable);
    return { ...idle, hidden: tight, hiddenCount: foldedControls(items, tight) };
  }
  if (!expanded) {
    return { hidden: folded, stacked: false, offerExpand: true, hiddenCount: reveal, rowMaxWidth: 0, needsScroll: false };
  }

  const widths = items.map((item) => item.width);
  // Full-width lines below the tabs, or narrower lines beside them?
  //
  // Stacking buys width but spends a whole line on the tab strip — a line
  // holding a row of tabs and then a void the width of the window, which is
  // the "random weird space" this bar was reported for. So it has to SAVE a
  // line, not merely break even: on a tie the controls stay beside the tabs,
  // where the first line is the collapsed line the user was already looking at
  // and the extra tools land under it. Only when the tab strip has eaten most
  // of the line — a narrow window — does full width win a line back.
  const beside = lineCount(widths, space.gap, space.inlineAvailable);
  const stacked = 1 + lineCount(widths, space.gap, space.fullAvailable) < beside;
  const room = stacked ? space.fullAvailable : space.inlineAvailable;
  const hidden = new Array<boolean>(items.length).fill(false);
  trimSeparators(items, hidden);
  return {
    hidden,
    stacked,
    offerExpand: true,
    hiddenCount: reveal,
    rowMaxWidth: balancedWidth(widths, space.gap, room),
    needsScroll: Math.max(...widths) > room,
  };
}
