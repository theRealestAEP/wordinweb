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
 * - Expanded — that same line, untouched, and the folded controls on the lines
 *   below it.
 *
 * **Expanding adds tools below and changes nothing else.** The bar used to
 * re-wrap every control across balanced lines when it opened, which squeezed
 * the first line and pushed controls the user was already pointing at down a
 * row. Reported as "why is it bringing items from the top that are already
 * available down a line?" — and it is the same complaint as the "random weird
 * space" before it: the affordance moved the furniture instead of adding to it.
 * So the collapsed line is now frozen on open, and `revealed` says which
 * controls go under it.
 *
 * The controls stay beside the tab strip when open. Giving them the full bar
 * width buys room for the revealed rows, but only by moving the collapsed line
 * below the tabs — motion, and a tab strip followed by a window-wide void.
 * A narrow window therefore gets more revealed rows rather than a bar that
 * rearranges itself at some widths and not others.
 *
 * Expanding costs the bar a line whatever it reveals, so this module also
 * decides whether the chevron is offered at all — see `MIN_REVEAL`.
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
  /** Per item, in the order given: not rendered at all. */
  hidden: boolean[];
  /**
   * Expanded: laid out below the collapsed line rather than on it. Empty of
   * `true` when collapsed. The caller renders these after a forced line break,
   * which is what keeps the line above them from re-wrapping.
   */
  revealed: boolean[];
  /** Offer the expand chevron: there is a line's worth of tools behind it. */
  offerExpand: boolean;
  /** How many controls (dividers do not count) the collapsed line hides. */
  hiddenCount: number;
  /** A single control is wider than the line, so that line must scroll. */
  needsScroll: boolean;
}

export interface RibbonSpace {
  /** Gap between two controls, in px. */
  gap: number;
  /** Width the controls may use while they share a line with the tab strip. */
  inlineAvailable: number;
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

/**
 * Hide separators that would start a line, end it, or double up.
 *
 * `region` is the run of items that share the line — the whole bar when
 * collapsed, and each of the two regions separately once the bar is open, so
 * the divider left dangling at the end of the frozen line is trimmed there
 * without hiding the one that legitimately starts the row beneath it.
 */
function trimSeparators(items: RibbonItem[], hidden: boolean[], region?: number[]): void {
  const visible = (region ?? items.map((_, i) => i)).filter((i) => !hidden[i]);
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
  const idle: RibbonPlan = {
    hidden: none,
    revealed: new Array<boolean>(items.length).fill(false),
    offerExpand: false,
    hiddenCount: 0,
    needsScroll: false,
  };
  if (items.length === 0 || space.inlineAvailable <= 0) return idle;
  if (lineWidth(items, none, space.gap) <= space.inlineAvailable) return idle;

  // The one-line bar, with room kept at its end for the chevron.
  const folded = foldToFit(items, space.gap, space.inlineAvailable - space.reserve);
  const reveal = foldedControls(items, folded);
  if (reveal < MIN_REVEAL) {
    // Too few to spend a line on, so try to avoid the trade entirely: drop
    // the chevron, give its reserved width back to the controls, and see
    // whether they all fit. Widening the budget can only un-fold, never fold.
    const tight = foldToFit(items, space.gap, space.inlineAvailable);
    if (foldedControls(items, tight) === 0) {
      // Everything fits without the chevron. Nothing is lost, so nothing is
      // offered — this is the "it does nothing except add a weird space" case.
      return { ...idle, hidden: tight };
    }
    // Something stays folded. MIN_REVEAL is a preference about whether to
    // spend a line, NEVER a licence to lose a control: this branch used to
    // return `tight` with `offerExpand: false`, which hid Hyphenation on
    // Layout and Compare Documents on Review at 900px with no way back to
    // them — no chevron, no menu, no shortcut (#158). Falling through offers
    // the chevron, and `folded` is the line that keeps room for it.
  }
  if (!expanded) {
    return { ...idle, hidden: folded, offerExpand: true, hiddenCount: reveal };
  }

  // Open. The collapsed line is exactly what it was a moment ago — the same
  // controls, in the same places — and everything it folded moves to the rows
  // underneath. Note that the folded set is NOT a suffix: foldToFit hides by
  // priority, so it can hide a control from the middle of the bar and keep a
  // later one. The caller must therefore ORDER the revealed controls below,
  // not merely break the line at some index; there is no index that separates
  // these two sets.
  const revealed = [...folded];
  const hidden = new Array<boolean>(items.length).fill(false);
  const above: number[] = [];
  const below: number[] = [];
  for (let i = 0; i < items.length; i++) (revealed[i] ? below : above).push(i);
  trimSeparators(items, hidden, above);
  trimSeparators(items, hidden, below);
  return {
    hidden,
    revealed,
    offerExpand: true,
    hiddenCount: reveal,
    // The revealed rows get the same width the collapsed line had, so a
    // control too wide for that line is too wide for these too.
    needsScroll: Math.max(...items.map((item) => item.width)) > space.inlineAvailable,
  };
}
