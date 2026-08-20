import { describe, expect, it } from "vitest";
import { MIN_REVEAL, planRibbon, type RibbonItem } from "../src/ribbon-layout.js";

/**
 * The arithmetic behind the toolbar's two states. These are the guarantees a
 * width sweep in a real window is supposed to demonstrate, stated as maths so
 * a regression names itself instead of showing up as a screenshot nobody
 * looked at: nothing is ever placed where it does not fit, the expanded bar
 * never ends on a line holding one stray control, and the chevron is only
 * offered when the line it costs comes back full of tools.
 */

const GAP = 2;
const RESERVE = 30;

/** A Home tab at its real measured widths (px, rounded), in bar order. */
const HOME: RibbonItem[] = [
  { width: 32, fold: -2, separator: false }, // undo
  { width: 32, fold: -2, separator: false }, // redo
  { width: 9, fold: 0, separator: true },
  { width: 92, fold: 0, separator: false },  // paragraph style
  { width: 30, fold: 0, separator: false },  // styles pane
  { width: 30, fold: 0, separator: false },  // format painter
  { width: 130, fold: -1, separator: false }, // font
  { width: 58, fold: -1, separator: false },  // size
  { width: 9, fold: 0, separator: true },
  { width: 30, fold: -3, separator: false },  // bold
  { width: 30, fold: -3, separator: false },  // italic
  { width: 30, fold: -3, separator: false },  // underline
  { width: 30, fold: 0, separator: false },   // strikethrough
  { width: 30, fold: 0, separator: false },   // superscript
  { width: 30, fold: 0, separator: false },   // subscript
  { width: 30, fold: 0, separator: false },   // text effects
  { width: 30, fold: 0, separator: false },   // clear formatting
  { width: 52, fold: 0, separator: false },   // change case
  { width: 30, fold: 0, separator: false },   // text color
  { width: 30, fold: 0, separator: false },   // highlight
  { width: 9, fold: 0, separator: true },
  { width: 30, fold: 0, separator: false },   // align left
  { width: 30, fold: 0, separator: false },   // center
  { width: 30, fold: 0, separator: false },   // align right
  { width: 30, fold: 0, separator: false },   // justify
  { width: 9, fold: 0, separator: true },
  { width: 30, fold: 0, separator: false },   // decrease indent
  { width: 30, fold: 0, separator: false },   // increase indent
  { width: 30, fold: 0, separator: false },   // tab stops
  { width: 44, fold: 0, separator: false },   // line spacing
  { width: 30, fold: 0, separator: false },   // bulleted list
  { width: 30, fold: 0, separator: false },   // numbered list
  { width: 44, fold: 0, separator: false },   // multilevel list
  { width: 9, fold: 0, separator: true },
  { width: 44, fold: 0, separator: false },   // borders
  { width: 9, fold: 0, separator: true },
];

/** Width of the line the plan leaves visible. */
function visibleWidth(items: RibbonItem[], hidden: boolean[]): number {
  const shown = items.filter((_, i) => !hidden[i]);
  return shown.reduce((sum, item) => sum + item.width, 0) + GAP * Math.max(0, shown.length - 1);
}

function collapsed(items: RibbonItem[], available: number) {
  return planRibbon(items, { gap: GAP, inlineAvailable: available, reserve: RESERVE }, false);
}

function opened(items: RibbonItem[], available: number) {
  return planRibbon(items, { gap: GAP, inlineAvailable: available, reserve: RESERVE }, true);
}

/** The controls that share the bar's first line, by index. */
function firstLine(items: RibbonItem[], plan: { hidden: boolean[]; revealed: boolean[] }): number[] {
  return items.map((_, i) => i).filter((i) => !plan.hidden[i] && !plan.revealed[i]);
}

describe("the collapsed line", () => {
  // The window widths the bar is used at, minus the tab strip and padding:
  // what is actually left for the controls.
  const AVAILABLE = [1240, 1140, 1040, 940, 840, 740, 640, 540, 440, 340, 240];

  it("never places a control where it does not fit, at any width", () => {
    for (const available of AVAILABLE) {
      const plan = collapsed(HOME, available);
      const budget = plan.offerExpand ? available - RESERVE : available;
      expect(visibleWidth(HOME, plan.hidden), `${available}px`).toBeLessThanOrEqual(budget);
    }
  });

  it("keeps writing controls longest as the window narrows", () => {
    const bold = HOME.findIndex((item) => item.fold === -3);
    const borders = HOME.length - 2;
    // 240px is narrower than the Home tab has ever been asked to be.
    const plan = collapsed(HOME, 240);
    expect(plan.hidden[bold], "bold survives").toBe(false);
    expect(plan.hidden[borders], "borders folds").toBe(true);
  });

  it("folds nothing when everything fits, and offers no chevron", () => {
    const plan = collapsed(HOME, 2000);
    expect(plan.offerExpand).toBe(false);
    expect(plan.hiddenCount).toBe(0);
    expect(plan.hidden.some(Boolean)).toBe(false);
  });

  it("leaves no separator stranded at either end of the line", () => {
    for (const available of AVAILABLE) {
      const plan = collapsed(HOME, available);
      const shown = HOME.filter((_, i) => !plan.hidden[i]);
      expect(shown[0]?.separator, `${available}px starts on a divider`).not.toBe(true);
      expect(shown[shown.length - 1]?.separator, `${available}px ends on a divider`).not.toBe(true);
      for (let i = 1; i < shown.length; i++) {
        expect(shown[i].separator && shown[i - 1].separator, `${available}px doubles a divider`).toBe(false);
      }
    }
  });

  it("hides nothing before anything has been measured", () => {
    // jsdom and server renders have no layout. Folding on a width of zero
    // would hide every control on a screen that had room for all of them.
    const plan = collapsed(HOME, 0);
    expect(plan.offerExpand).toBe(false);
    expect(plan.hidden.some(Boolean)).toBe(false);
  });

  it("spends the reserve on tools rather than showing a chevron for nothing", () => {
    // A hair too narrow: the ONLY reason anything folded is the width kept
    // free for a chevron. Give that width back and every control fits, so the
    // bar stays quiet — the "it does nothing except add a weird space" case.
    const total = HOME.reduce((sum, item) => sum + item.width, 0) + GAP * (HOME.length - 1);
    const plan = collapsed(HOME, total - 5);
    expect(plan.offerExpand, "a chevron hiding nothing is worse than no chevron").toBe(false);
    expect(plan.hiddenCount, "and nothing was lost to earn it").toBe(0);
    expect(visibleWidth(HOME, plan.hidden)).toBeLessThanOrEqual(total - 5);
    // The whole line is the budget: the bar keeps no seat for a chevron it is
    // not going to show, so it lands on the same line a bar with no chevron
    // at all would.
    const noSeat = planRibbon(HOME, { gap: GAP, inlineAvailable: total - 5, reserve: 0 }, false);
    expect(plan.hidden).toEqual(noSeat.hidden);
  });

  it("offers the chevron for even one folded control, because it is the only way back", () => {
    // #158. This branch used to return the tight line with `offerExpand`
    // false, reasoning that a whole line is too much to spend on one or two
    // controls. The reasoning was about the COST of expanding and never asked
    // what became of the controls: at 900px it hid Hyphenation on Layout and
    // Compare Documents on Review with no chevron, no menu and no shortcut to
    // reach them by. MIN_REVEAL may prefer fitting everything on one line; it
    // may not lose a tool.
    const total = HOME.reduce((sum, item) => sum + item.width, 0) + GAP * (HOME.length - 1);
    const plan = collapsed(HOME, total - 60);
    expect(plan.hiddenCount, "this width really does have to fold something").toBeGreaterThan(0);
    expect(plan.hiddenCount, "and fewer than the threshold, which is the whole point")
      .toBeLessThan(MIN_REVEAL);
    expect(plan.offerExpand, "folded with no way back is the defect").toBe(true);
    // Offering it means paying for it: the line keeps the chevron's seat.
    expect(visibleWidth(HOME, plan.hidden)).toBeLessThanOrEqual(total - 60 - RESERVE);
  });

  it("never hides a control without offering the chevron, at any width", () => {
    // The invariant the two cases above are instances of, swept across every
    // width the bar is used at. A control may be folded only if there is a
    // way to get it back.
    for (const available of AVAILABLE) {
      const plan = collapsed(HOME, available);
      if (plan.hiddenCount === 0) continue;
      expect(plan.offerExpand, `${available}px folds ${plan.hiddenCount} with no chevron`).toBe(true);
    }
  });

  it("offers the chevron once a line's worth of tools is behind it", () => {
    const total = HOME.reduce((sum, item) => sum + item.width, 0) + GAP * (HOME.length - 1);
    const plan = collapsed(HOME, total - 300);
    expect(plan.offerExpand).toBe(true);
    expect(plan.hiddenCount).toBeGreaterThanOrEqual(MIN_REVEAL);
    // Dividers are not tools: the count is what the chevron's label promises.
    expect(plan.hiddenCount).toBeLessThan(plan.hidden.filter(Boolean).length);
    expect(visibleWidth(HOME, plan.hidden)).toBeLessThanOrEqual(total - 300 - RESERVE);
  });

  it("answers the same at a width however the window got there", () => {
    // What the old tier table needed hysteresis for. The plan is a function
    // of the measured width alone, so a window shrinking through 640px and
    // one growing through it land on the identical bar.
    const shrinking = [1240, 940, 740, 640].map((w) => collapsed(HOME, w));
    const growing = [240, 440, 540, 640].map((w) => collapsed(HOME, w));
    expect(shrinking[3].hidden).toEqual(growing[3].hidden);
  });
});

describe("the expanded lines", () => {
  it("leaves the collapsed line exactly as it was", () => {
    // The reported bug, as arithmetic: "why is it bringing items from the top
    // that are already available down a line?" Opening the bar used to re-wrap
    // every control across balanced lines, which squeezed the first line and
    // pushed controls the user could already see down a row. Opening may only
    // ADD rows.
    for (const available of [1240, 940, 740, 640, 540, 440, 340]) {
      const shut = collapsed(HOME, available);
      if (!shut.offerExpand) continue;
      const open = opened(HOME, available);
      const before = HOME.map((_, i) => i).filter((i) => !shut.hidden[i]);
      expect(firstLine(HOME, open), `at ${available}px`).toEqual(before);
    }
  });

  it("reveals every folded control, and hides none of them", () => {
    const available = 740;
    const shut = collapsed(HOME, available);
    const open = opened(HOME, available);
    HOME.forEach((item, i) => {
      if (!shut.hidden[i] || item.separator) return;
      expect(open.revealed[i], "a folded control is revealed below").toBe(true);
      expect(open.hidden[i], "and is not hidden as well").toBe(false);
    });
    // Every control is reachable once open; only dividers may still go.
    expect(HOME.filter((item, i) => open.hidden[i] && !item.separator)).toEqual([]);
    expect(open.offerExpand).toBe(true);
  });

  it("strands no divider at the edge of either region", () => {
    // The two regions are two lines, so each one owns its own ends. A divider
    // that fell at the end of the collapsed line goes; so does one that would
    // now start the revealed row beneath it.
    const open = opened(HOME, 740);
    for (const region of [firstLine(HOME, open), HOME.map((_, i) => i).filter((i) => open.revealed[i] && !open.hidden[i])]) {
      expect(region.length).toBeGreaterThan(0);
      expect(HOME[region[0]].separator, "no leading divider").toBe(false);
      expect(HOME[region[region.length - 1]].separator, "no trailing divider").toBe(false);
    }
  });

  it("reports a control too wide for the line so it can scroll", () => {
    const wide: RibbonItem[] = [...HOME, { width: 900, fold: 0, separator: false }];
    expect(opened(wide, 400).needsScroll).toBe(true);
    expect(opened(HOME, 400).needsScroll).toBe(false);
  });
});
