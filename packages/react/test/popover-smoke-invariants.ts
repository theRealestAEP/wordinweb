/**
 * #146: the eight checks every popover and dialog on the bar must pass.
 *
 * They live here, in one function, so that the plain bar, the Table Format
 * tab and the Format tab all run the SAME checks rather than three drifting
 * copies of them. A new invariant is added once and every surface in the
 * package gains it.
 *
 * Each check reopens the surface from scratch. That costs a few milliseconds
 * and buys independence: a surface that ignores Escape fails the Escape check
 * only, instead of leaving a panel open and failing the three checks after it
 * as well — one defect must read as one failure.
 */
import { expect, it } from "vitest";
import {
  clickOutside,
  closeSurface,
  focusableControls,
  isOpen,
  openSurface,
  panelSignature,
  pressKey,
  tick,
  unnamedControls,
  type ConsoleWatch,
} from "./popover-smoke-harness.js";

export interface SmokeContext {
  /** The mounted toolbar, on the tab the surface lives on. */
  bar: HTMLElement;
  /** Recording since the surface's tab was selected. */
  watch: ConsoleWatch;
}

/**
 * Declare the eight tests for one surface.
 *
 * `context` is a getter rather than a value because the toolbar is mounted in
 * a `beforeEach`, which has not run when this function is called.
 */
export function surfaceInvariants(tip: string, what: string, context: () => SmokeContext) {
  it("opens when its control is activated, without throwing", async () => {
    const { bar, watch } = context();
    const { panel } = await openSurface(bar, tip);
    expect(isOpen(panel), `"${tip}" put a panel on screen`).toBe(true);
    expect(watch.messages, `opening "${tip}" logged to the console`).toEqual([]);
  });

  it("renders controls the user can reach", async () => {
    const { bar } = context();
    const { panel } = await openSurface(bar, tip);
    expect(
      focusableControls(panel).length,
      `"${tip}" opened a ${what} with no keyboard-reachable control. ` +
        `Panel HTML: ${panel.outerHTML.slice(0, 400)}`,
    ).toBeGreaterThan(0);
  });

  it("gives every control it renders a name", async () => {
    const { bar } = context();
    const { panel } = await openSurface(bar, tip);
    const unnamed = unnamedControls(panel);
    expect(
      unnamed.map((element) => element.outerHTML.slice(0, 140)),
      `"${tip}" renders ${unnamed.length} control(s) that announce as nothing`,
    ).toEqual([]);
  });

  it("closes on Escape", async () => {
    const { bar } = context();
    const { panel } = await openSurface(bar, tip);
    await pressKey(panel, "Escape");
    await tick(0);
    expect(isOpen(panel), `"${tip}" stayed open after Escape`).toBe(false);
  });

  it("closes on a click outside it", async () => {
    const { bar } = context();
    const { panel } = await openSurface(bar, tip);
    await clickOutside(panel);
    expect(isOpen(panel), `"${tip}" stayed open after a click outside it`).toBe(false);
  });

  it("stays quiet through open, close and reopen", async () => {
    const { bar, watch } = context();
    const first = await openSurface(bar, tip);
    expect(await closeSurface(first), `"${tip}" could not be dismissed at all`).toBe(true);
    // Reopening is where a panel that cleaned up badly the first time throws.
    const second = await openSurface(bar, tip);
    await closeSurface(second);
    await tick(0);
    expect(watch.messages, `opening, closing and reopening "${tip}" logged to the console`).toEqual([]);
  });

  it("offers the same thing the second time it is opened", async () => {
    // "Works once, then not again" is the shape of several defects the user
    // has hit, and a panel that keeps state across a close is where it comes
    // from: a form left on step two, a control left disabled, a list left
    // filtered. Nothing is typed into the panel here — the second open should
    // be indistinguishable from the first purely by virtue of being reopened.
    const { bar } = context();
    const first = await openSurface(bar, tip);
    const before = panelSignature(first.panel);
    await closeSurface(first);

    const second = await openSurface(bar, tip);
    expect(panelSignature(second.panel), `"${tip}" offers different controls on its second open`)
      .toEqual(before);
    await closeSurface(second);
  });

  it("dismisses the same way the second time it is opened", async () => {
    // Measured against the FIRST open, not against an absolute. A surface
    // that has never closed on Escape must not fail here as well — that is
    // #151, recorded once. What this catches is a surface that closed one way
    // the first time and stops doing so after a round trip, which is the
    // dismissal half of "works once, then not again".
    const { bar } = context();
    const firstEscape = await opensThenCloses(bar, tip, "escape");
    const secondEscape = await opensThenCloses(bar, tip, "escape");
    expect(secondEscape, `"${tip}" answered Escape differently on its second open`).toBe(firstEscape);

    const firstOutside = await opensThenCloses(bar, tip, "outside");
    const secondOutside = await opensThenCloses(bar, tip, "outside");
    expect(secondOutside, `"${tip}" answered a click outside differently on its second open`)
      .toBe(firstOutside);
  });
}

/** Open the surface, try one dismissal, report whether it worked, and tidy up. */
async function opensThenCloses(bar: HTMLElement, tip: string, how: "escape" | "outside"): Promise<boolean> {
  const opened = await openSurface(bar, tip);
  if (how === "escape") {
    await pressKey(opened.panel, "Escape");
    await tick(0);
  } else {
    await clickOutside(opened.panel);
  }
  const closed = !isOpen(opened.panel);
  await closeSurface(opened);
  return closed;
}
