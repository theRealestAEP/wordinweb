/**
 * #146: the five checks every popover and dialog on the bar must pass.
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
  pressKey,
  tick,
  type ConsoleWatch,
} from "./popover-smoke-harness.js";

export interface SmokeContext {
  /** The mounted toolbar, on the tab the surface lives on. */
  bar: HTMLElement;
  /** Recording since the surface's tab was selected. */
  watch: ConsoleWatch;
}

/**
 * Declare the five tests for one surface.
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
}
