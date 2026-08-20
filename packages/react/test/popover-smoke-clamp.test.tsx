// @vitest-environment jsdom
/**
 * #146 × #148: every popover on the bar stays inside the window, not just the
 * eleven that were reported.
 *
 * `popover-clamp` proved the clamp on the surfaces #148 named and built the
 * browser layout model that makes such a proof possible in jsdom. This runs
 * that SAME model — imported, not copied — over the smoke suite's whole
 * surface table, opened hard against the right edge of the window and then
 * hard against the left. A fix that moved eleven panels and missed the rest
 * is precisely how #148 came to be reported twice.
 *
 * WHY THIS LIVES APART from `popover-smoke-all`. `installLayout` patches
 * `getBoundingClientRect` on the prototype, so it is global while installed.
 * The context surfaces (Table Format, Format) are reached by driving the real
 * engine — placing a caret in a table cell, selecting a shape — and that path
 * reads element geometry, which should get jsdom's zeros rather than a model
 * built for toolbar panels. Keeping the model in its own file keeps it away
 * from them. The surfaces here need no engine beyond an empty document.
 */
import { act } from "react";
import { afterEach, describe, expect, it } from "vitest";
import {
  AT_LEFT_EDGE,
  AT_RIGHT_EDGE,
  expectInsideViewport,
  installLayout,
  MARGIN,
  mount,
  open,
  rectOf,
  selectTab,
  tick,
  VIEWPORT,
  type Mounted,
} from "./popover-layout-model.js";
import { SURFACES } from "./popover-smoke-surfaces.js";

class StubResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = StubResizeObserver;

/**
 * Surfaces a clamp assertion cannot say anything true about.
 *
 * Both are modals: their outermost element is a backdrop covering the whole
 * window, so "does it fit inside the window" is trivially yes and tells you
 * nothing about the dialog centred within it. The SmartArt one is also the
 * single panel in toolbar.tsx whose width jsdom drops outright — `min(560px,
 * calc(100vw - 32px))` carries no `var()`, so `style.width` reads "" and no
 * parser can recover it. Asserting on either would be a green that means
 * nothing, which is the thing this whole suite exists to avoid.
 */
const NOT_CLAMPABLE = new Set([
  "Help and keyboard shortcuts (Ctrl+/)",
  "Insert or edit SmartArt",
]);

const CLAMPABLE = SURFACES.filter((surface) => !NOT_CLAMPABLE.has(surface.tip));

let restoreLayout: (() => void) | null = null;
let mounted: Mounted | null = null;

afterEach(async () => {
  await mounted?.unmount();
  mounted = null;
  restoreLayout?.();
  restoreLayout = null;
});

describe("#146 × #148 · every popover on the bar is clamped, not just the reported ones", () => {
  it("covers the whole surface table apart from the two modals", () => {
    // A filter that quietly emptied itself would turn every test below into a
    // no-op, and the suite would still be green.
    expect(CLAMPABLE.length).toBe(SURFACES.length - NOT_CLAMPABLE.size);
    expect(CLAMPABLE.length).toBeGreaterThan(45);
  });

  /**
   * Proof that the fifty-odd greens below are not vacuous.
   *
   * `expectInsideViewport` already refuses a panel with no width, which is the
   * usual way a geometry assertion stops meaning anything. This closes the
   * other way: a panel pinned at left 0 would also satisfy "inside the
   * window" while proving nothing about clamping. So this pins the EXACT
   * arithmetic — a 236px panel whose control sits at x=864 must be pushed
   * back to 900 - 236 - 8 = 656, not left where it was asked to go, and not
   * flattened to zero.
   */
  it("moves a panel that would overflow, by exactly the amount needed", async () => {
    restoreLayout = installLayout();
    mounted = await mount();
    await selectTab(mounted.bar, "home");
    const panel = await open(mounted.bar, "Text color", AT_RIGHT_EDGE);
    const rect = rectOf(panel);

    expect(rect.width, "the colour menu models at Chromium's 236").toBe(236);
    expect(AT_RIGHT_EDGE.left, "the control really is near the right edge").toBe(864);
    expect(rect.left, "pushed left to fit, not left at the anchor and not at 0")
      .toBe(VIEWPORT.width - 236 - MARGIN);
    // Which is to say: it moved, and it moved because of the clamp.
    expect(rect.left).toBeLessThan(AT_RIGHT_EDGE.left);
    expect(rect.left).toBeGreaterThan(0);
  });

  for (const surface of CLAMPABLE) {
    it(`${surface.tab} › ${surface.tip} — opened at the right edge`, async () => {
      restoreLayout = installLayout();
      mounted = await mount();
      await selectTab(mounted.bar, surface.tab);
      const panel = await open(mounted.bar, surface.tip, AT_RIGHT_EDGE);
      expectInsideViewport(panel, `${surface.tip} at the right edge`);
    });

    it(`${surface.tab} › ${surface.tip} — opened at the left edge`, async () => {
      // Right-aligned panels are the ones at risk here: their right edge is
      // pinned to a control with almost no room to its left.
      restoreLayout = installLayout();
      mounted = await mount();
      await selectTab(mounted.bar, surface.tab);
      const panel = await open(mounted.bar, surface.tip, AT_LEFT_EDGE);
      expectInsideViewport(panel, `${surface.tip} at the left edge`);
    });

    it(`${surface.tab} › ${surface.tip} — still clamped the second time it opens`, async () => {
      // The half of the reopen invariant that `popover-smoke-all` cannot
      // carry, because it has no layout to ask. Placement is computed in an
      // effect from a ref, and a panel that measured itself once and kept the
      // answer would come back in the old place — "works once, then not
      // again", in geometry.
      restoreLayout = installLayout();
      mounted = await mount();
      await selectTab(mounted.bar, surface.tab);

      const first = await open(mounted.bar, surface.tip, AT_RIGHT_EDGE);
      expectInsideViewport(first, `${surface.tip}, first open`);
      const before = rectOf(first);

      await act(async () => {
        document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
      });
      await tick(20);
      expect(first.isConnected, `${surface.tip} did not close on Escape`).toBe(false);

      const second = await open(mounted.bar, surface.tip, AT_RIGHT_EDGE);
      expectInsideViewport(second, `${surface.tip}, second open`);
      expect(rectOf(second), `${surface.tip} came back in a different place`).toEqual(before);
    });
  }
});
