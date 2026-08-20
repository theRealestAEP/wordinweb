// @vitest-environment jsdom
/**
 * #146: one table-driven smoke test over every popover and dialog the bar can
 * open from a plain document.
 *
 * The measured gap this closes: of the toolbar's panels, four had a test
 * driving them and the rest had none. A panel could therefore ship crashing
 * on open, opening empty, or refusing to close, and the package stayed green.
 * A whole popover built from bare native `<select>`s looked wrong for weeks
 * (#141) and no suite could say so.
 *
 * The invariants in `popover-smoke-invariants` are deliberately cheap,
 * because breadth is the point — each is a bug a user would report on sight.
 * Depth lives in the per-surface suites (citations, styles pane, table
 * properties, review); this file is the floor under all of them.
 *
 * The surfaces that need a selection first — the Table Format and Format tabs
 * — are in `popover-smoke-context`.
 *
 * THIS SUITE IS RED ON PURPOSE. It found 27 real defects on the surfaces it
 * covers, and they are left failing rather than pinned as expected:
 *
 *   - 24 popovers ignore Escape. Only `ToolbarMenuSelect`, `ColorMenu` and
 *     `useAnchoredPopover` register a keydown handler; every panel that rolls
 *     its own `useEffect` listens for `mousedown` alone, so the keyboard has
 *     no way out of them.
 *   - Highlight and Table open galleries built from `<div onClick>`, with no
 *     `role` and no `tabindex`, so neither can be operated from the keyboard
 *     at all — and inserting a table has no other route.
 *
 * The fixes belong to whoever owns `toolbar.tsx`; softening these assertions
 * to reach green would delete the only evidence the defects exist.
 */
import { afterEach, beforeEach, describe } from "vitest";
import {
  mountToolbar,
  selectTab,
  tick,
  watchConsole,
  type ConsoleWatch,
  type MountedToolbar,
} from "./popover-smoke-harness.js";
import { surfaceInvariants } from "./popover-smoke-invariants.js";
import { SURFACES } from "./popover-smoke-surfaces.js";

class StubResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = StubResizeObserver;

let toolbar: MountedToolbar;
let watch: ConsoleWatch;

beforeEach(async () => {
  localStorage.clear();
  watch = watchConsole();
  toolbar = await mountToolbar();
});

afterEach(async () => {
  await toolbar.unmount();
  watch.stop();
});

describe.each(SURFACES)("$tab › $tip", ({ tab, tip, kind }) => {
  beforeEach(async () => {
    await selectTab(toolbar.bar, tab);
    await tick(0);
    // Mounting the view drives the engine asynchronously; React rightly warns
    // that those updates arrived outside act(). That is this file's setup
    // talking, not the bar's, so it is not carried into the checks.
    watch.reset();
  });

  surfaceInvariants(tip, kind, () => ({ bar: toolbar.bar, watch }));
});
