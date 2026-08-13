// @vitest-environment jsdom
/**
 * #158: no width may put a control out of reach.
 *
 * The bar folds controls it cannot fit and offers a chevron to get them back.
 * Those two decisions drifted apart: below `MIN_REVEAL` the layout hid the
 * overflow AND suppressed the chevron, so at 900px Hyphenation on Layout and
 * Compare Documents on Review were simply gone — no chevron, no menu, no
 * shortcut. Reported as controls that cannot be reached at all.
 *
 * WHY THIS COUNTS REACHABILITY AND NOT CONTROLS. A count would have passed
 * Review, which shows the right NUMBER of controls at some widths while
 * hiding two different ones. What matters to a user is whether the tool they
 * want can be got at — visible now, or visible after clicking the chevron —
 * so that is the set this compares, name by name, against the same tab at its
 * widest.
 */
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { measureLikeABrowser, resizeTo, ribbon, toolbarRoot } from "./toolbar-measure.js";

class StubResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = StubResizeObserver;
measureLikeABrowser();

function stubApi(): DocxViewApi {
  const methods = {
    getSelectedObjectContext: () => null,
    getTableCellFill: () => undefined,
    getSelectionFormat: () => ({ bold: false, italic: false, underline: false, strike: false }),
    getParagraphStyleId: () => null,
    getListType: () => null,
    listParagraphStyles: () => [{ id: "Heading1", name: "Heading 1" }],
    listStyles: () => [],
    imageAccept: () => "image/png,image/jpeg",
  };
  return new Proxy(methods, {
    get(target, property) {
      return property in target ? target[property as keyof typeof target] : () => null;
    },
  }) as unknown as DocxViewApi;
}

async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

async function mountToolbar() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(DocxToolbar, { api: stubApi() }));
  });
  return {
    container,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
}

/**
 * What this control is called, from the user's point of view. `title` moves
 * to `data-tip` on first hover, so both are read.
 */
function controlName(el: HTMLElement): string | null {
  return el.getAttribute("title") ?? el.getAttribute("data-tip") ?? el.getAttribute("aria-label");
}

/**
 * Every control the bar is RENDERING right now, named.
 *
 * Counted as buttons rather than as ribbon children on purpose: a tab is free
 * to wrap its controls in a group, and the Layout and Review tabs do. Reading
 * the children would then have called the whole Layout tab one control and
 * quietly compared a one-element set against itself at every width — a suite
 * that could not fail on the very tab the defect was reported against.
 *
 * A folded control is hidden by a `display: none` somewhere at or above it,
 * so the walk up to the ribbon is what decides whether it is on screen.
 */
function shownControls(container: HTMLElement): string[] {
  const bar = ribbon(container);
  const hidden = (el: HTMLElement): boolean => {
    for (let node: HTMLElement | null = el; node && node !== bar; node = node.parentElement) {
      if (node.style.display === "none") return true;
    }
    return false;
  };
  return Array.from(bar.querySelectorAll<HTMLElement>("button"))
    .filter((el) => el.dataset.dxwToolbarExpand === undefined && !hidden(el))
    .map(controlName)
    .filter((name): name is string => name !== null);
}

function expandChevron(container: HTMLElement): HTMLButtonElement | null {
  return toolbarRoot(container).querySelector<HTMLButtonElement>("[data-dxw-toolbar-expand]");
}

/**
 * Every control the user can get to at this width: what is on the bar, plus
 * whatever the chevron reveals. This is the whole point of the suite — a
 * control behind a chevron is reachable, a control behind nothing is not.
 */
async function reachableControls(container: HTMLElement, width: number): Promise<string[]> {
  await resizeTo(container, width);
  const collapsed = shownControls(container);
  const chevron = expandChevron(container);
  if (!chevron) return [...new Set(collapsed)].sort();
  await click(chevron);
  const expanded = shownControls(container);
  return [...new Set([...collapsed, ...expanded])].sort();
}

const TABS = ["home", "insert", "draw", "layout", "review"] as const;
/** Widest first: the reference set is taken at 1400. */
const WIDTHS = [1400, 1100, 900, 700] as const;

beforeEach(() => {
  localStorage.clear();
});

describe("#158 · every control stays reachable at every width", () => {
  for (const tab of TABS) {
    it(`${tab} keeps all its controls reachable from 1400 down to 700`, async () => {
      const reference: string[] = [];
      for (const width of WIDTHS) {
        // A fresh bar per width: the expanded choice persists in
        // localStorage, and carrying it between widths would measure the
        // previous case rather than this one.
        const t = await mountToolbar();
        try {
          await click(t.container.querySelector<HTMLButtonElement>(`button[data-tab="${tab}"]`)!);
          const reachable = await reachableControls(t.container, width);

          if (width === WIDTHS[0]) {
            reference.push(...reachable);
            // A reference that collapsed to nothing would make every
            // comparison below trivially true.
            expect(reference.length, `${tab} has controls to lose at its widest`).toBeGreaterThan(4);
            continue;
          }

          const missing = reference.filter((name) => !reachable.includes(name));
          expect(
            missing,
            `${tab} at ${width}px: ${missing.length} control(s) cannot be reached — ` +
              `not on the bar, and not behind the chevron`,
          ).toEqual([]);
        } finally {
          await t.unmount();
        }
      }
    });
  }

  it("offers the chevron on any tab at any width that hides something", async () => {
    // The invariant behind the defect, stated directly: a control may be
    // folded only if there is a way to get it back. MIN_REVEAL is allowed to
    // prefer fitting everything on one line; it is not allowed to lose a tool.
    for (const tab of TABS) {
      for (const width of WIDTHS) {
        const t = await mountToolbar();
        try {
          await click(t.container.querySelector<HTMLButtonElement>(`button[data-tab="${tab}"]`)!);
          await resizeTo(t.container, width);
          const collapsed = shownControls(t.container);
          const chevron = expandChevron(t.container);
          if (chevron) continue; // something to get them back with: fine

          await click(t.container.querySelector<HTMLButtonElement>(`button[data-tab="${tab}"]`)!);
          const widest = await reachableControls(t.container, 1400);
          const lost = widest.filter((name) => !collapsed.includes(name));
          expect(
            lost,
            `${tab} at ${width}px hides ${lost.length} control(s) and offers no chevron`,
          ).toEqual([]);
        } finally {
          await t.unmount();
        }
      }
    }
  });
});
