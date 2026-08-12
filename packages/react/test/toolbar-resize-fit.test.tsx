// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { availableWidth, controlWidth, measureLikeABrowser, resizeTo, ribbon, visibleControls } from "./toolbar-measure.js";

/**
 * The bar across a width sweep. This file replaces the tier-hysteresis test:
 * there are no tiers left to flap between, because the bar no longer guesses
 * from breakpoints — it measures its controls and fits them. The property the
 * hysteresis gap used to fake (the same width always giving the same bar) now
 * falls out of the arithmetic, and the property the tiers never had (nothing
 * cut off at the window edge) is the thing worth pinning.
 *
 * A user reported the failure this guards: at a narrow window the Home tab
 * ended with "Size" sliced in half by the right edge of the window.
 */

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
    listParagraphStyles: () => [],
    listStyles: () => [],
    imageAccept: () => "image/png,image/jpeg",
  };
  return new Proxy(methods, {
    get(target, property) {
      return property in target ? target[property as keyof typeof target] : () => null;
    },
  }) as unknown as DocxViewApi;
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

async function selectTab(container: HTMLElement, name: string) {
  const tab = container.querySelector<HTMLButtonElement>(`button[data-tab="${name}"]`);
  expect(tab, `${name} tab`).toBeTruthy();
  await act(async () => {
    tab!.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

/** Width the visible controls occupy, gaps included. */
function lineWidth(container: HTMLElement): number {
  const shown = visibleControls(container);
  return shown.reduce((sum, el) => sum + controlWidth(el), 0) + 2 * Math.max(0, shown.length - 1);
}

/** The controls on the bar, named by their tooltip. */
function names(container: HTMLElement): string[] {
  return visibleControls(container).map(
    (el) =>
      el.getAttribute("title") ??
      el.getAttribute("data-tip") ??
      el.querySelector("[title],[data-tip]")?.getAttribute("title") ??
      el.querySelector("[title],[data-tip]")?.getAttribute("data-tip") ??
      "",
  );
}

// The chevron's reservation, kept free at the right end of the line.
const RESERVE = 30;

const WIDTHS = [1600, 1500, 1400, 1300, 1200, 1100, 1000, 900, 800, 700, 600];

beforeEach(() => {
  localStorage.clear();
});

describe("the bar across a width sweep", () => {
  for (const tabName of ["home", "insert"]) {
    it(`never runs a ${tabName} control past the edge of the window`, async () => {
      const t = await mountToolbar();
      await selectTab(t.container, tabName);
      for (const width of [...WIDTHS, ...[...WIDTHS].reverse()]) {
        await resizeTo(t.container, width);
        const folded = ribbon(t.container).querySelectorAll("[data-dxw-folded]").length > 0;
        const room = availableWidth(t.container) - (folded ? RESERVE : 0);
        expect(lineWidth(t.container), `${tabName} at ${width}px`).toBeLessThanOrEqual(room);
      }
      await t.unmount();
    });
  }

  it("shows the same bar at a width however the window got there", async () => {
    const t = await mountToolbar();
    await resizeTo(t.container, 1600);
    for (const width of WIDTHS) await resizeTo(t.container, width);
    const shrinking = names(t.container);

    await resizeTo(t.container, 600);
    for (const width of [...WIDTHS].reverse()) await resizeTo(t.container, width);
    await resizeTo(t.container, 600);
    expect(names(t.container)).toEqual(shrinking);
    await t.unmount();
  });

  it("keeps a usable bar at the narrowest window, not an arbitrary one", async () => {
    const t = await mountToolbar();
    await resizeTo(t.container, 600);
    const shown = names(t.container).join("|");
    for (const kept of ["Undo", "Bold", "Italic", "Underline"]) {
      expect(shown, `${kept} at 600px`).toContain(kept);
    }
    // And the way to everything else is offered exactly once.
    expect(t.container.querySelectorAll("[data-dxw-toolbar-expand]").length).toBe(1);
    await t.unmount();
  });

  it("expanded, holds every control and leaves no line hanging", async () => {
    const t = await mountToolbar();
    await selectTab(t.container, "insert");
    await resizeTo(t.container, 900);
    const foldedCount = ribbon(t.container).querySelectorAll("[data-dxw-folded]").length;
    expect(foldedCount, "controls are folded while collapsed").toBeGreaterThan(0);

    await act(async () => {
      t.container
        .querySelector<HTMLButtonElement>("[data-dxw-toolbar-expand]")!
        .dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    });
    expect(ribbon(t.container).querySelectorAll("[data-dxw-folded]").length).toBe(0);
    // The wrapping container is squeezed to the width that spreads the
    // controls evenly over the same number of lines.
    const squeezed = parseFloat(ribbon(t.container).style.maxWidth);
    expect(squeezed).toBeGreaterThan(0);
    expect(squeezed).toBeLessThanOrEqual(availableWidth(t.container) + 400);
    await t.unmount();
  });
});
