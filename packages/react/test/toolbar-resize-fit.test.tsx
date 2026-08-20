// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { availableWidth, controlWidth, firstLineControls, measureLikeABrowser, resizeTo, ribbon, visibleControls } from "./toolbar-measure.js";

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

function chevron(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>("[data-dxw-toolbar-expand]");
}

/** One control, named by its tooltip. */
function nameOf(el: HTMLElement): string {
  return (
    el.getAttribute("title") ??
    el.getAttribute("data-tip") ??
    el.querySelector("[title],[data-tip]")?.getAttribute("title") ??
    el.querySelector("[title],[data-tip]")?.getAttribute("data-tip") ??
    ""
  );
}

/** The controls on the bar, named by their tooltip. */
function names(container: HTMLElement): string[] {
  return visibleControls(container).map(nameOf);
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
        // The seat at the end of the line is kept for a chevron that is going
        // to be there. When the bar decides not to offer one, the line gets
        // that width back, and the controls are allowed to use it.
        const room = availableWidth(t.container) - (chevron(t.container) ? RESERVE : 0);
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

  it("keeps the tab strip's line in use rather than emptying it", async () => {
    // What the user saw and reported: expanding moved every control onto a
    // line of its own, leaving the tab strip alone on the first line with a
    // window's width of empty bar beside it. The controls only leave that
    // line when going full-width saves the bar a whole line.
    const t = await mountToolbar();
    await selectTab(t.container, "home");
    for (const width of [1500, 1400, 1300, 1200]) {
      await resizeTo(t.container, width);
      const toggle = chevron(t.container);
      if (!toggle) continue;
      await act(async () => {
        toggle.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
      });
      expect(visibleControls(t.container).length, `controls beside the tabs at ${width}px`).toBeGreaterThan(0);
      expect(ribbon(t.container).parentElement!.style.flex, `${width}px stays beside the tabs`).toBe("1 1 0%");
      await act(async () => {
        chevron(t.container)!.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
      });
    }
    await t.unmount();
  });

  it("expanded, holds every control and moves none of the ones already on the bar", async () => {
    const t = await mountToolbar();
    await selectTab(t.container, "insert");
    await resizeTo(t.container, 900);
    const foldedCount = ribbon(t.container).querySelectorAll("[data-dxw-folded]").length;
    expect(foldedCount, "controls are folded while collapsed").toBeGreaterThan(0);
    const before = names(t.container);

    await act(async () => {
      t.container
        .querySelector<HTMLButtonElement>("[data-dxw-toolbar-expand]")!
        .dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    });
    // Every control is reachable...
    expect(ribbon(t.container).querySelectorAll("[data-dxw-folded]").length).toBe(0);
    expect(visibleControls(t.container).length).toBeGreaterThan(before.length);
    // ...and the line the user was already pointing at did not move. This is
    // the reported bug: opening the bar used to re-wrap the whole tab, which
    // pushed controls that were already visible down a row.
    expect(firstLineControls(t.container).map(nameOf)).toEqual(before);
    // Held below by a forced break, not by squeezing the container — the
    // squeeze is what moved the first line.
    expect(ribbon(t.container).querySelector("[data-dxw-ribbon-break]")).toBeTruthy();
    expect(ribbon(t.container).style.maxWidth).toBe("");
    await t.unmount();
  });
});
