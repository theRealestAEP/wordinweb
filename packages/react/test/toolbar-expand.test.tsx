// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { MIN_REVEAL } from "../src/ribbon-layout.js";
import { DocxToolbar } from "../src/toolbar.js";
import { measureLikeABrowser, resizeTo, ribbon, visibleControls } from "./toolbar-measure.js";

/**
 * The expand chevron — the toolbar's one and only "there is more" affordance
 * since the ⋮ overflow menu was removed. What is checked: it appears only
 * when there is a line's worth of tools behind it, expanding puts every
 * control of the tab straight into the bar (no popover, nothing behind a
 * menu), and the choice survives a remount.
 *
 * jsdom has no layout, so control widths are stubbed (see toolbar-measure).
 * Without that the bar measures a width of zero, correctly refuses to fold
 * anything, and there is nothing here to test.
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

async function mountToolbar(width = 900) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(DocxToolbar, { api: stubApi() }));
  });
  await resizeTo(container, width);
  return {
    container,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
}

async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

function control(container: HTMLElement, tip: string): HTMLElement | undefined {
  return visibleControls(container).find((el) =>
    (el.getAttribute("title") ?? el.getAttribute("data-tip") ?? "").includes(tip),
  );
}

function chevron(container: HTMLElement): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>("[data-dxw-toolbar-expand]");
}

/** Folded controls a user would call tools: the dividers between them are
 * folded too, and counting those is how "Show 4 more tools" once meant two. */
function foldedControls(container: HTMLElement): number {
  return Array.from(ribbon(container).querySelectorAll<HTMLElement>("[data-dxw-folded]")).filter(
    (el) => el.dataset.dxwSep === undefined,
  ).length;
}

async function openInsertTab(container: HTMLElement) {
  const tab = container.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
  expect(tab, "insert tab").toBeTruthy();
  await click(tab!);
}

// A control at the far end of the Insert tab: the first thing to fold away.
const DEEP_CONTROL = "Embed a file in this document";

beforeEach(() => {
  localStorage.clear();
});

describe("the toolbar expand chevron", () => {
  it("expanding brings every folded control onto the bar", async () => {
    const t = await mountToolbar(900);
    await openInsertTab(t.container);
    expect(control(t.container, DEEP_CONTROL), "folded while collapsed").toBeUndefined();

    const toggle = chevron(t.container);
    expect(toggle, "chevron offered").toBeTruthy();
    expect(toggle!.getAttribute("aria-expanded")).toBe("false");
    expect(toggle!.getAttribute("aria-label")).toMatch(/^Show \d+ more tools?$/);
    await click(toggle!);

    expect(control(t.container, DEEP_CONTROL), "on the bar once expanded").toBeTruthy();
    expect(
      ribbon(t.container).querySelectorAll("[data-dxw-folded]").length,
      "nothing left folded",
    ).toBe(0);
    const pressed = chevron(t.container)!;
    expect(pressed.getAttribute("aria-expanded")).toBe("true");
    expect(pressed.getAttribute("aria-label")).toBe("Hide the extra tools");
    await t.unmount();
  });

  it("persists the choice in localStorage and restores it on remount", async () => {
    const t = await mountToolbar(900);
    await openInsertTab(t.container);
    await click(chevron(t.container)!);
    expect(localStorage.getItem("dxw-toolbar-expanded")).toBe("1");
    await t.unmount();

    const again = await mountToolbar(900);
    await openInsertTab(again.container);
    expect(chevron(again.container)!.getAttribute("aria-expanded")).toBe("true");
    expect(control(again.container, DEEP_CONTROL)).toBeTruthy();
    await again.unmount();
  });

  it("collapsing folds the extra controls away again", async () => {
    localStorage.setItem("dxw-toolbar-expanded", "1");
    const t = await mountToolbar(900);
    await openInsertTab(t.container);
    expect(control(t.container, DEEP_CONTROL)).toBeTruthy();

    await click(chevron(t.container)!);
    expect(localStorage.getItem("dxw-toolbar-expanded")).toBe("0");
    expect(control(t.container, DEEP_CONTROL)).toBeUndefined();
    await t.unmount();
  });

  it("offers no chevron at a width where every control fits", async () => {
    // Nothing is folded, so there is nothing to reveal and no chrome
    // pretending otherwise.
    const t = await mountToolbar(3000);
    expect(chevron(t.container), "no chevron on a wide bar").toBeNull();
    expect(ribbon(t.container).querySelectorAll("[data-dxw-folded]").length).toBe(0);
    await t.unmount();
  });

  it("stays silent only while it is hiding nothing", async () => {
    // Two complaints, and one rule that answers both.
    //
    // The first was a chevron that cost a whole line of the window to put two
    // icons back — "it does nothing except add a random weird space". The
    // answer is the width reserved for the chevron: give it back to the
    // controls and, just below the width where everything fits, everything
    // still fits. That band is real and the sweep crosses it.
    //
    // The second (#158) was the overcorrection. Below that band the bar hid
    // the overflow and stayed silent anyway, so at 900px Hyphenation and
    // Compare Documents were gone with no chevron, no menu and no shortcut.
    // So silence is allowed only while nothing is hidden: once a control
    // folds, the chevron is the only route back to it and must be offered,
    // however few are behind it.
    const t = await mountToolbar();
    let sawSilentAndFull = false;
    let sawChevronBelowThreshold = false;
    for (let width = 2000; width >= 700; width -= 20) {
      await resizeTo(t.container, width);
      const waiting = foldedControls(t.container);
      if (chevron(t.container)) {
        expect(waiting, `chevron at ${width}px promises nothing`).toBeGreaterThan(0);
        if (waiting < MIN_REVEAL) sawChevronBelowThreshold = true;
      } else {
        expect(waiting, `silent bar at ${width}px has hidden ${waiting} control(s)`).toBe(0);
        sawSilentAndFull = true;
      }
    }
    expect(sawSilentAndFull, "the sweep never saw the bar silent").toBe(true);
    // Without this, the rule above would also be satisfied by a bar that never
    // folds fewer than MIN_REVEAL at any width — so the sweep would pass
    // without ever entering the band #158 lived in.
    expect(sawChevronBelowThreshold, "the sweep never crossed the band #158 lived in").toBe(true);
    await t.unmount();
  });
});
