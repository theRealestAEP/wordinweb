// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { measureLikeABrowser, resizeTo, ribbon, visibleControls } from "./toolbar-measure.js";

/**
 * The expand chevron — the toolbar's one and only "there is more" affordance
 * since the ⋮ overflow menu was removed. What is checked: it appears only
 * when controls are actually folded away, expanding puts every control of the
 * tab straight into the bar (no popover, nothing behind a menu), and the
 * choice survives a remount.
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
});
