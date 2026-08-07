// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/**
 * The expand/collapse chevron. jsdom has no layout, so a stub ResizeObserver
 * lets the toolbar's measure() run: clientWidth is 0 there, which lands on
 * the narrowest tier and folds the low-frequency groups into the ⋮ overflow
 * menu — exactly the situation the chevron exists to undo. What is checked:
 * expanding puts the folded Insert controls straight into the DOM (no popover
 * needed), the choice persists in localStorage across a remount, and
 * collapsing restores the folding.
 */

class StubResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = StubResizeObserver;

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

async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

function button(container: HTMLElement, tip: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
    (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").includes(tip),
  );
}

async function openInsertTab(container: HTMLElement) {
  const tab = container.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
  expect(tab, "insert tab").toBeTruthy();
  await click(tab!);
}

// A control the narrow tiers fold into ⋮ on the Insert tab.
const DEEP_CONTROL = "Insert blank page";

beforeEach(() => {
  localStorage.clear();
});

describe("the toolbar expand chevron", () => {
  it("expanding reveals folded Insert controls inline, without any popover", async () => {
    const t = await mountToolbar();
    await openInsertTab(t.container);
    // Collapsed at the narrow tier: the deep control lives behind ⋮ and is
    // not in the DOM until that menu opens.
    expect(button(t.container, DEEP_CONTROL)).toBeUndefined();
    expect(t.container.querySelector("[data-dxw-overflow]")).toBeTruthy();

    const toggle = button(t.container, "Expand the toolbar");
    expect(toggle, "expand toggle").toBeTruthy();
    expect(toggle!.getAttribute("aria-expanded")).toBe("false");
    await click(toggle!);

    // Expanded: the control is directly in the DOM, nothing is folded, and
    // the ⋮ button is gone. The toggle reads pressed.
    expect(button(t.container, DEEP_CONTROL)).toBeTruthy();
    expect(t.container.querySelector("[data-dxw-overflow]")).toBeNull();
    const pressed = button(t.container, "Collapse the toolbar");
    expect(pressed).toBeTruthy();
    expect(pressed!.getAttribute("aria-expanded")).toBe("true");
    expect(pressed!.style.background).not.toBe("transparent");
    await t.unmount();
  });

  it("persists the choice in localStorage and restores it on remount", async () => {
    const t = await mountToolbar();
    await click(button(t.container, "Expand the toolbar")!);
    expect(localStorage.getItem("dxw-toolbar-expanded")).toBe("1");
    await t.unmount();

    const again = await mountToolbar();
    await openInsertTab(again.container);
    expect(button(again.container, "Collapse the toolbar")).toBeTruthy();
    expect(button(again.container, DEEP_CONTROL)).toBeTruthy();
    await again.unmount();
  });

  it("collapsing restores the folding and records the choice", async () => {
    localStorage.setItem("dxw-toolbar-expanded", "1");
    const t = await mountToolbar();
    await openInsertTab(t.container);
    expect(button(t.container, DEEP_CONTROL)).toBeTruthy();

    await click(button(t.container, "Collapse the toolbar")!);
    expect(localStorage.getItem("dxw-toolbar-expanded")).toBe("0");
    expect(button(t.container, DEEP_CONTROL)).toBeUndefined();
    expect(t.container.querySelector("[data-dxw-overflow]")).toBeTruthy();
    await t.unmount();
  });
});
