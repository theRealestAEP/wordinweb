// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { measureLikeABrowser, resizeTo, ribbon, toolbarRoot } from "./toolbar-measure.js";

/**
 * A control's menu must open over the document, never inside the bar.
 *
 * The bar this replaced folded its extra controls into a ⋮ popover, and that
 * popover both animated with a `transform` and scrolled with `overflow: auto`.
 * A transform makes an element the containing block for `position: fixed`
 * descendants, so the style dropdown opened by a control inside the popover
 * stopped being positioned against the window and was laid over the popover's
 * own rows, then clipped by its scroll box. The controls now live on the bar
 * itself, and these are the two structural properties that keep any menu they
 * open free of it.
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
    listParagraphStyles: () => [
      { id: "Heading1", name: "Heading 1" },
      { id: "Heading2", name: "Heading 2" },
    ],
    listStyles: () => [],
    imageAccept: () => "image/png,image/jpeg",
  };
  return new Proxy(methods, {
    get(target, property) {
      return property in target ? target[property as keyof typeof target] : () => null;
    },
  }) as unknown as DocxViewApi;
}

async function mountToolbar(width: number) {
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

/** Every element between `from` and the document body. */
function ancestors(from: HTMLElement, until: HTMLElement): HTMLElement[] {
  const chain: HTMLElement[] = [];
  let el = from.parentElement;
  while (el) {
    chain.push(el);
    if (el === until) break;
    el = el.parentElement;
  }
  return chain;
}

beforeEach(() => {
  localStorage.clear();
});

describe("menus opened from the bar", () => {
  it("are positioned against the window, with nothing in the bar able to trap them", async () => {
    const t = await mountToolbar(800);
    // Expanded: the state where the bar is at its tallest and a menu opened
    // from a lower line has the most bar above and below it.
    await click(t.container.querySelector<HTMLButtonElement>("[data-dxw-toolbar-expand]")!);

    const trigger = [...t.container.querySelectorAll<HTMLButtonElement>("[data-dxw-menu-select-trigger]")]
      .find((el) => (el.getAttribute("data-tip") ?? "") === "Paragraph style");
    expect(trigger, "paragraph style trigger").toBeTruthy();
    await click(trigger!);

    const menu = t.container.querySelector<HTMLElement>("[data-dxw-menu-select-menu]");
    expect(menu, "the style menu is open").toBeTruthy();
    expect(menu!.style.position, "positioned against the window").toBe("fixed");

    for (const parent of ancestors(menu!, toolbarRoot(t.container))) {
      const style = getComputedStyle(parent);
      expect(["", "none"], `${describeEl(parent)} must not become a containing block`)
        .toContain(style.transform);
      expect(["", "visible"], `${describeEl(parent)} must not clip the menu`)
        .toContain(style.overflow || "");
    }
    await t.unmount();
  });

  it("folds controls away rather than clipping them", async () => {
    // The bar hides what does not fit; it never keeps a control on the line
    // and cuts it off with a scroll box.
    const t = await mountToolbar(700);
    const line = ribbon(t.container);
    expect(getComputedStyle(line).overflow || "").toMatch(/^(|visible)$/);
    expect(line.querySelectorAll("[data-dxw-folded]").length).toBeGreaterThan(0);
    await t.unmount();
  });
});

function describeEl(el: HTMLElement): string {
  return el.tagName.toLowerCase() + (el.dataset.dxwToolbarRibbon !== undefined ? "[ribbon]" : "");
}
