// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";
import type { DocxViewApi, StyleGalleryEntry } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/**
 * The Home tab's two style controls: the character-style menu and the format
 * painter. Both are thin — the engine work is in core — so what is checked is
 * the WIRING: which api call each fires, with what argument, and the painter's
 * two-click contract (copy, then paint, then hand the brush back).
 */

const CHARACTER_STYLES: StyleGalleryEntry[] = [
  {
    id: "Emphasis",
    name: "Emphasis",
    type: "character",
    quickStyle: true,
    defined: true,
    usageCount: 0,
    preview: { italic: true },
  },
  {
    id: "BookTitle",
    name: "Book Title",
    type: "character",
    quickStyle: false,
    // Not a quick style, but the document uses it, so the menu offers it.
    defined: true,
    usageCount: 3,
    preview: { bold: true },
  },
  {
    id: "Hidden",
    name: "Hidden",
    type: "character",
    quickStyle: false,
    defined: true,
    usageCount: 0,
    preview: {},
  },
];

function styleApi(options: { styles?: StyleGalleryEntry[]; characterStyleId?: string | null } = {}) {
  const spies = {
    applyFormat: vi.fn(),
    applyCopiedFormatting: vi.fn(),
    copyFormatting: vi.fn(() => ({ bold: true, italic: false, underline: false, strike: false })),
  };
  const methods = {
    ...spies,
    getSelectedObjectContext: () => null,
    getSelectedChart: () => null,
    getTableCellFill: () => undefined,
    getSelectionFormat: () => ({
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      characterStyleId: options.characterStyleId ?? null,
    }),
    getParagraphStyleId: () => null,
    getListType: () => null,
    listParagraphStyles: () => [],
    listStyles: () => options.styles ?? CHARACTER_STYLES,
    imageAccept: () => "image/png,image/jpeg",
  };
  const api = new Proxy(methods, {
    get(target, property) {
      return property in target ? target[property as keyof typeof target] : () => null;
    },
  }) as unknown as DocxViewApi;
  return { api, ...spies };
}

async function mountToolbar(options: Parameters<typeof styleApi>[0] = {}) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const stub = styleApi(options);
  await act(async () => {
    root.render(createElement(DocxToolbar, { api: stub.api }));
  });
  await act(async () => {
    document.dispatchEvent(new Event("dxw-selection"));
    await new Promise<void>((r) => setTimeout(r, 5));
  });
  return {
    container,
    ...stub,
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

function optionText(item: Element): string {
  return (item.textContent ?? "").replace(/[✓ ]/g, "").trim();
}

/** The Home tab's menus name themselves with `title` (the toolbar swaps it for
 * data-tip on hover), so both spellings are accepted. */
function menuTrigger(container: HTMLElement, title: string): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>(
    `button[title="${title}"], button[data-tip="${title}"]`,
  );
}

async function openMenu(container: HTMLElement, title: string) {
  const trigger = menuTrigger(container, title);
  expect(trigger, `menu "${title}"`).toBeTruthy();
  await click(trigger!);
  return [...container.querySelectorAll<HTMLButtonElement>('[role="option"]')];
}

const CHAR_STYLE = "Character style";

describe("Home tab: the character-style menu", () => {
  it("offers quick styles and styles the document already uses, and nothing else", async () => {
    const t = await mountToolbar();
    const options = (await openMenu(t.container, CHAR_STYLE)).map(optionText);
    expect(options).toEqual(["None", "Emphasis", "Book Title"]);
    await t.unmount();
  });

  it("applies the chosen style as a run format patch", async () => {
    const t = await mountToolbar();
    const options = await openMenu(t.container, CHAR_STYLE);
    await click(options.find((o) => optionText(o) === "Emphasis")!);
    expect(t.applyFormat).toHaveBeenLastCalledWith({ characterStyleId: "Emphasis" });
    await t.unmount();
  });

  it("removes the style with null rather than clearing the run", async () => {
    const t = await mountToolbar({ characterStyleId: "Emphasis" });
    const options = await openMenu(t.container, CHAR_STYLE);
    await click(options.find((o) => optionText(o) === "None")!);
    expect(t.applyFormat).toHaveBeenLastCalledWith({ characterStyleId: null });
    await t.unmount();
  });

  it("stays out of the toolbar when the document offers no character style", async () => {
    const t = await mountToolbar({ styles: [] });
    expect(menuTrigger(t.container, CHAR_STYLE)).toBeNull();
    await t.unmount();
  });
});

describe("Home tab: the format painter", () => {
  function painterButton(container: HTMLElement): HTMLButtonElement {
    const button = [...container.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
      (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").includes("format painter"),
    );
    expect(button, "format painter button").toBeTruthy();
    return button!;
  }

  it("copies on the first click and paints on the second", async () => {
    const t = await mountToolbar();
    await click(painterButton(t.container));
    expect(t.copyFormatting).toHaveBeenCalledTimes(1);
    expect(t.applyCopiedFormatting).not.toHaveBeenCalled();

    // The button's tooltip flips once it is loaded, which is how the user
    // knows the brush is armed.
    const armed = [...t.container.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
      (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").includes("Paint the copied"),
    );
    expect(armed, "armed painter button").toBeTruthy();

    await click(armed!);
    expect(t.applyCopiedFormatting).toHaveBeenCalledWith({
      bold: true,
      italic: false,
      underline: false,
      strike: false,
    });
    await t.unmount();
  });

  it("is single-use: the third click copies again rather than painting twice", async () => {
    const t = await mountToolbar();
    const armed = () =>
      [...t.container.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
        (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").includes("Paint the copied"),
      );
    await click(painterButton(t.container));
    await click(armed()!);
    expect(armed()).toBeUndefined();
    await click(painterButton(t.container));
    expect(t.copyFormatting).toHaveBeenCalledTimes(2);
    expect(t.applyCopiedFormatting).toHaveBeenCalledTimes(1);
    await t.unmount();
  });
});
