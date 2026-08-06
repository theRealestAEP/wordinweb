// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/**
 * The Review tab. Like the other toolbar suites, what is checked is the
 * WIRING: the tab appears in advanced mode, the Track Changes button drives
 * setSuggesting and reflects isSuggesting, and the Accept menu's "all" entry
 * fires acceptAllRevisions. The engine work behind those calls is core's.
 */

function reviewApi() {
  let suggesting = false;
  const spies = {
    setSuggesting: vi.fn((on: boolean) => {
      suggesting = on;
    }),
    acceptRevisionAtCaret: vi.fn(() => true),
    acceptAllRevisions: vi.fn(() => 2),
    rejectAllRevisions: vi.fn(() => 2),
    replaceAll: vi.fn(() => 3),
  };
  const methods = {
    ...spies,
    isSuggesting: () => suggesting,
    revisionCount: () => 2,
    getSelectedObjectContext: () => null,
    getTableCellFill: () => undefined,
    getSelectionFormat: () => ({ bold: false, italic: false, underline: false, strike: false }),
    getParagraphStyleId: () => null,
    getListType: () => null,
    listParagraphStyles: () => [],
    listStyles: () => [],
    imageAccept: () => "image/png,image/jpeg",
  };
  const api = new Proxy(methods, {
    get(target, property) {
      return property in target ? target[property as keyof typeof target] : () => null;
    },
  }) as unknown as DocxViewApi;
  return { api, ...spies };
}

async function mountToolbar(props: { features?: Record<string, boolean>; mode?: "simple" | "advanced" } = {}) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const stub = reviewApi();
  await act(async () => {
    root.render(createElement(DocxToolbar, { api: stub.api, ...props }));
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

function button(container: HTMLElement, tip: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
    (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").includes(tip),
  );
}

async function openReviewTab(t: Awaited<ReturnType<typeof mountToolbar>>) {
  const tab = t.container.querySelector<HTMLButtonElement>('button[data-tab="review"]');
  expect(tab, "review tab").toBeTruthy();
  await click(tab!);
}

describe("the Review tab", () => {
  it("renders in advanced mode with its controls and the revision count", async () => {
    const t = await mountToolbar();
    await openReviewTab(t);
    expect(button(t.container, "Record edits as tracked changes")).toBeTruthy();
    expect(t.container.querySelector('button[aria-label="Accept tracked changes"]')).toBeTruthy();
    expect(t.container.querySelector('button[aria-label="Reject tracked changes"]')).toBeTruthy();
    expect(t.container.querySelector("[data-dxw-revision-count]")?.textContent).toBe("2 changes");
    await t.unmount();
  });

  it("stays out of simple mode and honors the review feature gate", async () => {
    const simple = await mountToolbar({ mode: "simple" });
    expect(simple.container.querySelector('button[data-tab="review"]')).toBeNull();
    await simple.unmount();
    const gated = await mountToolbar({ features: { review: false } });
    expect(gated.container.querySelector('button[data-tab="review"]')).toBeNull();
    await gated.unmount();
  });

  it("Track Changes toggles suggesting and shows the pressed state", async () => {
    const t = await mountToolbar();
    await openReviewTab(t);
    const toggle = button(t.container, "Record edits as tracked changes");
    expect(toggle).toBeTruthy();
    await click(toggle!);
    expect(t.setSuggesting).toHaveBeenLastCalledWith(true);
    // The button re-reads isSuggesting on re-render: tooltip flips and the
    // background carries the active tint (the toolbar's pressed idiom).
    const pressed = button(t.container, "Stop tracking changes");
    expect(pressed).toBeTruthy();
    expect(pressed!.style.background).not.toBe("transparent");
    await click(pressed!);
    expect(t.setSuggesting).toHaveBeenLastCalledWith(false);
    await t.unmount();
  });

  it("Accept all changes fires acceptAllRevisions", async () => {
    const t = await mountToolbar();
    await openReviewTab(t);
    await click(t.container.querySelector<HTMLButtonElement>('button[aria-label="Accept tracked changes"]')!);
    const all = [...t.container.querySelectorAll<HTMLButtonElement>('[role="option"]')]
      .find((o) => o.textContent?.includes("Accept all changes"));
    expect(all, "accept-all option").toBeTruthy();
    await click(all!);
    expect(t.acceptAllRevisions).toHaveBeenCalledTimes(1);
    expect(t.acceptRevisionAtCaret).not.toHaveBeenCalled();
    await t.unmount();
  });

  it("Replace all reports how many replacements applied", async () => {
    const t = await mountToolbar();
    await openReviewTab(t);
    await click(button(t.container, "Find & replace")!);
    const find = t.container.querySelector<HTMLInputElement>('input[aria-label="Find text"]')!;
    const replace = t.container.querySelector<HTMLInputElement>('input[aria-label="Replace with"]')!;
    await act(async () => {
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
      set.call(find, "cat");
      find.dispatchEvent(new Event("input", { bubbles: true }));
      set.call(replace, "dog");
      replace.dispatchEvent(new Event("input", { bubbles: true }));
    });
    const replaceAll = [...t.container.querySelectorAll<HTMLButtonElement>("button")]
      .find((b) => b.textContent === "Replace all");
    await click(replaceAll!);
    expect(t.replaceAll).toHaveBeenCalledWith("cat", "dog");
    expect(t.container.querySelector("[data-dxw-find-status]")?.textContent).toBe("Replaced 3 matches");
    await t.unmount();
  });
});
