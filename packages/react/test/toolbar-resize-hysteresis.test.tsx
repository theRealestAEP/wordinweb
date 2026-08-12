// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { beforeEach, describe, expect, it } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/**
 * Hysteresis on the tier-1/tier-2 boundary (720px fold / 760px unfold — see
 * TIER_FOLD_WIDTH/TIER_UNFOLD_WIDTH in toolbar.tsx). Without the 40px gap
 * between the fold and unfold widths, a toolbar clientWidth that lands in
 * that band flaps between tiers on every ResizeObserver tick, which reads to
 * a user as controls jittering in and out of the ⋮ menu while the window
 * settles at a particular width.
 *
 * jsdom has no real layout, so a stub ResizeObserver is used and widths are
 * driven directly by stubbing `clientWidth` on the toolbar root, then firing
 * the same `resize` listener the component itself registers as a fallback
 * (`window.addEventListener("resize", measure)`).
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

function button(container: HTMLElement, tip: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
    (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").includes(tip),
  );
}

/** Drives the toolbar's measure() the same way a real resize would: update
 * the root's clientWidth, then fire the `resize` event it listens for. */
async function resizeTo(container: HTMLElement, width: number) {
  const el = container.querySelector<HTMLElement>("[data-dxw-toolbar-mode]");
  expect(el, "toolbar root").toBeTruthy();
  Object.defineProperty(el, "clientWidth", { value: width, configurable: true });
  await act(async () => {
    window.dispatchEvent(new Event("resize"));
  });
}

// "Align left" only folds into ⋮ at tier 2 (the alignment group isn't in the
// tier-1 overflow set), so it's a clean marker for the 720/760 boundary.
const ALIGNMENT_CONTROL = "Align left";

beforeEach(() => {
  localStorage.clear();
});

describe("tier hysteresis at the 720/760 boundary", () => {
  it("does not flap while the width jitters inside the fold/unfold gap", async () => {
    const t = await mountToolbar();

    // Establish tier 1 (unfolded alignment) from a mid-range width.
    await resizeTo(t.container, 800);
    expect(button(t.container, ALIGNMENT_CONTROL), "inline at 800px").toBeTruthy();

    // Cross below the fold width: tier 2, alignment now behind ⋮.
    await resizeTo(t.container, 700);
    expect(button(t.container, ALIGNMENT_CONTROL), "folded at 700px").toBeUndefined();

    // Jitter through the 720-760 hysteresis band without ever reaching the
    // 760px unfold width. A naive single-breakpoint (>=720) implementation
    // would unfold on every one of these steps since they're all >=720; the
    // hysteresis gap must keep it folded throughout.
    for (const w of [740, 750, 710, 745, 725]) {
      await resizeTo(t.container, w);
      expect(button(t.container, ALIGNMENT_CONTROL), `still folded at ${w}px`).toBeUndefined();
    }

    // Clear the unfold width: back to tier 1.
    await resizeTo(t.container, 765);
    expect(button(t.container, ALIGNMENT_CONTROL), "unfolded at 765px").toBeTruthy();

    // Jitter back down through the same band from the tier-1 side without
    // reaching the 720px fold width — must stay unfolded throughout.
    for (const w of [730, 745, 722, 750]) {
      await resizeTo(t.container, w);
      expect(button(t.container, ALIGNMENT_CONTROL), `still inline at ${w}px`).toBeTruthy();
    }

    // Cross the fold width again: back to tier 2.
    await resizeTo(t.container, 715);
    expect(button(t.container, ALIGNMENT_CONTROL), "folded again at 715px").toBeUndefined();

    await t.unmount();
  });
});
