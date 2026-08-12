import { act } from "react";
import { expect } from "vitest";

/**
 * jsdom has no layout engine: every element measures zero, and a toolbar that
 * measures zero correctly decides it cannot fold anything. These helpers give
 * the DOM plausible control widths so the bar's layout engine has something
 * real to fit, which is the only way to test folding without a browser.
 *
 * The widths are the shapes the bar actually renders — a divider is a hairline,
 * an icon button is square, a menu select is as wide as its `width` prop, and a
 * text button grows with its label.
 */
export function measureLikeABrowser(): void {
  Object.defineProperty(HTMLElement.prototype, "offsetWidth", {
    configurable: true,
    get(this: HTMLElement): number {
      if (this.style.display === "none") return 0;
      if (this.dataset.dxwSep !== undefined) return 1;
      if (this.hasAttribute("data-dxw-menu-select")) return parseFloat(this.style.width) || 90;
      if (this.tagName === "BUTTON") return Math.max(30, 12 + (this.textContent ?? "").trim().length * 7);
      if (this.tagName === "INPUT") return 0;
      let total = 0;
      for (const child of Array.from(this.children) as HTMLElement[]) {
        const placement = getComputedStyle(child).position;
        if (placement === "fixed" || placement === "absolute") continue;
        total += child.offsetWidth;
      }
      return total;
    },
  });
  stubRects();
}

/**
 * jsdom's getBoundingClientRect is all zeros, and the bar measures with it
 * (fractions and margins matter when thirty controls share a line). Mirror the
 * stubbed offsetWidth into it so the engine sees the same fake layout.
 */
function stubRects(): void {
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement): DOMRect {
    const width = this.offsetWidth;
    return { x: 0, y: 0, top: 0, left: 0, right: width, bottom: 26, width, height: 26, toJSON: () => ({}) } as DOMRect;
  };
}

/** The room a control takes on the line: its box plus its margins, the way
 * the bar's engine counts it. */
export function controlWidth(el: HTMLElement): number {
  const style = getComputedStyle(el);
  return Math.ceil(
    el.getBoundingClientRect().width + (parseFloat(style.marginLeft) || 0) + (parseFloat(style.marginRight) || 0),
  );
}

export function toolbarRoot(container: HTMLElement): HTMLElement {
  const root = container.querySelector<HTMLElement>("[data-dxw-toolbar-mode]");
  expect(root, "toolbar root").toBeTruthy();
  return root!;
}

export function ribbon(container: HTMLElement): HTMLElement {
  const el = container.querySelector<HTMLElement>("[data-dxw-toolbar-ribbon]");
  expect(el, "toolbar ribbon").toBeTruthy();
  return el!;
}

/** The controls the bar is showing right now, in bar order. */
export function visibleControls(container: HTMLElement): HTMLElement[] {
  return (Array.from(ribbon(container).children) as HTMLElement[]).filter(
    (el) => el.style.display !== "none" && el.offsetWidth > 0,
  );
}

/** Resize the window the way a user does, and let the bar re-fit. */
export async function resizeTo(container: HTMLElement, width: number): Promise<void> {
  const root = toolbarRoot(container);
  Object.defineProperty(root, "clientWidth", { value: width, configurable: true });
  await act(async () => {
    window.dispatchEvent(new Event("resize"));
  });
}

/** The space the engine has for controls: the bar, less padding, the tab
 * strip and the trailing controls — computed here independently of the
 * engine, so a test can catch the engine disagreeing with the browser. */
export function availableWidth(container: HTMLElement): number {
  const root = toolbarRoot(container);
  const style = getComputedStyle(root);
  const padding = (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0);
  const tabs = root.querySelector<HTMLElement>("[data-dxw-toolbar-tabs]");
  const trailing = root.querySelector<HTMLElement>("[data-dxw-toolbar-trailing]");
  return root.clientWidth - padding - (tabs?.offsetWidth ?? 0) - (trailing?.offsetWidth ?? 0) - 4;
}
