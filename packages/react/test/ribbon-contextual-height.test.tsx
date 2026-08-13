/**
 * #155: the bar must not change height when a contextual tab appears.
 *
 * The caret entering a table adds the Table Format tab, the tab strip grows by
 * ~102px, the controls beside it lose 72px of width — and they wrapped onto a
 * second row, taking the bar from 39px to 65px and pushing the whole document
 * down 26px under the pointer. Measured in Chromium on parity-tables.docx:
 *
 *     body caret   bar 39   page-1 top 164.6
 *     table caret  bar 65   page-1 top 190.6
 *
 * The layout engine was supposed to prevent exactly that by folding controls
 * to fit one line, and it folded NOTHING. Both contextual tabs group their
 * controls inside a `display: contents` span, which has no box of its own: the
 * engine measured that one child at 0x0, skipped it as zero-width, and came
 * away with an empty list to fit, while the browser promoted the controls
 * inside it into the line and wrapped them freely.
 *
 * jsdom performs no layout, so the widths below are supplied. That is the
 * point of the test rather than a limitation of it: the defect is entirely
 * about which elements the engine MEASURES, so giving it measurable elements
 * and asking what it folded is the whole question.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { DocxToolbar } from "../src/toolbar.js";
import type { DocxViewApi } from "../src/index.js";

/** Wide enough for a handful of controls, narrow enough to force folding. */
const BAR_WIDTH = 420;
const CONTROL_WIDTH = 90;
const TAB_STRIP_WIDTH = 200;

let root: Root | null = null;
let host: HTMLElement | null = null;
let restore: (() => void) | null = null;

afterEach(async () => {
  if (root) await act(async () => root!.unmount());
  host?.remove();
  restore?.();
  root = null;
  host = null;
  restore = null;
});

/** An engine that reports the caret inside a table, raising the contextual tab. */
function tableCaretApi(): DocxViewApi {
  const impl: Record<string, unknown> = {
    getSelectionFormat: () => ({ bold: false, italic: false, underline: false, strike: false }),
    // undefined means "not in a table"; a string raises the Table Format tab.
    getTableCellFill: () => "FFFFFF",
    getTableBorderStyle: () => "single",
  };
  return new Proxy({} as Record<string, unknown>, {
    get: (_t, key: string) => (key in impl ? impl[key] : () => undefined),
  }) as unknown as DocxViewApi;
}

/**
 * Give the bar a width and every control a width, so the engine has something
 * to fit. A `display: contents` wrapper keeps its real 0x0 box, exactly as a
 * browser reports it — that is the condition under test.
 */
function installWidths(bar: HTMLElement): () => void {
  const priorRect = Element.prototype.getBoundingClientRect;
  const priorClientWidth = Object.getOwnPropertyDescriptor(Element.prototype, "clientWidth");
  const priorOffsetWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, "offsetWidth");

  const widthOf = (el: Element): number => {
    if (!(el instanceof HTMLElement)) return 0;
    if (el.dataset.dxwToolbarMode !== undefined) return BAR_WIDTH;
    if (el.dataset.dxwToolbarTabs !== undefined) return TAB_STRIP_WIDTH;
    // The wrapper itself has no box. This is the whole defect.
    if (el.style.display === "contents") return 0;
    const ribbon = bar.querySelector("[data-dxw-toolbar-ribbon]");
    if (el === ribbon) return BAR_WIDTH - TAB_STRIP_WIDTH;
    // A control is anything mounted directly on a ribbon line.
    if (ribbon?.contains(el) && el.parentElement && ribbon !== el) return CONTROL_WIDTH;
    return 0;
  };

  Element.prototype.getBoundingClientRect = function () {
    const width = widthOf(this);
    return { x: 0, y: 0, left: 0, top: 0, right: width, bottom: 26, width, height: 26,
      toJSON() { return this; } } as DOMRect;
  };
  Object.defineProperty(Element.prototype, "clientWidth", {
    configurable: true, get(this: Element) { return widthOf(this); },
  });
  Object.defineProperty(HTMLElement.prototype, "offsetWidth", {
    configurable: true, get(this: HTMLElement) { return widthOf(this); },
  });
  return () => {
    Element.prototype.getBoundingClientRect = priorRect;
    if (priorClientWidth) Object.defineProperty(Element.prototype, "clientWidth", priorClientWidth);
    if (priorOffsetWidth) Object.defineProperty(HTMLElement.prototype, "offsetWidth", priorOffsetWidth);
  };
}

async function mountBar(api: DocxViewApi): Promise<HTMLElement> {
  host = document.createElement("div");
  document.body.appendChild(host);
  root = createRoot(host);
  await act(async () => { root!.render(createElement(DocxToolbar, { api })); });
  const bar = host.querySelector<HTMLElement>("[data-dxw-toolbar-mode]");
  if (!bar) throw new Error("the toolbar never rendered");
  return bar;
}

describe("#155 · a contextual tab does not change the bar's height", () => {
  it("folds the controls a contextual tab groups behind display:contents", async () => {
    const bar = await mountBar(tableCaretApi());
    restore = installWidths(bar);
    // Re-run the layout engine now that the elements have widths.
    await act(async () => { document.dispatchEvent(new Event("dxw-selection")); });
    await act(async () => { window.dispatchEvent(new Event("resize")); });

    const ribbon = bar.querySelector<HTMLElement>("[data-dxw-toolbar-ribbon]");
    expect(ribbon, "no ribbon rendered").toBeTruthy();
    const wrapper = bar.querySelector<HTMLElement>("[data-dxw-table-format]");
    expect(wrapper, "the Table Format tab is not showing, so nothing is under test").toBeTruthy();
    const controls = [...wrapper!.children] as HTMLElement[];
    expect(controls.length, "the contextual tab rendered no controls").toBeGreaterThan(3);

    // Before the fix the engine measured the wrapper at 0x0, skipped it, and
    // folded nothing at all — so the controls wrapped onto a second row and
    // the bar grew.
    const folded = controls.filter((el) => el.dataset.dxwFolded === "1");
    expect(
      folded.length,
      `nothing folded: ${controls.length} controls of ${CONTROL_WIDTH}px in ${BAR_WIDTH - TAB_STRIP_WIDTH}px of line`,
    ).toBeGreaterThan(0);
  });

  it("keeps every folded control reachable behind the chevron", async () => {
    const bar = await mountBar(tableCaretApi());
    restore = installWidths(bar);
    await act(async () => { document.dispatchEvent(new Event("dxw-selection")); });
    await act(async () => { window.dispatchEvent(new Event("resize")); });

    // Folding a control it cannot then offer is #158's defect, and trading one
    // for the other would be no fix at all.
    const chevron = [...bar.querySelectorAll<HTMLElement>("button")].find((el) =>
      /more tools/i.test(el.getAttribute("title") ?? el.textContent ?? ""),
    );
    const wrapper = bar.querySelector<HTMLElement>("[data-dxw-table-format]");
    const folded = [...(wrapper?.children ?? [])].filter((el) => (el as HTMLElement).dataset.dxwFolded === "1");
    if (folded.length > 0) {
      expect(chevron, "controls were folded away with no chevron to reveal them").toBeTruthy();
    }
  });
});
