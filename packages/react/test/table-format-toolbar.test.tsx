// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";
import type { DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";

/**
 * The Table Format tab's new controls: borders, table style, the six style
 * options, autofit and the header band.
 *
 * The tab appears whenever the caret is inside a table, which the api reports
 * by returning something other than `undefined` from getTableCellFill. Each
 * control is checked for the ARGUMENTS it sends, not merely that it fired —
 * "No borders" and "Clear direct borders" write different files, and a
 * control that confused them would look fine from a call count.
 */

type Look = NonNullable<ReturnType<DocxViewApi["getTableLook"]>>;

const DEFAULT_LOOK: Look = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  bandedRows: true,
  bandedCols: false,
};

function tableApi(options: { styleId?: string | null; look?: Look } = {}) {
  const spies = {
    setTableBorders: vi.fn(),
    setTableStyle: vi.fn(),
    setTableLook: vi.fn(),
    setTableWidth: vi.fn(),
    setTableLayout: vi.fn(),
    setTableHeaderRows: vi.fn(),
    insertTableFormula: vi.fn(() => true),
  };
  const methods = {
    ...spies,
    getSelectedObjectContext: () => null,
    getSelectedChart: () => null,
    // Not undefined: the caret is inside a table, with no direct cell fill.
    getTableCellFill: () => null,
    getTableStyleId: () => options.styleId ?? null,
    getTableLook: () => options.look ?? DEFAULT_LOOK,
    listTableStyles: () => [{ id: "GridBlue", name: "Grid Blue" }],
    getSelectionFormat: () => null,
    getParagraphStyleId: () => null,
    getListType: () => null,
    listParagraphStyles: () => [],
    imageAccept: () => "image/png,image/jpeg",
  };
  const api = new Proxy(methods, {
    get(target, property) {
      return property in target ? target[property as keyof typeof target] : () => null;
    },
  }) as unknown as DocxViewApi;
  return { api, ...spies };
}

async function mountToolbar(options: { styleId?: string | null; look?: Look } = {}) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const stub = tableApi(options);
  await act(async () => {
    root.render(createElement(DocxToolbar, { api: stub.api }));
  });
  // The toolbar recomputes its context on a selection event, and the caret
  // landing in a table is what raises the Table Format tab.
  await act(async () => {
    document.dispatchEvent(new Event("dxw-selection"));
    await new Promise<void>((r) => setTimeout(r, 5));
  });
  const tab = container.querySelector<HTMLButtonElement>('button[data-tab="tableFormat"]');
  expect(tab, "Table Format tab").toBeTruthy();
  await click(tab!);
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

/** An option's text without the tick marks the menus draw for state. */
function optionText(item: Element): string {
  return (item.textContent ?? "").replace(/[\u2713\u2007]/g, "").trim();
}

/** Open the menu whose trigger carries `title`, then pick the option named
 * `label`. */
async function pick(container: HTMLElement, title: string, label: string) {
  const trigger = container.querySelector<HTMLButtonElement>(`button[aria-label="${title}"]`);
  expect(trigger, `menu "${title}"`).toBeTruthy();
  await click(trigger!);
  const options = [...container.querySelectorAll<HTMLButtonElement>('[role="option"]')];
  const option = options.find((item) => optionText(item) === label);
  expect(option, `option "${label}" among ${options.map(optionText).join(", ")}`).toBeTruthy();
  await click(option!);
}

const BORDERS = "Set or clear the borders of the table or the current cell";
const STYLE = "Apply a table style defined in this document";
const OPTIONS = "Choose which parts of the table style apply";
const AUTOFIT = "Choose how the table sizes its columns";
const HEADERS = "Repeat leading rows at the top of every page";

describe("Table Format tab: borders, style, autofit, header rows", () => {
  it("keeps 'No borders' and 'Clear direct borders' apart", async () => {
    const t = await mountToolbar();
    await pick(t.container, BORDERS, "No borders");
    // Suppression: an explicit style, so the table style's rules stay hidden.
    expect(t.setTableBorders).toHaveBeenLastCalledWith(
      "table",
      ["top", "bottom", "left", "right", "insideH", "insideV"],
      { style: "none" },
    );
    await pick(t.container, BORDERS, "Clear direct borders");
    // Removal: null, so the style's rules come back.
    expect(t.setTableBorders).toHaveBeenLastCalledWith(
      "table",
      ["top", "bottom", "left", "right", "insideH", "insideV"],
      null,
    );
    await t.unmount();
  });

  it("never asks for an edge the engine refuses at that scope", async () => {
    const t = await mountToolbar();
    await pick(t.container, BORDERS, "Inside borders");
    expect(t.setTableBorders).toHaveBeenLastCalledWith("table", ["insideH", "insideV"], { style: "single", sz: 4 });

    // Every entry in the menu, driven. Cell scope refuses insideH/insideV and
    // table scope refuses the diagonals, so a control that sent one would be
    // dead on arrival.
    const labels = ["All borders", "Outside borders", "Inside borders", "No borders", "Clear direct borders"];
    for (const label of labels) await pick(t.container, BORDERS, label);
    for (const [scope, edges] of t.setTableBorders.mock.calls as [string, string[]][]) {
      const allowed =
        scope === "cell"
          ? ["top", "bottom", "left", "right", "tl2br", "tr2bl"]
          : ["top", "bottom", "left", "right", "insideH", "insideV"];
      expect(edges.length, `${scope} edges`).toBeGreaterThan(0);
      for (const edge of edges) expect(allowed, `${edge} at ${scope} scope`).toContain(edge);
    }
    await t.unmount();
  });

  it("applies and clears a table style from the document's own list", async () => {
    const t = await mountToolbar();
    await pick(t.container, STYLE, "Grid Blue");
    expect(t.setTableStyle).toHaveBeenLastCalledWith("GridBlue");
    await t.unmount();

    const applied = await mountToolbar({ styleId: "GridBlue" });
    await pick(applied.container, STYLE, "No style");
    expect(applied.setTableStyle).toHaveBeenLastCalledWith(null);
    await applied.unmount();
  });

  it("offers the style options only once a style is applied, and toggles them", async () => {
    const none = await mountToolbar({ styleId: null });
    expect(none.container.querySelector(`button[aria-label="${OPTIONS}"]`)).toBeNull();
    await none.unmount();

    const t = await mountToolbar({ styleId: "GridBlue" });
    // Banded rows are on, so picking them turns them OFF.
    await pick(t.container, OPTIONS, "Banded rows");
    expect(t.setTableLook).toHaveBeenLastCalledWith({ bandedRows: false });
    // Total row is off, so picking it turns it ON.
    await pick(t.container, OPTIONS, "Total row");
    expect(t.setTableLook).toHaveBeenLastCalledWith({ lastRow: true });
    await t.unmount();
  });

  it("distinguishes autofit to contents from autofit to window", async () => {
    const t = await mountToolbar();
    await pick(t.container, AUTOFIT, "AutoFit to contents");
    expect(t.setTableLayout).toHaveBeenLastCalledWith("autofit");
    expect(t.setTableWidth).not.toHaveBeenCalled();

    await pick(t.container, AUTOFIT, "AutoFit to window");
    expect(t.setTableLayout).toHaveBeenLastCalledWith("autofit");
    // To window is autofit measured against a full-width target.
    expect(t.setTableWidth).toHaveBeenLastCalledWith("pct", 100);

    await pick(t.container, AUTOFIT, "Fixed column width");
    expect(t.setTableLayout).toHaveBeenLastCalledWith("fixed");
    expect(t.setTableWidth).toHaveBeenCalledTimes(1); // still only the "window" call
    await t.unmount();
  });

  it("sets a header BAND, not only a single row", async () => {
    const t = await mountToolbar();
    await pick(t.container, HEADERS, "First two rows");
    expect(t.setTableHeaderRows).toHaveBeenLastCalledWith(2);
    await pick(t.container, HEADERS, "None");
    expect(t.setTableHeaderRows).toHaveBeenLastCalledWith(0);
    await t.unmount();
  });

  it("inserts a formula through the Formula dialog, validating the grammar live", async () => {
    const t = await mountToolbar();
    const trigger = t.container.querySelector<HTMLButtonElement>(
      'button[title=\'Insert a formula field ("=SUM(ABOVE)") in the current cell\']',
    );
    expect(trigger, "Formula button").toBeTruthy();
    await click(trigger!);
    const dialog = t.container.querySelector<HTMLElement>('[role="dialog"][aria-label="Formula"]');
    expect(dialog, "Formula dialog").toBeTruthy();
    const apply = dialog!.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]")!;
    // Prefilled with Word's own default, so Apply starts enabled.
    const formula = dialog!.querySelector<HTMLInputElement>('input[aria-label="Formula"]')!;
    expect(formula.value).toBe("=SUM(ABOVE)");
    expect(apply.disabled).toBe(false);
    // A formula outside the simple tier disables Apply rather than submitting
    // an instruction the engine refuses.
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
      setter.call(formula, "=IF(1,2,3)");
      formula.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(dialog!.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]")!.disabled).toBe(true);
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
      setter.call(formula, "=SUM(LEFT)");
      formula.dispatchEvent(new Event("input", { bubbles: true }));
    });
    const fmt = dialog!.querySelector<HTMLInputElement>('input[aria-label="Number format"]')!;
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
      setter.call(fmt, "#,##0.00");
      fmt.dispatchEvent(new Event("input", { bubbles: true }));
    });
    await click(t.container.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]")!);
    expect(t.insertTableFormula).toHaveBeenLastCalledWith("=SUM(LEFT)", "#,##0.00");
    await t.unmount();
  });
});
