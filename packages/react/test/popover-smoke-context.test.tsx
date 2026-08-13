// @vitest-environment jsdom
/**
 * #146, part two: the popovers that exist only once something is selected.
 *
 * The Table Format tab is not on the bar until the caret is inside a table,
 * and the Format tab is not there until a drawing is selected, so the breadth
 * pass in `popover-smoke-all` cannot reach either — and the panels behind
 * them (cell fill, borders, the formula and properties dialogs, the shape's
 * size/position/rotation forms) are among the least-driven in the file.
 *
 * Same five invariants, one extra step to get there.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  caretIntoTable,
  click,
  controlByTip,
  mountToolbar,
  selectAShape,
  TABLE_FIXTURE,
  tick,
  watchConsole,
  type ConsoleWatch,
  type MountedToolbar,
} from "./popover-smoke-harness.js";
import { surfaceInvariants } from "./popover-smoke-invariants.js";

class StubResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = StubResizeObserver;

/**
 * Surfaces reachable only from the Table Format tab.
 *
 * The tab's plain action buttons ("Convert this table to text", "Delete the
 * current table") are left out: they do their work on the spot and open no
 * panel, so there is nothing here to assert about them.
 */
const TABLE_SURFACES = [
  "Cell fill color",
  "Align text inside the current cell",
  "Edit rows and columns around the current cell",
  "Merge or split the current cell",
  "Sort rows by the current column (repeating header rows stay in place)",
  "Set or clear the borders of the table or the current cell",
  "Apply a table style defined in this document",
  "Choose how the table sizes its columns",
  "Repeat leading rows at the top of every page",
  'Insert a formula field ("=SUM(ABOVE)") in the current cell',
  "Table properties: exact widths, cell margins and header rows",
];

/**
 * Surfaces reachable only with a drawing selected, same exclusion.
 *
 * "Edit shape text" is excluded for a second reason: it opens no panel at
 * all. It calls the engine's `editText`, which puts the caret inside the
 * shape's own story so the Home tab's font controls apply to it — a mode
 * switch, not a dialog, so Escape and a click outside are not its contract.
 */
const OBJECT_SURFACES = [
  "Fill color",
  "Outline color, weight, and style",
  "Wrap",
  "How the shape's text and its box fit each other",
  "Exact size",
  "Exact page position",
  "Set rotation",
];

describe("Table Format tab", () => {
  let toolbar: MountedToolbar;
  let watch: ConsoleWatch;

  beforeEach(async () => {
    localStorage.clear();
    watch = watchConsole();
    toolbar = await mountToolbar(TABLE_FIXTURE);
    await caretIntoTable(toolbar);
    const tab = toolbar.bar.querySelector<HTMLButtonElement>('button[data-tab="tableFormat"]');
    expect(tab, "the tab appears with the caret in a table").toBeTruthy();
    await click(tab!);
    // Driving the caret is asynchronous, and React rightly warns that those
    // updates arrived outside act(). That is this file's setup talking, not
    // the bar's, so it is not carried into the checks.
    await tick();
    watch.reset();
  });

  afterEach(async () => {
    await toolbar.unmount();
    watch.stop();
  });

  it("carries every surface this suite drives", () => {
    for (const tip of TABLE_SURFACES) {
      expect(() => controlByTip(toolbar.bar, tip), `${tip} is on the tab`).not.toThrow();
    }
  });

  for (const tip of TABLE_SURFACES) {
    describe(tip, () => surfaceInvariants(tip, "table panel", () => ({ bar: toolbar.bar, watch })));
  }
});

describe("Format tab, with a shape selected", () => {
  let toolbar: MountedToolbar;
  let watch: ConsoleWatch;

  beforeEach(async () => {
    localStorage.clear();
    watch = watchConsole();
    toolbar = await mountToolbar();
    await selectAShape(toolbar);
    const tab = toolbar.bar.querySelector<HTMLButtonElement>('button[data-tab="format"]');
    expect(tab, "the tab appears with an object selected").toBeTruthy();
    await click(tab!);
    await tick();
    watch.reset();
  });

  afterEach(async () => {
    await toolbar.unmount();
    watch.stop();
  });

  it("carries every surface this suite drives", () => {
    for (const tip of OBJECT_SURFACES) {
      expect(() => controlByTip(toolbar.bar, tip), `${tip} is on the tab`).not.toThrow();
    }
  });

  for (const tip of OBJECT_SURFACES) {
    describe(tip, () => surfaceInvariants(tip, "object panel", () => ({ bar: toolbar.bar, watch })));
  }
});
