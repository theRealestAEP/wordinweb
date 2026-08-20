// @vitest-environment jsdom
/**
 * #152: a table can be inserted with the keyboard alone.
 *
 * The grid was 80 `<div onClick>` driven by `onMouseEnter`, and it is the ONLY
 * route to inserting a table anywhere in the product — no shortcut binds it,
 * the context menu does not offer it, and the help guide names the grid as the
 * way. So this was not an accessibility nicety: a keyboard user could not
 * insert a table at all.
 *
 * WHAT THIS FILE ASSERTS, AND WHY IT IS THIS AND NOT SOMETHING CHEAPER. The
 * size is chosen by POINTING. Making the 80 cells focusable would satisfy
 * "the panel has a control you can reach" and leave the feature exactly as
 * unusable, because reaching a cell is not choosing a size. So the assertion
 * is on the document: a table of the CHOSEN dimensions exists in the file
 * afterwards, and the only input was keys.
 */
import { describe, expect, it } from "vitest";
import { act } from "react";
import { serializeXml, type DocxDocument } from "@wordinweb/core";
import { createRoot } from "react-dom/client";
import { createElement } from "react";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import {
  caretIntoBody,
  click,
  controlByTip,
  focusableControls,
  openSurface,
  pressKey,
  selectTab,
  tick,
} from "./popover-smoke-harness.js";

class StubResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as unknown as { ResizeObserver: unknown }).ResizeObserver = StubResizeObserver;

const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t xml:space="preserve">Body text</w:t></w:r></w:p></w:body></w:document>`,
  ),
});

/** A view plus a bar, keeping the parsed document so the file can be read. */
async function mount() {
  const page = document.createElement("div");
  document.body.appendChild(page);
  const pageRoot = createRoot(page);
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const barRoot = createRoot(bar);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    pageRoot.render(createElement(DocxView, {
      source: FIXTURE,
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
    }));
  });
  for (let i = 0; i < 40 && !page.querySelector(".dxw-page"); i++) await tick();
  await act(async () => { barRoot.render(createElement(DocxToolbar, { api: seen.api })); });
  return {
    bar,
    page,
    api: () => seen.api!,
    xml: () => serializeXml(seen.doc!.docRoot),
    unmount: async () => {
      await act(async () => { barRoot.unmount(); pageRoot.unmount(); });
      bar.remove();
      page.remove();
    },
  };
}

/** The shape of the first table in the file: rows, and columns in row one. */
function firstTableShape(xml: string): { rows: number; cols: number } | null {
  const table = /<w:tbl>[\s\S]*?<\/w:tbl>/.exec(xml)?.[0];
  if (!table) return null;
  const rows = table.match(/<w:tr[ >]/g)?.length ?? 0;
  const firstRow = /<w:tr[ >][\s\S]*?<\/w:tr>/.exec(table)?.[0] ?? "";
  return { rows, cols: firstRow.match(/<w:tc[ >]/g)?.length ?? 0 };
}

/** Press a key on whatever currently has focus, as a real keyboard does. */
async function typeKey(key: string) {
  const target = (document.activeElement as HTMLElement | null) ?? document.body;
  await act(async () => {
    target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
    target.dispatchEvent(new KeyboardEvent("keyup", { key, bubbles: true, cancelable: true }));
  });
  await tick(0);
}

describe("#152: inserting a table without a mouse", () => {
  it("inserts a table of the size the arrow keys chose", async () => {
    const t = await mount();
    try {
      await caretIntoBody(t, "Body text");
      await selectTab(t.bar, "insert");
      expect(firstTableShape(t.xml()), "no table to begin with").toBeNull();

      const { panel } = await openSurface(t.bar, "Table");

      // Enter the grid the way Tab would: the single cell in the tab order.
      const entry = panel.querySelector<HTMLButtonElement>('[data-dxw-table-size][tabindex="0"]');
      expect(entry, "the grid offers exactly one tab stop, not eighty").toBeTruthy();
      expect(panel.querySelectorAll('[data-dxw-table-size][tabindex="0"]')).toHaveLength(1);
      entry!.focus();
      await tick(0);

      // 1x1 to 3x4: two presses down, three right. Nothing here is a click.
      await typeKey("ArrowDown");
      await typeKey("ArrowDown");
      await typeKey("ArrowRight");
      await typeKey("ArrowRight");
      await typeKey("ArrowRight");

      expect(panel.textContent, "the panel shows the size the keys chose").toContain("3 × 4");

      // Enter on a focused button is a click, which jsdom does not synthesise.
      (document.activeElement as HTMLButtonElement).click();
      await tick();

      expect(firstTableShape(t.xml()), "a 3x4 table reached the file").toEqual({ rows: 3, cols: 4 });
    } finally {
      await t.unmount();
    }
  });

  it("moves the size with the arrows rather than only the focus", async () => {
    // The distinction the whole ticket turns on. A grid whose cells take focus
    // but whose SIZE still needs a pointer would pass a focus-only check and
    // remain unusable.
    const t = await mount();
    try {
      await selectTab(t.bar, "insert");
      const { panel } = await openSurface(t.bar, "Table");
      expect(panel.textContent, "nothing chosen yet").toContain("Insert table");

      panel.querySelector<HTMLButtonElement>('[data-dxw-table-size="1x1"]')!.focus();
      await tick(0);
      await typeKey("ArrowRight");
      await typeKey("ArrowDown");

      expect(panel.textContent).toContain("2 × 2");
      expect(
        document.activeElement?.getAttribute("data-dxw-table-size"),
        "focus followed the size, so the next arrow moves from here",
      ).toBe("2x2");
    } finally {
      await t.unmount();
    }
  });

  it("clamps at the edges instead of wrapping or escaping the grid", async () => {
    const t = await mount();
    try {
      await selectTab(t.bar, "insert");
      const { panel } = await openSurface(t.bar, "Table");
      panel.querySelector<HTMLButtonElement>('[data-dxw-table-size="1x1"]')!.focus();
      await tick(0);

      await typeKey("ArrowUp");
      await typeKey("ArrowLeft");
      expect(panel.textContent, "held at the first cell").toContain("1 × 1");

      for (let i = 0; i < 12; i++) await typeKey("ArrowRight");
      for (let i = 0; i < 12; i++) await typeKey("ArrowDown");
      expect(panel.textContent, "held at the last cell of an 8x10 grid").toContain("8 × 10");
    } finally {
      await t.unmount();
    }
  });

  it("offers the table operations and cell fills to the keyboard too", async () => {
    // The rest of the same panel was the same defect: menu rows and fill
    // swatches were divs, so a keyboard could reach none of them either.
    const t = await mount();
    try {
      await selectTab(t.bar, "insert");
      const { panel } = await openSurface(t.bar, "Table");
      const names = focusableControls(panel).map((el) => el.textContent?.trim() || el.getAttribute("aria-label"));
      expect(names, "the row operations are reachable").toContain("Insert row above");
      expect(names, "so is deleting the table").toContain("Delete table");
      expect(names, "and the cell fills").toContain("Cell fill #FFF2CC");
      expect(names, "including clearing one").toContain("No cell fill");
    } finally {
      await t.unmount();
    }
  });

  it("starts from nothing chosen when it is reopened", async () => {
    // A size swept by a pointer and then dismissed is not a choice; carrying
    // it into the next open would offer to insert something nobody asked for.
    const t = await mount();
    try {
      await selectTab(t.bar, "insert");
      const first = await openSurface(t.bar, "Table");
      first.panel.querySelector<HTMLButtonElement>('[data-dxw-table-size="4x4"]')!.focus();
      await tick(0);
      expect(first.panel.textContent).toContain("4 × 4");

      await pressKey(first.panel, "Escape");
      await tick(0);

      const second = await openSurface(t.bar, "Table");
      expect(second.panel.textContent, "the old size did not survive the close").toContain("Insert table");
    } finally {
      await t.unmount();
    }
  });
});

describe("#152: the highlight swatches", () => {
  it("are buttons a keyboard can reach, each announcing its colour", async () => {
    const t = await mount();
    try {
      await selectTab(t.bar, "home");
      const { panel } = await openSurface(t.bar, "Highlight color");
      const names = focusableControls(panel).map((el) => el.getAttribute("aria-label"));
      expect(names).toEqual([
        "Highlight yellow",
        "Highlight green",
        "Highlight cyan",
        "Highlight magenta",
        "Highlight lightGray",
        "No highlight",
      ]);
    } finally {
      await t.unmount();
    }
  });

  it("applies the highlight when its button is activated", async () => {
    const t = await mount();
    try {
      await caretIntoBody(t, "Body text");
      t.api().find("Body text");
      await tick();
      await selectTab(t.bar, "home");
      const { panel } = await openSurface(t.bar, "Highlight color");
      await click(panel.querySelector<HTMLButtonElement>('[data-dxw-highlight="green"]')!);
      await tick();
      expect(t.xml(), "the highlight reached the file").toContain('w:highlight w:val="green"');
    } finally {
      await t.unmount();
    }
  });

  it("marks the current highlight as pressed", async () => {
    const t = await mount();
    try {
      await selectTab(t.bar, "home");
      const { panel } = await openSurface(t.bar, "Highlight color");
      // Nothing highlighted at the caret, so "No highlight" is the active one.
      expect(
        panel.querySelector('[data-dxw-highlight=""]')?.getAttribute("aria-pressed"),
      ).toBe("true");
      expect(
        panel.querySelector('[data-dxw-highlight="yellow"]')?.getAttribute("aria-pressed"),
      ).toBe("false");
      expect(controlByTip(t.bar, "Highlight color"), "opened from the bar").toBeTruthy();
    } finally {
      await t.unmount();
    }
  });
});
