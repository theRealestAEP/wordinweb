// @vitest-environment jsdom
/**
 * THE TABLE FORMAT TAB'S TWO DIALOGS: Table Properties and Custom Border.
 *
 * The ribbon could already apply preset borders and a header band of 0, 1 or
 * 2 rows. It could not express a NUMBER — an exact table width, one column's
 * width, cell margins — nor any border beyond a ½pt black single rule, though
 * setTableWidth/setTableColumnWidth/setTableCellMargins/setTableBorders have
 * taken all of it since the tables work landed.
 *
 * Driven against a LIVE DocxView, because both dialogs are about values that
 * have to survive the trip into the file: the assertions read the resulting
 * w:tblW, w:gridCol, w:tblCellMar and w:tblBorders rather than call arguments.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

/** 2400 twips = 120pt per column, 4800 twips = 240pt total. */
function cell(text: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
}

const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">Before</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
    `<w:tr>${cell("r0c0")}${cell("r0c1")}</w:tr>` +
    `<w:tr>${cell("r1c0")}${cell("r1c1")}</w:tr></w:tbl>` +
    `<w:p><w:r><w:t xml:space="preserve">After</w:t></w:r></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
})();

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}

async function click(element: HTMLElement) {
  await act(async () => {
    element.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
  });
}

/** React tracks the DOM value itself, so a plain assignment is invisible to
 * it; the prototype setter is the standard way to make a synthetic edit look
 * like a user's. */
async function type(element: HTMLInputElement | HTMLSelectElement, value: string) {
  const proto = element instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto, "value")!.set!.call(element, value);
  await act(async () => {
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
  });
}

async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const barRoot = createRoot(bar);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(
      createElement(DocxView, {
        source: FIXTURE,
        editable: true,
        onReady: (api: DocxViewApi) => {
          seen.api = api;
        },
        onLoad: (info: { document: DocxDocument }) => {
          seen.doc = info.document;
        },
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  await act(async () => {
    barRoot.render(createElement(DocxToolbar, { api: seen.api }));
  });
  // Caret into the FIRST cell: everything here addresses the table around the
  // caret. find() leaves a selection, so collapse it the way a user would.
  seen.api!.find("r0c0");
  await tick();
  await act(async () => {
    const target =
      (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
    await new Promise((r) => setTimeout(r, 2));
  });
  expect(seen.api!.getTableCellFill(), "caret is inside the table").not.toBe(undefined);
  // The Table Format tab appears once the toolbar sees the caret in a table.
  await act(async () => {
    document.dispatchEvent(new Event("dxw-selection"));
  });
  await tick();
  const tab = bar.querySelector<HTMLButtonElement>('button[data-tab="tableFormat"]');
  expect(tab, "Table Format tab").toBeTruthy();
  await click(tab!);
  return {
    bar,
    api: () => seen.api!,
    xml: () => serializeXml(seen.doc!.docRoot),
    unmount: async () => {
      await act(async () => {
        barRoot.unmount();
        root.unmount();
      });
      bar.remove();
      container.remove();
    },
  };
}

function field(bar: HTMLElement, label: string): HTMLInputElement | HTMLSelectElement {
  const el = bar.querySelector<HTMLInputElement | HTMLSelectElement>(`[aria-label="${label}"]`)
    ?? document.querySelector<HTMLInputElement | HTMLSelectElement>(`[aria-label="${label}"]`);
  expect(el, `field "${label}"`).toBeTruthy();
  return el!;
}

async function openProperties(bar: HTMLElement) {
  const trigger = [...bar.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
    (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").startsWith("Table properties"),
  );
  expect(trigger, "Properties button").toBeTruthy();
  await click(trigger!);
}

async function applyDialog(label: string) {
  const dialog = document.querySelector<HTMLElement>(`[role="dialog"][aria-label="${label}"]`);
  expect(dialog, `dialog "${label}"`).toBeTruthy();
  const apply = dialog!.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]");
  expect(apply, "Apply button").toBeTruthy();
  await click(apply!);
  await tick(20);
}

describe("Table Format tab: the Table Properties dialog", () => {
  it("prefills from the document rather than from a guess", async () => {
    const t = await mount();
    await openProperties(t.bar);
    // 4800 twips is 240pt; 2400 twips per column is 120pt.
    expect((field(t.bar, "Table width unit") as HTMLSelectElement).value).toBe("pt");
    expect((field(t.bar, "Table width") as HTMLInputElement).value).toBe("240");
    expect((field(t.bar, "Column width (points)") as HTMLInputElement).value).toBe("120");
    expect((field(t.bar, "Repeating header rows") as HTMLInputElement).value).toBe("0");
    // The table declares no cell margins, so the boxes are blank rather than
    // claiming a zero the file does not contain.
    expect((field(t.bar, "Top cell margin (points)") as HTMLInputElement).value).toBe("");
    // …and it names the column the caret is actually in.
    const dialog = document.querySelector('[role="dialog"][aria-label="Table Properties"]')!;
    expect(dialog.textContent).toContain("Column 1 of 2");
    await t.unmount();
  });

  it("writes ONLY the boxes the user changed", async () => {
    const t = await mount();
    await openProperties(t.bar);
    await type(field(t.bar, "Top cell margin (points)") as HTMLInputElement, "6");
    await type(field(t.bar, "Left cell margin (points)") as HTMLInputElement, "9");
    await applyDialog("Table Properties");
    const xml = t.xml();
    // 6pt = 120 twips, 9pt = 180 twips.
    expect(xml).toContain(`<w:tblCellMar>`);
    expect(xml).toContain(`<w:top w:w="120" w:type="dxa"/>`);
    expect(xml).toContain(`<w:left w:w="180" w:type="dxa"/>`);
    // The width boxes were untouched, so the table keeps the exact tblW it
    // was authored with — no rewrite, no tracked change for a value nobody
    // edited.
    expect(xml).toContain(`<w:tblW w:w="4800" w:type="dxa"/>`);
    await t.unmount();
  });

  it("sets an exact table width in points and in percent", async () => {
    const t = await mount();
    await openProperties(t.bar);
    await type(field(t.bar, "Table width") as HTMLInputElement, "300");
    await applyDialog("Table Properties");
    // 300pt = 6000 twips.
    expect(t.xml()).toContain(`<w:tblW w:w="6000" w:type="dxa"/>`);

    await openProperties(t.bar);
    await type(field(t.bar, "Table width unit") as HTMLSelectElement, "pct");
    await type(field(t.bar, "Table width") as HTMLInputElement, "80");
    await applyDialog("Table Properties");
    // Word writes percent as fiftieths.
    expect(t.xml()).toContain(`<w:tblW w:w="4000" w:type="pct"/>`);
    await t.unmount();
  });

  it("resizes the caret's column, and re-totals the table", async () => {
    const t = await mount();
    await openProperties(t.bar);
    await type(field(t.bar, "Column width (points)") as HTMLInputElement, "180");
    await applyDialog("Table Properties");
    const xml = t.xml();
    // 180pt = 3600 twips on column 1; column 2 keeps its 2400.
    expect(xml).toContain(`<w:gridCol w:w="3600"/>`);
    expect(xml).toContain(`<w:gridCol w:w="2400"/>`);
    // A dxa table re-totals, so the three places that carry a width agree.
    expect(xml).toContain(`<w:tblW w:w="6000" w:type="dxa"/>`);
    await t.unmount();
  });

  it("sets a header band of any size, not only the menu's 0/1/2", async () => {
    const t = await mount();
    await openProperties(t.bar);
    await type(field(t.bar, "Repeating header rows") as HTMLInputElement, "2");
    await applyDialog("Table Properties");
    expect(t.xml().match(/<w:tblHeader\/>/g)?.length).toBe(2);
    // Re-opening reads the band back, which is what makes the dialog honest.
    await openProperties(t.bar);
    expect((field(t.bar, "Repeating header rows") as HTMLInputElement).value).toBe("2");
    await t.unmount();
  });
});

describe("Table Format tab: the Custom Border dialog", () => {
  async function openCustomBorder(bar: HTMLElement) {
    const trigger = bar.querySelector<HTMLButtonElement>(
      'button[aria-label="Set or clear the borders of the table or the current cell"]',
    );
    expect(trigger, "Borders menu").toBeTruthy();
    await click(trigger!);
    const option = [...bar.querySelectorAll<HTMLButtonElement>('[role="option"]')].find(
      (item) => (item.textContent ?? "").trim() === "Custom border…",
    );
    expect(option, "Custom border option").toBeTruthy();
    await click(option!);
  }

  it("writes a style, a weight and a color the presets cannot reach", async () => {
    const t = await mount();
    await openCustomBorder(t.bar);
    await type(field(t.bar, "Border style") as HTMLSelectElement, "dotDash");
    await type(field(t.bar, "Border width (points)") as HTMLSelectElement, "2.25");
    // The color menu is the same one the text and cell-fill colors use.
    const dialog = document.querySelector<HTMLElement>('[role="dialog"][aria-label="Custom Border"]')!;
    const colorTrigger = dialog.querySelector<HTMLButtonElement>("[data-dxw-color-trigger]");
    expect(colorTrigger, "Border color menu").toBeTruthy();
    await click(colorTrigger!);
    const swatch = dialog.querySelector<HTMLElement>('[data-dxw-color="#ff0000"]');
    expect(swatch, "red swatch").toBeTruthy();
    await click(swatch!);
    await applyDialog("Custom Border");
    const xml = t.xml();
    // 2.25pt = w:sz 18 (eighths of a point).
    expect(xml).toContain(`w:val="dotDash"`);
    expect(xml).toContain(`w:sz="18"`);
    expect(xml).toContain(`w:color="FF0000"`);
    await t.unmount();
  });

  it("offers only the edges the chosen scope owns", async () => {
    const t = await mount();
    await openCustomBorder(t.bar);
    const dialog = document.querySelector<HTMLElement>('[role="dialog"][aria-label="Custom Border"]')!;
    const edges = () =>
      [...dialog.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].map((c) =>
        c.getAttribute("aria-label"),
      );
    // Table scope has inside rules and no diagonals.
    expect(edges()).toContain("Inside horizontal");
    expect(edges()).not.toContain("Diagonal ↘");
    await type(field(t.bar, "Border scope") as HTMLSelectElement, "cell");
    // Cell scope is the mirror image — setTableBorders refuses the other set,
    // so offering it would be a checkbox that does nothing.
    expect(edges()).not.toContain("Inside horizontal");
    expect(edges()).toContain("Diagonal ↘");
    await t.unmount();
  });

  it("puts a rule on exactly the edges that are ticked", async () => {
    const t = await mount();
    await openCustomBorder(t.bar);
    const dialog = document.querySelector<HTMLElement>('[role="dialog"][aria-label="Custom Border"]')!;
    const box = (label: string) =>
      dialog.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!;
    // Default is the four outside edges; swap them for the inside rules.
    for (const label of ["Top", "Bottom", "Left", "Right"]) await click(box(label));
    await click(box("Inside horizontal"));
    await type(field(t.bar, "Border style") as HTMLSelectElement, "double");
    await applyDialog("Custom Border");
    const xml = t.xml();
    const borders = xml.slice(xml.indexOf("<w:tblBorders>"), xml.indexOf("</w:tblBorders>"));
    expect(borders).toContain(`<w:insideH`);
    expect(borders).not.toContain(`<w:top`);
    expect(borders).toContain(`w:val="double"`);
    await t.unmount();
  });
});
