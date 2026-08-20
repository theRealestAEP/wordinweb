// @vitest-environment jsdom
/**
 * THE QUICK PARTS MENU (Insert tab, `quickParts` group): save the current
 * selection as a named building block, then insert it back at the caret from
 * a gallery grouped by category. Driven against a LIVE DocxView (the
 * citations-toolbar discipline): the interesting failure is not "did the
 * button call the method" but "did the block land in a real glossary part,
 * and did a real clone of it land in the document body".
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

const FIXTURE = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t xml:space="preserve">Best regards, Jane</w:t></w:r></w:p></w:body></w:document>`,
  ),
});

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

async function type(input: HTMLInputElement, value: string) {
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
    setter.call(input, value);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

async function mount(inRoom = false) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  const intents: ({ kind: string } & Record<string, unknown>)[] = [];
  let nextId = 900_000;
  const bar = document.createElement("div");
  document.body.appendChild(bar);
  const barRoot = createRoot(bar);
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
        ...(inRoom
          ? {
              collab: {
                submit: (intent: unknown) => {
                  intents.push(intent as { kind: string } & Record<string, unknown>);
                },
                submitOp: (intent: { kind: string } & Record<string, unknown>) => {
                  intents.push(intent);
                },
                allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
              },
            }
          : {}),
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  await act(async () => {
    barRoot.render(createElement(DocxToolbar, { api: seen.api }));
  });
  const page = container.querySelector<HTMLElement>(".dxw-page")!;
  const span = page.querySelector("span") ?? page;
  await act(async () => {
    const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    span.dispatchEvent(new MouseEvent("mousedown", opts));
    span.dispatchEvent(new MouseEvent("mouseup", opts));
  });
  await tick();
  const selectAll = async () => {
    const target = container.querySelector("textarea") ?? container;
    await act(async () => {
      target.dispatchEvent(new KeyboardEvent("keydown", { key: "a", ctrlKey: true, bubbles: true, cancelable: true }));
    });
    await tick();
  };
  return {
    bar,
    intents,
    selectAll,
    api: () => seen.api!,
    doc: () => seen.doc!,
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

/** Open the Insert tab and the Quick Parts popover. */
async function openQuickParts(bar: HTMLElement): Promise<HTMLElement> {
  const insert = bar.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
  expect(insert, "Insert tab").toBeTruthy();
  await click(insert!);
  const trigger = [...bar.querySelectorAll<HTMLButtonElement>("button")].find(
    (b) => b.title === "Quick Parts: save and reuse content",
  );
  expect(trigger, "Quick Parts button").toBeTruthy();
  await click(trigger!);
  const menu = bar.querySelector<HTMLElement>("[data-dxw-quickparts-menu]");
  expect(menu, "Quick Parts popover").toBeTruthy();
  return menu!;
}

async function saveSelectionAs(bar: HTMLElement, menu: HTMLElement, name: string, category = "") {
  const save = [...menu.querySelectorAll<HTMLButtonElement>("button")].find(
    (b) => (b.textContent ?? "").trim() === "Save selection as Quick Part",
  );
  expect(save, "Save selection button").toBeTruthy();
  await click(save!);
  await type(bar.querySelector<HTMLInputElement>('input[aria-label="Quick Part name"]')!, name);
  if (category) await type(bar.querySelector<HTMLInputElement>('input[aria-label="Quick Part category"]')!, category);
  const submit = [...bar.querySelectorAll<HTMLButtonElement>("button")].find(
    (b) => (b.textContent ?? "").trim() === "Save",
  );
  await click(submit!);
  await tick(20);
}

describe("Insert tab: the Quick Parts menu", () => {
  it("saves the selection under a name and category, then inserts a clone at the caret", async () => {
    const t = await mount();
    await t.selectAll();
    const menu = await openQuickParts(t.bar);
    await saveSelectionAs(t.bar, menu, "Sign-off", "Letters");

    expect(t.api().listBuildingBlocks()).toEqual([{ name: "Sign-off", category: "Letters" }]);

    // The popover stays open (in list view) after a successful save.
    const menu2 = t.bar.querySelector<HTMLElement>("[data-dxw-quickparts-menu]")!;
    const insertBtn = [...menu2.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.title === "Insert Sign-off");
    expect(insertBtn, "Insert button").toBeTruthy();
    await click(insertBtn!);
    await tick(20);
    // The clone landed as a second occurrence of the saved text.
    expect(t.xml().match(/Best regards, Jane/g)?.length).toBe(2);
    await t.unmount();
  });

  it("defaults the category to General, and groups the gallery by category", async () => {
    const t = await mount();
    await t.selectAll();
    const menu = await openQuickParts(t.bar);
    await saveSelectionAs(t.bar, menu, "Plain Block");
    expect(t.api().listBuildingBlocks()).toEqual([{ name: "Plain Block", category: "General" }]);
    const menu2 = t.bar.querySelector<HTMLElement>("[data-dxw-quickparts-menu]")!;
    expect(menu2.textContent).toContain("General");
    await t.unmount();
  });

  it("reports an error with nothing selected", async () => {
    const t = await mount();
    const menu = await openQuickParts(t.bar);
    await saveSelectionAs(t.bar, menu, "Nothing");
    expect(t.bar.querySelector('[role="alert"]')?.textContent).toContain("select");
    expect(t.api().listBuildingBlocks()).toEqual([]);
    await t.unmount();
  });

  it("deletes a saved Quick Part", async () => {
    const t = await mount();
    await t.selectAll();
    const menu = await openQuickParts(t.bar);
    await saveSelectionAs(t.bar, menu, "Temp");
    const menu2 = t.bar.querySelector<HTMLElement>("[data-dxw-quickparts-menu]")!;
    const del = [...menu2.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.title === "Delete Temp");
    expect(del, "Delete button").toBeTruthy();
    await click(del!);
    await tick(20);
    expect(t.api().listBuildingBlocks()).toEqual([]);
    await t.unmount();
  });
});

describe("Insert tab: the Quick Parts menu in a room", () => {
  it("emits intents instead of forking the room", async () => {
    const t = await mount(true);
    const before = t.xml();
    await t.selectAll();
    const menu = await openQuickParts(t.bar);
    await saveSelectionAs(t.bar, menu, "Sign-off");
    expect(t.intents.map((i) => i.kind)).toContain("createBuildingBlock");
    // In a room the intent is the whole of the edit; the local document (and
    // the local glossary part) mutate only when the op comes back applied.
    expect(t.xml()).toBe(before);
    expect(t.api().listBuildingBlocks()).toEqual([]);
    await t.unmount();
  });
});
