// @vitest-environment jsdom
/**
 * THE CITATIONS MENU (Insert tab, `citations` group) — the References
 * cluster: manage sources, cite one, pick the style, insert the generated
 * bibliography. Driven against a LIVE DocxView (the contents-toolbar
 * discipline): the interesting failure is not "did the button call the
 * method" but "did a real source land in a real sources part, and did a real
 * CITATION / BIBLIOGRAPHY field land in the document".
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
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8(
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body><w:p><w:r><w:t xml:space="preserve">Anchor</w:t></w:r></w:p></w:body></w:document>`,
  ),
  "word/styles.xml": strToU8(
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
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
  return {
    bar,
    intents,
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

/** Open the Insert tab and the Citations popover. */
async function openCitations(bar: HTMLElement): Promise<HTMLElement> {
  const insert = bar.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
  expect(insert, "Insert tab").toBeTruthy();
  await click(insert!);
  const trigger = [...bar.querySelectorAll<HTMLButtonElement>("button")].find(
    (b) => b.title === "Citations and bibliography",
  );
  expect(trigger, "Citations button").toBeTruthy();
  await click(trigger!);
  const menu = bar.querySelector<HTMLElement>("[data-dxw-citations-menu]");
  expect(menu, "Citations popover").toBeTruthy();
  return menu!;
}

async function fillNewSource(bar: HTMLElement, menu: HTMLElement) {
  const newSource = [...menu.querySelectorAll<HTMLButtonElement>("button")].find(
    (b) => (b.textContent ?? "").trim() === "New source",
  );
  expect(newSource, "New source button").toBeTruthy();
  await click(newSource!);
  const field = (label: string) => {
    const input = bar.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`);
    expect(input, `${label} input`).toBeTruthy();
    return input!;
  };
  await type(field("Authors"), "Doe, Jane");
  await type(field("Title"), "A Study of Things");
  await type(field("Year"), "2003");
  await type(field("Publisher"), "Acme Press");
  await type(field("Source tag"), "Doe03");
  const add = [...bar.querySelectorAll<HTMLButtonElement>("button")].find(
    (b) => (b.textContent ?? "").trim() === "Add source",
  );
  await click(add!);
  await tick(20);
}

describe("Insert tab: the Citations menu", () => {
  it("adds a source, cites it, and inserts the generated bibliography", async () => {
    const t = await mount();
    const menu = await openCitations(t.bar);
    await fillNewSource(t.bar, menu);

    // The source landed in a real b:Sources part.
    expect(t.api().listCitationSources().map((s) => s.tag)).toEqual(["Doe03"]);
    expect(t.api().getCitationStyle()).toBe("APA");

    // Cite it: a CITATION field with the APA parenthetical as cached result.
    const cite = [...t.bar.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.title === "Cite Doe03");
    expect(cite, "Cite button").toBeTruthy();
    await click(cite!);
    await tick(20);
    expect(t.xml()).toContain(`w:instr=" CITATION Doe03 \\l 1033 "`);
    expect(t.xml()).toContain("(Doe, 2003)");

    // Insert the bibliography: the generated field, one entry per source.
    const menu2 = await openCitations(t.bar);
    const bib = [...menu2.querySelectorAll<HTMLButtonElement>("button")].find(
      (b) => (b.textContent ?? "").trim() === "Insert bibliography",
    );
    await click(bib!);
    await tick(20);
    expect(t.xml()).toContain(`<w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText>`);
    expect(t.xml()).toContain("Doe, J. (2003). A Study of Things. Acme Press.");
    await t.unmount();
  });

  it("switches the style, restyling citations and the bibliography together", async () => {
    const t = await mount();
    const menu = await openCitations(t.bar);
    await fillNewSource(t.bar, menu);
    const cite = [...t.bar.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.title === "Cite Doe03");
    await click(cite!);
    await tick(20);
    const menu2 = await openCitations(t.bar);
    const bib = [...menu2.querySelectorAll<HTMLButtonElement>("button")].find(
      (b) => (b.textContent ?? "").trim() === "Insert bibliography",
    );
    await click(bib!);
    await tick(20);

    const menu3 = await openCitations(t.bar);
    const select = menu3.querySelector<HTMLSelectElement>('select[aria-label="Citation style"]')!;
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!;
      setter.call(select, "MLA");
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await tick(20);
    expect(t.api().getCitationStyle()).toBe("MLA");
    // Both renderers moved in the one gesture.
    expect(t.xml()).toContain("(Doe)");
    expect(t.xml()).toContain("Doe, Jane. A Study of Things. Acme Press, 2003.");
    expect(t.xml()).not.toContain("(Doe, 2003)");
    await t.unmount();
  });

  it("refuses to delete a cited source, with the reason on screen", async () => {
    const t = await mount();
    const menu = await openCitations(t.bar);
    await fillNewSource(t.bar, menu);
    const cite = [...t.bar.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.title === "Cite Doe03");
    await click(cite!);
    await tick(20);
    const menu2 = await openCitations(t.bar);
    const del = [...menu2.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.title === "Delete Doe03");
    expect(del, "Delete button").toBeTruthy();
    await click(del!);
    await tick(20);
    expect(t.api().listCitationSources().map((s) => s.tag)).toEqual(["Doe03"]);
    expect(t.bar.querySelector('[role="alert"]')?.textContent).toContain("cited");
    await t.unmount();
  });
});

describe("Insert tab: the Citations menu in a room", () => {
  it("emits intents instead of forking the room", async () => {
    const t = await mount(true);
    const before = t.xml();
    const menu = await openCitations(t.bar);
    await fillNewSource(t.bar, menu);
    expect(t.intents.map((i) => i.kind)).toContain("createCitationSource");
    // In a room the intent is the whole of the edit; the local document (and
    // the local sources part) mutate only when the op comes back applied.
    expect(t.xml()).toBe(before);
    expect(t.api().listCitationSources()).toEqual([]);
    await t.unmount();
  });
});
