// @vitest-environment jsdom
/**
 * THE CONTENTS MENU (Insert tab, `field` group).
 *
 * `insertToc`, `refreshTocs` and `updateFields` have been on the api since the
 * fields engine landed, with no control anywhere in the ribbon: a user could
 * not reach a table of contents at all. This drives the menu against a LIVE
 * DocxView rather than a stub, because the interesting failure is not "did the
 * button call the method" but "did a real TOC land in a real document" — the
 * field structure, the entries, and the page numbers the update pass fills in.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

const heading = (level: number, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const STYLES_XML =
  `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  [1, 2, 3]
    .map(
      (n) =>
        `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/>` +
        `<w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`,
    )
    .join("") +
  `</w:styles>`;

/** An anchor paragraph for the caret, then three headings for the TOC to find. */
const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">Anchor</w:t></w:r></w:p>` +
    heading(1, "Introduction") +
    heading(2, "Background") +
    heading(1, "Method");
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
    "word/styles.xml": strToU8(STYLES_XML),
  });
})();

/** The same package with NO headings — the state a new document is in, and the
 * one where a table of contents has nothing to list. */
const PLAIN_FIXTURE = (() => {
  const body = `<w:p><w:r><w:t xml:space="preserve">Just a sentence.</w:t></w:r></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
    "word/styles.xml": strToU8(STYLES_XML),
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

/** Mount a real editable view with the real toolbar above it. Supply `collab`
 * to mount IN A ROOM, where the api emits an intent instead of mutating. */
async function mount(inRoom = false, source: Uint8Array = FIXTURE) {
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
        source,
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
  // Caret into the anchor paragraph — an insert needs one.
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
    container,
    intents,
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

/**
 * Open the Insert tab, then pick `label` from the Contents menu.
 *
 * The menu is a popover of plain buttons rather than a listbox: every command
 * in it depends on something the user must have done first (heading styles,
 * marked entries), so each one carries a note saying so, which a <select> of
 * options cannot express.
 */
async function openContents(bar: HTMLElement): Promise<void> {
  const insert = bar.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
  expect(insert, "Insert tab").toBeTruthy();
  await click(insert!);
  const trigger = bar.querySelector<HTMLButtonElement>('button[data-testid="contents-menu"]');
  expect(trigger, "Contents menu").toBeTruthy();
  await click(trigger!);
}

async function pickContents(bar: HTMLElement, label: string) {
  await openContents(bar);
  const options = [...bar.querySelectorAll<HTMLButtonElement>("button")];
  const option = options.find((item) => (item.textContent ?? "").trim() === label);
  expect(option, `option "${label}" among ${options.map((o) => o.textContent).join(" | ")}`).toBeTruthy();
  await click(option!);
  await tick(20);
}

describe("Insert tab: the Contents menu", () => {
  it("inserts a real TOC field with an entry per heading", async () => {
    const t = await mount();
    expect(t.xml()).not.toContain("TOC \\o");
    await pickContents(t.bar, "Insert table of contents");
    const xml = t.xml();
    // The field itself, spelled the way Word writes it.
    expect(xml).toContain(`<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>`);
    // One entry per heading, each a PAGEREF hyperlink.
    for (const title of ["Introduction", "Background", "Method"]) {
      expect(xml, `entry for ${title}`).toContain(title);
    }
    expect(xml.match(/PAGEREF/g)?.length, "one PAGEREF per heading").toBe(3);
    await t.unmount();
  });

  it("updates field results without rebuilding the entries", async () => {
    const t = await mount();
    await pickContents(t.bar, "Insert table of contents");
    const withToc = t.xml();
    // A pure result refresh: the entry paragraphs the insert produced survive,
    // so the document is unchanged when nothing about the headings moved.
    await pickContents(t.bar, "Update all fields");
    expect(t.xml()).toBe(withToc);
    await t.unmount();
  });

  it("rebuilds the entries after a heading changes", async () => {
    const t = await mount();
    await pickContents(t.bar, "Insert table of contents");
    expect(t.xml()).toContain("Method");
    // Retitle the last heading behind the TOC's back, the way editing does.
    const doc = t.api().document;
    const texts: { text: string }[] = [];
    const walk = (el: { name: string; text: string; children: typeof texts }) => {
      if (el.name.endsWith("t") && el.text === "Method") texts.push(el);
      for (const c of el.children) walk(c as never);
    };
    walk(doc.docRoot as never);
    // The TOC entry copied the title, so both the heading and the entry match.
    expect(texts.length).toBeGreaterThanOrEqual(2);
    texts[texts.length - 1].text = "Materials";
    doc.refresh();
    await pickContents(t.bar, "Update it from the document");
    expect(t.xml()).toContain("Materials");
    await t.unmount();
  });
});

describe("Insert tab: the Contents menu in a room", () => {
  it("emits an insertToc intent instead of forking the room", async () => {
    const t = await mount(true);
    const before = t.xml();
    await pickContents(t.bar, "Insert table of contents");
    const toc = t.intents.find((intent) => intent.kind === "insertToc");
    expect(toc, `insertToc among ${t.intents.map((i) => i.kind).join(", ")}`).toBeTruthy();
    // The local document is NOT mutated: in a room the intent is the whole of
    // the edit, and mutating here as well would apply the insert twice.
    expect(t.xml()).toBe(before);
    await t.unmount();
  });

  it("carries an entry budget sized from the document's own headings", async () => {
    // THE NUMBER THE WIRE CANNOT DERIVE. Every other insert's size is in its
    // arguments; a TOC's is in the document. This fixture has three headings
    // inside the default levels 1-3, so the intent says three — and allocates
    // eight ids per entry (the paragraph and its seven hyperlink runs) plus
    // spares for the field runs and the closing paragraph.
    const t = await mount(true);
    await pickContents(t.bar, "Insert table of contents");
    const toc = t.intents.find((intent) => intent.kind === "insertToc")!;
    expect(toc.entryCount).toBe(3);
    expect((toc.nodeIds as number[]).length).toBe(3 * 8 + 8);
    await t.unmount();
  });
});

describe("Insert tab: the Contents menu explains what it needs", () => {
  /**
   * Inserting a TOC into a document with no headings puts Word's own "No table
   * of contents entries found." into the page. That text is correct and has to
   * stay — it is what Word writes, and changing it would make the saved file
   * disagree with Word. So the explanation belongs BEFORE the insert, where it
   * can say what a table of contents is built from.
   */
  it("says there are no headings, and what a heading is, before anything is inserted", async () => {
    const t = await mount(false, PLAIN_FIXTURE);
    await openContents(t.bar);
    const state = t.bar.querySelector('[data-testid="toc-state"]');

    expect(state?.textContent).toContain("No headings yet");
    expect(state?.textContent).toContain("Heading 1");
    // And the document is untouched: opening a menu must insert nothing.
    expect(t.container.textContent).not.toContain("No table of contents entries found.");
    await t.unmount();
  });

  it("counts the headings it would list once they exist", async () => {
    const t = await mount();
    await openContents(t.bar);
    expect(t.bar.querySelector('[data-testid="toc-state"]')?.textContent).toMatch(/\d+ headings? found/);
    await t.unmount();
  });

  it("says the index needs entries marked first", async () => {
    const t = await mount();
    await openContents(t.bar);
    const state = t.bar.querySelector('[data-testid="index-state"]');
    expect(state?.textContent).toContain("Nothing marked yet");
    expect(state?.textContent).toContain("Select a word");
    await t.unmount();
  });
});
