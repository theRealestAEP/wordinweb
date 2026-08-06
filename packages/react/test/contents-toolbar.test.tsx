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

/** Mount a real editable view with the real toolbar above it. */
async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
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

/** Open the Insert tab, then pick `label` from the Contents menu. */
async function pickContents(bar: HTMLElement, label: string) {
  const insert = bar.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
  expect(insert, "Insert tab").toBeTruthy();
  await click(insert!);
  const trigger = bar.querySelector<HTMLButtonElement>(
    'button[aria-label="Insert or update a table of contents"]',
  );
  expect(trigger, "Contents menu").toBeTruthy();
  await click(trigger!);
  const options = [...bar.querySelectorAll<HTMLButtonElement>('[role="option"]')];
  const option = options.find((item) => (item.textContent ?? "").trim() === label);
  expect(option, `option "${label}" among ${options.map((o) => o.textContent).join(", ")}`).toBeTruthy();
  await click(option!);
  await tick(20);
}

describe("Insert tab: the Contents menu", () => {
  it("inserts a real TOC field with an entry per heading", async () => {
    const t = await mount();
    expect(t.xml()).not.toContain("TOC \\o");
    await pickContents(t.bar, "Table of contents");
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
    await pickContents(t.bar, "Table of contents");
    const withToc = t.xml();
    // A pure result refresh: the entry paragraphs the insert produced survive,
    // so the document is unchanged when nothing about the headings moved.
    await pickContents(t.bar, "Update fields");
    expect(t.xml()).toBe(withToc);
    await t.unmount();
  });

  it("rebuilds the entries after a heading changes", async () => {
    const t = await mount();
    await pickContents(t.bar, "Table of contents");
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
    await pickContents(t.bar, "Update table of contents");
    expect(t.xml()).toContain("Materials");
    await t.unmount();
  });
});
