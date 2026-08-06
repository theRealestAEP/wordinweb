// @vitest-environment jsdom
/**
 * THE STYLES PANE (Home tab).
 *
 * listStyles/createStyle/modifyStyle/deleteStyle have been on the api since
 * the styles work landed, reachable only from the two dropdowns that APPLY a
 * style. Nothing in the toolbar could define one, rename one, or say how much
 * of the document a style was holding up.
 *
 * Driven against a LIVE DocxView, and asserted against styles.xml, because
 * the load-bearing property is not "which method ran" but WHAT LANDED IN THE
 * DEFINITION — in particular that modifying a style writes only the property
 * the user moved. The form prefills from the RESOLVED preview, so a form that
 * saved itself wholesale would copy every inherited property onto the
 * definition and quietly cut it out of its own cascade.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { DocxToolbar } from "../src/toolbar.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

/** Normal, a bold Body Text based on it, and an unused Caption. */
const STYLES_XML =
  `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/>` +
  `<w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="Caption"/>` +
  `<w:basedOn w:val="Normal"/></w:style>` +
  `</w:styles>`;

const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">plain one</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t xml:space="preserve">styled one</w:t></w:r></w:p>`;
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
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  await act(async () => {
    barRoot.render(createElement(DocxToolbar, { api: seen.api }));
  });
  // Caret into the first, unstyled paragraph.
  seen.api!.find("plain one");
  await tick();
  await act(async () => {
    const target =
      (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft", bubbles: true, cancelable: true }));
    await new Promise((r) => setTimeout(r, 2));
  });
  return {
    bar,
    api: () => seen.api!,
    body: () => serializeXml(seen.doc!.docRoot),
    styles: () => serializeXml(seen.doc!.stylesTree()!),
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

async function openPane(bar: HTMLElement) {
  const trigger = [...bar.querySelectorAll<HTMLButtonElement>("button")].find((b) =>
    (b.getAttribute("title") ?? b.getAttribute("data-tip") ?? "").startsWith("Styles pane"),
  );
  expect(trigger, "Styles button").toBeTruthy();
  await click(trigger!);
  const pane = document.querySelector<HTMLElement>("[data-dxw-styles-pane]");
  expect(pane, "styles pane").toBeTruthy();
  return pane!;
}

function control(pane: HTMLElement, label: string): HTMLElement {
  const el = pane.querySelector<HTMLElement>(`[aria-label="${label}"]`);
  expect(el, `control "${label}"`).toBeTruthy();
  return el!;
}

/** The w:style element for `id`, as XML. */
function definition(styles: string, id: string): string {
  const start = styles.indexOf(`w:styleId="${id}"`);
  expect(start, `definition for ${id}`).toBeGreaterThan(-1);
  const open = styles.lastIndexOf("<w:style ", start);
  return styles.slice(open, styles.indexOf("</w:style>", start) + 10);
}

describe("Home tab: the styles pane", () => {
  it("lists the document's styles with what each one is holding up", async () => {
    const t = await mount();
    const pane = await openPane(t.bar);
    const rows = [...pane.querySelectorAll<HTMLButtonElement>('button[aria-label^="Apply "]')].map(
      (b) => b.textContent ?? "",
    );
    expect(rows.some((row) => row.includes("Body Text") && row.includes("1 use"))).toBe(true);
    // A style nothing references says so rather than showing a bare zero.
    expect(rows.some((row) => row.includes("Caption") && row.includes("unused"))).toBe(true);
    await t.unmount();
  });

  it("applies a paragraph style to the caret's paragraph on click", async () => {
    const t = await mount();
    const pane = await openPane(t.bar);
    await click(control(pane, "Apply Caption"));
    await tick(20);
    expect(t.body()).toContain(`<w:pStyle w:val="Caption"/>`);
    await t.unmount();
  });

  it("creates a style, with the id Word would derive from the name", async () => {
    const t = await mount();
    const pane = await openPane(t.bar);
    await click([...pane.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "New style")!);
    await type(control(pane, "Style name") as HTMLInputElement, "Pull Quote");
    await click(control(pane, "Italic"));
    await type(control(pane, "Style font size (points)") as HTMLInputElement, "14");
    await type(control(pane, "Style alignment") as HTMLSelectElement, "center");
    await click(pane.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]")!);
    await tick(20);
    const style = definition(t.styles(), "PullQuote");
    expect(style).toContain(`<w:name w:val="Pull Quote"/>`);
    expect(style).toContain(`<w:i/>`);
    // 14pt is 28 half-points.
    expect(style).toContain(`w:val="28"`);
    expect(style).toContain(`<w:jc w:val="center"/>`);
    expect(style).toContain(`w:customStyle="1"`);
    await t.unmount();
  });

  it("renames a style through modifyStyle without touching anything else", async () => {
    const t = await mount();
    const pane = await openPane(t.bar);
    await click(control(pane, "Modify Body Text"));
    await type(control(pane, "Style name") as HTMLInputElement, "Body Copy");
    await click(pane.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]")!);
    await tick(20);
    const style = definition(t.styles(), "BodyText");
    expect(style).toContain(`<w:name w:val="Body Copy"/>`);
    // The rename is the whole edit: the definition keeps its parent and its
    // single run property, and gains nothing.
    expect(style).toContain(`<w:basedOn w:val="Normal"/>`);
    expect(style).toContain(`<w:b/>`);
    await t.unmount();
  });

  it("writes ONLY the property the form moved, never the inherited ones", async () => {
    const t = await mount();
    const pane = await openPane(t.bar);
    // Caption resolves through Normal, so its preview is plain; Body Text's
    // preview is BOLD, inherited from its own definition. Editing Caption and
    // touching only italic must not stamp a bold or a size onto it.
    await click(control(pane, "Modify Caption"));
    await click(control(pane, "Italic"));
    await click(pane.querySelector<HTMLButtonElement>("[data-dxw-dialog-apply]")!);
    await tick(20);
    const style = definition(t.styles(), "Caption");
    expect(style).toContain(`<w:i/>`);
    expect(style).not.toContain(`<w:sz`);
    expect(style).not.toContain(`<w:color`);
    await t.unmount();
  });

  it("asks before deleting, and re-points the paragraphs that used the style", async () => {
    const t = await mount();
    const pane = await openPane(t.bar);
    // One click arms, and the label says so; the style survives it.
    await click(control(pane, "Delete Body Text"));
    expect(t.styles()).toContain(`w:styleId="BodyText"`);
    await click(control(pane, "Confirm delete Body Text"));
    await tick(20);
    expect(t.styles()).not.toContain(`w:styleId="BodyText"`);
    // The paragraph that used it inherits the deleted style's parent rather
    // than keeping a reference nothing declares.
    expect(t.body()).toContain(`<w:pStyle w:val="Normal"/>`);
    await t.unmount();
  });
});
