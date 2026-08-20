// @vitest-environment jsdom
/**
 * Multilevel numbering gallery: applyNumberingPreset compiles onto the
 * existing setListType / setNumberingLevel operations — locally it lands a
 * multilevel definition in numbering.xml; in a room it EMITS those same
 * registered intents and mutates nothing outside the canonical apply.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8, unzipSync, strFromU8 } from "fflate";
import { DocxView, NUMBERING_PRESETS, type DocxViewApi } from "../src/index.js";
import { type DocxDocument } from "@wordinweb/core";

const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">first item</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">second item</w:t></w:r></w:p>`;
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

async function mount(collabIntents?: Array<{ kind: string } & Record<string, unknown>>) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let nextId = 900_000;
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(
      createElement(DocxView, {
        source: FIXTURE,
        editable: true,
        onReady: (api: DocxViewApi) => { seen.api = api; },
        onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
        ...(collabIntents
          ? {
              collab: {
                submit: (intent: never) => { collabIntents.push(intent); },
                submitOp: (intent: { kind: string } & Record<string, unknown>) => { collabIntents.push(intent); },
                allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
              },
            }
          : {}),
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  const click = async () => {
    const page = container.querySelector<HTMLElement>(".dxw-page")!;
    const span = page.querySelector("span") ?? page;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      span.dispatchEvent(new MouseEvent("mousedown", opts));
      span.dispatchEvent(new MouseEvent("mouseup", opts));
    });
    await tick();
  };
  return {
    container, click,
    api: () => seen.api!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

describe("multilevel numbering gallery", () => {
  it("locally lands a multilevel definition in numbering.xml", async () => {
    const m = await mount();
    await m.click();
    let ok = false;
    await act(async () => { ok = m.api().applyNumberingPreset("decimalNested"); });
    expect(ok).toBe(true);
    const saved = unzipSync(m.api().save());
    const numbering = strFromU8(saved["word/numbering.xml"]);
    expect(numbering).toContain(`w:lvlText w:val="%1."`);
    expect(numbering).toContain(`w:lvlText w:val="%1.%2."`);
    expect(numbering).toContain(`w:lvlText w:val="%1.%2.%3."`);
    // The caret's paragraph joined the list.
    expect(strFromU8(saved["word/document.xml"])).toContain("numPr");
    await m.unmount();
  });

  it("the legal preset writes Article/Section labels", async () => {
    const m = await mount();
    await m.click();
    await act(async () => { m.api().applyNumberingPreset("articleSection"); });
    const numbering = strFromU8(unzipSync(m.api().save())["word/numbering.xml"]);
    expect(numbering).toContain(`w:lvlText w:val="Article %1"`);
    expect(numbering).toContain(`w:lvlText w:val="Section %1.%2"`);
    expect(numbering).toContain(`w:numFmt w:val="upperRoman"`);
    await m.unmount();
  });

  it("in a room the preset EMITS the registered intents and mutates nothing itself", async () => {
    const intents: Array<{ kind: string } & Record<string, unknown>> = [];
    const m = await mount(intents);
    await m.click();
    await act(async () => { m.api().applyNumberingPreset("outline"); });
    const kinds = intents.map((i) => i.kind);
    expect(kinds[0]).toBe("setListType");
    expect(kinds.filter((k) => k === "setNumberingLevel")).toHaveLength(NUMBERING_PRESETS.outline.levels.length);
    await m.unmount();
  });
});
