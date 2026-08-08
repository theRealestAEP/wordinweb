// @vitest-environment jsdom
/**
 * #111 — DocxViewApi.selectRange: select an exact text range by the
 * stable-addressed wire shape, in a LOCAL (non-collab) mount. The desktop
 * spellcheck's select-and-replace flow: encode an address, select it, then
 * replace through the typing path (insertSymbol).
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8, unzipSync, strFromU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { localName, type DocxDocument, type XmlElement } from "@wordinweb/core";

const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">The teh cat</w:t></w:r></w:p>` +
    // A word split across two runs ("beau" + "tiful") — the case the app's
    // spellcheck could squiggle but never replace.
    `<w:p><w:r><w:t xml:space="preserve">beau</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">tiful</w:t></w:r></w:p>`;
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

async function mount() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(
      createElement(DocxView, {
        source: FIXTURE,
        editable: true,
        onReady: (api: DocxViewApi) => { seen.api = api; },
        onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  return {
    container,
    api: seen.api!,
    doc: seen.doc!,
    unmount: async () => {
      await act(async () => root.unmount());
      container.remove();
    },
  };
}

/** Address a paragraph's nth run through the doc's stable-id table, the way
 * a host computes ranges (enableStableIds is public and idempotent). */
function runAddress(doc: DocxDocument, paraIdx: number, runIdx: number): { blockId: number; runId: number } {
  const ids = doc.enableStableIds();
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const p = body.children.filter((c) => localName(c.name) === "p")[paraIdx];
  const r = p.children.filter((c: XmlElement) => localName(c.name) === "r")[runIdx];
  return { blockId: ids.idOf(p)!, runId: ids.idOf(r)! };
}

describe("selectRange (#111)", () => {
  it("selects an exact wire range in a local mount and replaces it through the typing path", async () => {
    const t = await mount();
    const addr = runAddress(t.doc, 0, 0); // "The teh cat"
    // Select "teh" — [4, 7) in the run's wire basis.
    await act(async () => {
      expect(t.api.selectRange({ ...addr, start: 4, end: 7 })).toBe(true);
    });
    expect(t.container.querySelector(".dxw-sel")).toBeTruthy();
    await act(async () => {
      expect(t.api.insertSymbol("the")).toBe(true);
    });
    const saved = unzipSync(t.api.save());
    expect(strFromU8(saved["word/document.xml"])).toContain("The the cat");
    await t.unmount();
  });

  it("selects a word split across runs from several wire ranges", async () => {
    const t = await mount();
    const a = runAddress(t.doc, 1, 0); // "beau"
    const b = runAddress(t.doc, 1, 1); // "tiful"
    await act(async () => {
      expect(t.api.selectRange([
        { ...a, start: 0, end: 4 },
        { ...b, start: 0, end: 5 },
      ])).toBe(true);
    });
    await act(async () => {
      expect(t.api.insertSymbol("plain")).toBe(true);
    });
    const saved = unzipSync(t.api.save());
    const xml = strFromU8(saved["word/document.xml"]);
    expect(xml).toContain("plain");
    expect(xml).not.toContain("beau");
    expect(xml).not.toContain("tiful");
    await t.unmount();
  });

  it("declines ranges that resolve to nothing", async () => {
    const t = await mount();
    expect(t.api.selectRange({ blockId: 999999, runId: 999998, start: 0, end: 3 })).toBe(false);
    const addr = runAddress(t.doc, 0, 0);
    expect(t.api.selectRange({ ...addr, start: 3, end: 3 })).toBe(false);
    await t.unmount();
  });
});
