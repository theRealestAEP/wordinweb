// @vitest-environment jsdom
/**
 * Find & Replace depth + Go To, driven against a LIVE DocxView: matching in
 * header/footnote stories, whole-word and match-case options, per-story
 * replace counts, and page/bookmark navigation.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8, unzipSync, strFromU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { type DocxDocument } from "@wordinweb/core";

const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">Body cat one</w:t></w:r></w:p>` +
    `<w:bookmarkStart w:id="1" w:name="mark"/><w:bookmarkEnd w:id="1"/>` +
    `<w:p><w:r><w:t xml:space="preserve">Second cat here</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">concatenate</w:t></w:r></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  const headerXml =
    `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:p><w:r><w:t xml:space="preserve">Header cat text</w:t></w:r></w:p></w:hdr>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`),
    "word/header1.xml": strToU8(headerXml),
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

describe("find & replace depth through the view api", () => {
  it("finds in every story, honors the options, and counts per story", async () => {
    const t = await mount();
    // "cat" appears twice in the body, once in the header, and inside
    // "concatenate" — whole-word excludes the latter.
    expect(t.api.find("cat")).toBe(4);
    expect(t.api.find("cat", { wholeWord: true })).toBe(3);
    expect(t.api.find("CAT", { matchCase: true })).toBe(0);
    const result = t.api.replaceAll("cat", "dog", { wholeWord: true });
    expect(result.total).toBe(3);
    expect(result.byStory).toEqual({ body: 2, header: 1 });
    // The header's own part carries the replacement.
    const saved = unzipSync(t.api.save());
    expect(strFromU8(saved["word/header1.xml"])).toContain("Header dog text");
    expect(strFromU8(saved["word/document.xml"])).toContain("concatenate");
    await t.unmount();
  });

  it("goes to a page and a bookmark", async () => {
    const t = await mount();
    expect(t.api.goToPage(1)).toBe(true);
    expect(t.api.goToPage(99)).toBe(false);
    expect(t.api.goToBookmark("mark")).toBe(true);
    expect(t.api.goToBookmark("missing")).toBe(false);
    await t.unmount();
  });
});
