// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { buildPrintHtml, printStylesheet, renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

describe("print page sizes", () => {
  // A landscape section printed at the first page's portrait size: one
  // @page rule covered every sheet. Each distinct page size now gets a named
  // @page, and each page takes the one matching its own size.
  const portrait = `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>`;
  const landscape = `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>`;
  const body =
    `<w:p><w:pPr><w:sectPr>${portrait}</w:sectPr></w:pPr><w:r><w:t>Portrait page</w:t></w:r></w:p>` +
    p("Landscape page") +
    `<w:sectPr>${landscape}</w:sectPr>`;

  function render(zoom: number): HTMLElement {
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
    const layout = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const container = document.createElement("div");
    renderToDom(doc, layout, container, { zoom, pageGap: 0, pageShadow: false, comments: false });
    return container;
  }

  it("gives each page size its own named @page, at the unzoomed size", () => {
    const root = render(1.5);
    const pages = [...root.querySelectorAll<HTMLElement>(".dxw-page")];
    expect(pages.map((el) => `${el.dataset.pageWidth}x${el.dataset.pageHeight}`)).toEqual(["816x1056", "1056x816"]);
    const css = printStylesheet(816, 1056, root);
    expect(css).toContain("@page dxw-816x1056 { size: 8.5in 11in; margin: 0; }");
    expect(css).toContain("@page dxw-1056x816 { size: 11in 8.5in; margin: 0; }");
    expect(css).toContain(`.dxw-page[data-page-width="1056"][data-page-height="816"] { page: dxw-1056x816; }`);
  });

  it("carries the named pages into the print document", () => {
    const html = buildPrintHtml(render(1), 816, 1056);
    expect(html).toContain("@page dxw-1056x816 { size: 11in 8.5in; margin: 0; }");
    // The plain @page stays as the fallback for a page without a size stamp.
    expect(html).toContain("@page { size: 8.5in 11in; margin: 0; }");
  });
});
