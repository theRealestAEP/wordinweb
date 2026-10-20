// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** The text of each rendered page in the editor's DOM. */
function domPageTexts(bodyXml: string): string[] {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const container = document.createElement("div");
  document.body.appendChild(container);
  renderToDom(doc, result, container);
  return [...container.querySelectorAll<HTMLElement>(".dxw-page")].map((el) => el.textContent ?? "");
}

describe("DOM renderer: tables that break across pages", () => {
  it("shows every row of a floating table taller than the page", () => {
    const num = (i: number) => String(i).padStart(3, "0");
    const rows = Array.from(
      { length: 120 },
      (_, i) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${p(`row${num(i + 1)}`)}</w:tc></w:tr>`,
    ).join("");
    const table =
      `<w:tbl><w:tblPr><w:tblpPr w:horzAnchor="margin" w:vertAnchor="text" w:tblpY="300"/>` +
      `<w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
      `${rows}</w:tbl>`;
    const pages = domPageTexts(p("Before") + table + p("After"));
    expect(pages.length).toBeGreaterThanOrEqual(3);
    const all = pages.join("");
    for (let i = 1; i <= 120; i++) expect(all).toContain(`row${num(i)}`);
    expect(pages[pages.length - 1]).toContain("row120");
    expect(pages[pages.length - 1]).toContain("After");
  });
});
