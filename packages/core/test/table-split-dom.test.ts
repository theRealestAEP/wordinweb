// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** The rendered pages of a document in the editor's DOM. */
function domPages(bodyXml: string): HTMLElement[] {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const container = document.createElement("div");
  document.body.appendChild(container);
  renderToDom(doc, result, container);
  return [...container.querySelectorAll<HTMLElement>(".dxw-page")];
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
    const pages = domPages(p("Before") + table + p("After")).map((el) => el.textContent ?? "");
    expect(pages.length).toBeGreaterThanOrEqual(3);
    const all = pages.join("");
    for (let i = 1; i <= 120; i++) expect(all).toContain(`row${num(i)}`);
    expect(pages[pages.length - 1]).toContain("row120");
    expect(pages[pages.length - 1]).toContain("After");
  });
});

describe("DOM renderer: vertical text in a merged cell", () => {
  it("lays the text along the merged height, not the first row", () => {
    for (const dir of ["btLr", "tbRl"]) {
      const table =
        `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="1400"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="240"/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/><w:vMerge w:val="restart"/><w:textDirection w:val="${dir}"/></w:tcPr>` +
        `${p("Pickup practices conducted")}</w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/></w:tcPr>${p("x")}</w:tc></w:tr>` +
        `<w:tr><w:trPr><w:trHeight w:val="4000"/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/><w:vMerge/><w:textDirection w:val="${dir}"/></w:tcPr><w:p/></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/></w:tcPr>${p("y")}</w:tc></w:tr></w:tbl>`;
      const [page] = domPages(table + p(""));
      const rotated = [...page.querySelectorAll<HTMLElement>("*")].filter(
        (el) => el.style.transform.includes("rotate(") && el.textContent,
      );
      expect(rotated.map((el) => el.textContent).join("").replace(/\s/g, "")).toBe("Pickuppracticesconducted");
      // One rotated line holds the whole heading; the old layout cut it into
      // "Pi", "ck", "up", ... on lines of their own.
      expect(new Set(rotated.map((el) => el.style.top)).size).toBe(1);
    }
  });
});
