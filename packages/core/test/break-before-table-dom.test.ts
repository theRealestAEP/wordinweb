// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { W_NS, makeDocx, p, wrapDocument } from "./helpers.js";

/**
 * The browser view draws the same layout as the PDF. These cases check, in
 * the DOM that renderToDom builds, the page-break fixes that
 * break-before-table.test.ts checks in the layout.
 */

/** Each page's spans, as text with their top position in px. */
function render(bodyXml: string, parts: Record<string, string> = {}): { text: string; top: number }[][] {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml), ...parts }));
  const container = document.createElement("div");
  document.body.appendChild(container);
  renderToDom(doc, layoutDocument(doc, { measurer: new ApproxMeasurer() }), container);
  return [...container.querySelectorAll<HTMLElement>(".dxw-page")].map((page) =>
    [...page.querySelectorAll<HTMLElement>("span")]
      .filter((el) => el.children.length === 0 && (el.textContent ?? "").trim() !== "")
      .map((el) => ({ text: el.textContent!.trim(), top: parseFloat(el.style.top) })),
  );
}

const pageOf = (pages: { text: string }[][], text: string) => pages.findIndex((pg) => pg.some((s) => s.text === text)) + 1;
const topOf = (pages: { text: string; top: number }[][], text: string) =>
  pages.flat().find((s) => s.text === text)!.top;

describe("page breaks in the DOM view", () => {
  it("draws the paragraph after a table that follows a page break below the table", () => {
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:tc>${p("A")}</w:tc><w:tc>${p("one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen seventeen eighteen")}</w:tc></w:tr></w:tbl>`;
    const pages = render(p("Before") + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + table + p("After"));
    expect([pageOf(pages, "Before"), pageOf(pages, "A"), pageOf(pages, "eighteen"), pageOf(pages, "After")]).toEqual([1, 2, 2, 2]);
    expect(topOf(pages, "After")).toBeGreaterThan(topOf(pages, "eighteen"));
  });

  it("draws a numbered heading's label beside its text after a leading page break", () => {
    const numbering =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W_NS}>` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/>' +
      '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
    const pages = render(
      p("AAA") +
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
        '<w:r><w:br w:type="page"/></w:r><w:r><w:t>CITE</w:t></w:r></w:p>',
      { "word/numbering.xml": numbering },
    );
    expect(pages.length).toBe(2);
    expect([pageOf(pages, "I."), pageOf(pages, "CITE")]).toEqual([2, 2]);
    expect(topOf(pages, "I.")).toBeCloseTo(topOf(pages, "CITE"), 1);
  });
});
