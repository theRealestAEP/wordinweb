import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

// A continuous break into a 2-column section that balances before the next
// continuous break. word95-check-a R4 (govdocs 133189, 143726): the section
// starts under the 1-column text on the shared page, as Word and LibreOffice
// lay it, even when it runs on past that page.

const PAGE = `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>`;

function firstPageOf(space: number, paragraphs: number): { pages: number; firstPage: string } {
  const body = [
    p("Intro"),
    `<w:p><w:pPr><w:sectPr>${PAGE}</w:sectPr></w:pPr></w:p>`,
    ...Array.from({ length: paragraphs }, (_, i) => p(`Column ${i + 1} text runs on for a while here.`)),
    `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="${space}"/></w:sectPr></w:pPr></w:p>`,
    p("Closing"),
    `<w:sectPr><w:type w:val="continuous"/>${PAGE}</w:sectPr>`,
  ];
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body.join("")) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const firstPage = result.pages[0].items.map((item) => (item.kind === "text" ? item.text : "")).join(" ");
  return { pages: result.totalPages, firstPage };
}

describe("a 2-column section after a continuous break", () => {
  it("starts on the shared page when it is too long for the rest of it", () => {
    const { pages, firstPage } = firstPageOf(720, 200);
    expect(pages).toBeGreaterThan(1);
    expect(firstPage).toContain("Intro");
    expect(firstPage).toMatch(/Column\s+1\s/);
  });

  it("moves to a fresh page when its columns are too narrow to hold a glyph (wild-multicolumn)", () => {
    const { firstPage } = firstPageOf(360000, 6);
    expect(firstPage).toContain("Intro");
    expect(firstPage).not.toMatch(/C\s+o\s+l/);
  });
});
