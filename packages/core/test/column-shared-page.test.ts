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

  it("keeps a heading and its keepLines paragraph together past equally short shared columns", () => {
    const body = [
      // The intro leaves 54pt of the 648pt body for each shared column.
      p("Intro", `<w:pPr><w:spacing w:line="11880" w:lineRule="exact"/></w:pPr>`),
      `<w:p><w:pPr><w:sectPr>${PAGE}</w:sectPr></w:pPr></w:p>`,
      p("Heading", `<w:pPr><w:keepNext/><w:spacing w:line="360" w:lineRule="exact"/></w:pPr>`),
      p(["Body1", "Body2", "Body3", "Body4"].join("</w:t><w:br/><w:t>"),
        `<w:pPr><w:keepLines/><w:spacing w:line="360" w:lineRule="exact"/></w:pPr>`),
      p("Following", `<w:pPr><w:spacing w:line="3600" w:lineRule="exact"/></w:pPr>`),
      `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>`,
      p("Closing"),
      `<w:sectPr><w:type w:val="continuous"/>${PAGE}</w:sectPr>`,
    ];
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body.join("")) }));
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const positions = result.pages.flatMap((page, index) => page.items.flatMap((item) =>
      item.kind === "text" && /^(Heading|Body[1-4])$/.test(item.text)
        ? [{ text: item.text, page: index + 1, x: item.x }]
        : []));
    expect(result.totalPages).toBe(2);
    expect(positions.map(({ text, page }) => ({ text, page }))).toEqual([
      { text: "Heading", page: 2 },
      { text: "Body1", page: 2 },
      { text: "Body2", page: 2 },
      { text: "Body3", page: 2 },
      { text: "Body4", page: 2 },
    ]);
    expect(new Set(positions.map(({ x }) => x)).size).toBe(1);
  });

  it("takes the anchors to the next page when neither shared column holds their paragraph", () => {
    const body = [
      p("Intro", `<w:pPr><w:spacing w:line="10500" w:lineRule="exact"/></w:pPr>`),
      `<w:p><w:pPr><w:sectPr>${PAGE}</w:sectPr></w:pPr></w:p>`,
      `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>`,
      p("Closing"),
      `<w:sectPr><w:type w:val="continuous"/>${PAGE}</w:sectPr>`,
    ];
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body.join("")) }));
    // A wrapping box covers the column and leaves the anchor line below it.
    // Its 187px height exceeds the 164px left in each shared column.
    doc.sections[1].blocks.unshift({
      type: "paragraph", props: {}, children: [{
        type: "run", props: {}, content: [
          { kind: "anchor", shape: {
            type: "textbox", x: 0, y: 0, width: 624, height: 187,
            hRel: "column", vRel: "text", wrap: "square",
            blocks: [{ type: "paragraph", props: {}, children: [{
              type: "run", props: {}, content: [{ kind: "text", text: "Survey box" }],
            }] }],
          } },
          { kind: "text", text: "Anchor line" },
        ],
      }],
    });
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const texts = result.pages.map((page) => page.items
      .map((item) => item.kind === "text" ? item.text : "").join(" ").replace(/\s+/g, " "));
    expect(result.totalPages).toBe(2);
    expect(texts[0]).toContain("Intro");
    expect(texts[0]).not.toContain("Survey box");
    expect(texts[1]).toContain("Survey box");
    expect(texts[1]).toContain("Anchor line");
    expect(texts[1]).toContain("Closing");
  });
});
