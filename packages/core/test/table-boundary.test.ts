import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

describe("table boundary geometry", () => {
  it("places cell content and surrounding flow inside the outer border halves", () => {
    const table = `<w:tbl>
      <w:tblPr><w:tblBorders>
        <w:top w:val="single" w:sz="6" w:color="000000"/>
        <w:bottom w:val="single" w:sz="6" w:color="000000"/>
        <w:left w:val="single" w:sz="6" w:color="000000"/>
        <w:right w:val="single" w:sz="6" w:color="000000"/>
      </w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="6000"/></w:tcPr>${p("cell")}</w:tc></w:tr>
    </w:tbl>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(p("lead") + table + p("after")) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const items = result.pages[0].items;
    const text = (value: string) => {
      const item = items.find((candidate) => candidate.kind === "text" && candidate.text === value);
      if (item?.kind !== "text") throw new Error(`missing ${value}`);
      return item;
    };
    const rules = items
      .filter((item) => item.kind === "edge" && item.y1 === item.y2)
      .sort((a, b) => (a.kind === "edge" && b.kind === "edge" ? a.y1 - b.y1 : 0));
    if (rules[0]?.kind !== "edge" || rules[1]?.kind !== "edge") throw new Error("missing table rules");

    const lead = text("lead");
    const cell = text("cell");
    const after = text("after");
    const halfRule = rules[0].border.width / 2;

    expect(rules[0].y1 - (lead.lineTop + lead.lineHeight)).toBeCloseTo(halfRule, 5);
    expect(cell.lineTop - rules[0].y1).toBeCloseTo(halfRule, 5);
    expect(rules[1].y1 - rules[0].y1).toBeCloseTo(cell.lineHeight + rules[0].border.width, 5);
    expect(after.lineTop - rules[1].y1).toBeCloseTo(halfRule, 5);
  });

  it("moves a cantSplit row when its line box crosses the body bottom", () => {
    const lead = `<w:p><w:pPr><w:spacing w:line="1320" w:lineRule="exact"/></w:pPr>` +
      `<w:r><w:t>lead</w:t></w:r></w:p>`;
    const table = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
      <w:tr><w:trPr><w:cantSplit/></w:trPr>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="6000"/></w:tcPr>${p("cell")}</w:tc>
      </w:tr>
    </w:tbl>`;
    const section = `<w:sectPr><w:pgSz w:w="12240" w:h="3000"/>` +
      `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>`;
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(lead + table + section) }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });

    expect(result.totalPages).toBe(2);
    expect(result.pages[0].items.some((item) => item.kind === "text" && item.text === "cell")).toBe(false);
    expect(result.pages[1].items.some((item) => item.kind === "text" && item.text === "cell")).toBe(true);
  });
});
