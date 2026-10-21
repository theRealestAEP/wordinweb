// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const cell = (text: string, width: number, props = "") =>
  `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${width}"/><w:shd w:fill="EEEEEE"/>${props}</w:tcPr>${p(text)}</w:tc>`;
const row = (cells: string, props = "") => `<w:tr><w:trPr>${props}</w:trPr>${cells}</w:tr>`;

function layout(grid: number[], rows: string, layout = "fixed", width?: number) {
  const table = `<w:tbl><w:tblPr><w:tblLayout w:type="${layout}"/>
    ${width === undefined ? "" : `<w:tblW w:type="dxa" w:w="${width}"/>`}
    <w:tblBorders>${["top", "bottom", "left", "right", "insideH", "insideV"].map(
      (side) => `<w:${side} w:val="single" w:sz="6"/>`,
    ).join("")}</w:tblBorders></w:tblPr>
    <w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>${rows}</w:tbl>`;
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(table) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const items = result.pages.flatMap((page) => page.items);
  const fills = items.filter((item) => item.kind === "rect").filter((item) => item.role === "table-fill");
  const origin = items.find((item) => item.kind === "grip" && item.axis === "move")!;
  if (origin.kind !== "grip") throw new Error("missing table origin");
  return { doc, result, items, fills, x: origin.x };
}

describe("omitted table grid columns", () => {
  it.each(["fixed", "autofit"])("starts a %s row after gridBefore and keeps wide text on one line", (mode) => {
    // The narrow first column caused character wrapping when the cell used it.
    const { doc, result, items, fills, x } = layout([150, 4500], row(cell("Readable", 4500),
      '<w:gridBefore w:val="1"/><w:wBefore w:type="dxa" w:w="300"/>'), mode);
    expect(fills).toHaveLength(1);
    expect(fills[0].x).toBeCloseTo(x + 10);
    expect(fills[0].width).toBeCloseTo(300);
    const text = items.filter((item) => item.kind === "text").filter((item) => item.text.trim());
    expect(text).toHaveLength(1);
    expect(text[0]).toMatchObject({ text: "Readable" });
    expect(text[0].x).toBeCloseTo(x + 11);
    for (const edge of items.filter((item) => item.kind === "edge")) {
      expect(Math.min(edge.x1, edge.x2)).toBeGreaterThanOrEqual(x + 10);
    }

    const container = document.createElement("div");
    const rendered = renderToDom(doc, result, container);
    expect(container.textContent).toContain("Readable");
    const fill = [...container.querySelectorAll<HTMLElement>("div")].find(
      (el) => el.style.background === "rgb(238, 238, 238)",
    )!;
    expect(parseFloat(fill.style.left)).toBeCloseTo(x + 10);
    expect(parseFloat(fill.style.width)).toBeCloseTo(300);
    rendered.destroy();
  });

  it("starts a spanning cell after two omitted columns", () => {
    const { fills, x } = layout([150, 300, 3000, 1500, 600], row(
      cell("Spanning", 4500, '<w:gridSpan w:val="2"/>') + cell("Tail", 600),
      '<w:gridBefore w:val="2"/><w:wBefore w:type="dxa" w:w="450"/>',
    ));
    expect(fills.map((fill) => [fill.x - x, fill.width])).toEqual([[30, 300], [330, 40]]);
  });

  it("reserves gridAfter columns and paints through the last cell edge", () => {
    const { fills, items, x } = layout([3000, 1500], row(cell("End", 3000),
      '<w:gridAfter w:val="1"/><w:wAfter w:type="dxa" w:w="3000"/>'));
    expect(fills.map((fill) => [fill.x - x, fill.width])).toEqual([[0, 200]]);
    const grip = items.find((item) => item.kind === "grip" && item.axis === "col")!;
    expect(grip.kind === "grip" && grip.renderedWidths).toEqual([200, 100]);
    for (const edge of items.filter((item) => item.kind === "edge")) {
      expect(Math.max(edge.x1, edge.x2)).toBeLessThanOrEqual(x + 200);
    }
  });

  it("uses wBefore and wAfter when the grid is absent", () => {
    const { fills, x } = layout([], row(cell("Fallback", 3000),
      '<w:gridBefore w:val="2"/><w:wBefore w:type="dxa" w:w="900"/>' +
      '<w:gridAfter w:val="1"/><w:wAfter w:type="dxa" w:w="600"/>'), "fixed", 4500);
    expect(fills.map((fill) => [fill.x - x, fill.width])).toEqual([[60, 200]]);
  });

  it("uses wAfter beyond a partial grid", () => {
    const { fills, items, x } = layout([3000], row(cell("Partial", 3000),
      '<w:gridAfter w:val="2"/><w:wAfter w:type="dxa" w:w="1500"/>'), "fixed", 4500);
    expect(fills.map((fill) => [fill.x - x, fill.width])).toEqual([[0, 200]]);
    const grip = items.find((item) => item.kind === "grip" && item.axis === "col")!;
    expect(grip.kind === "grip" && grip.renderedWidths).toEqual([200, 50, 50]);
  });

  it("uses wBefore for the missing part of a leading grid gap", () => {
    const { fills, x } = layout([300], row(cell("Partial", 3000),
      '<w:gridBefore w:val="2"/><w:wBefore w:type="dxa" w:w="900"/>'), "fixed", 3900);
    expect(fills.map((fill) => [fill.x - x, fill.width])).toEqual([[60, 200]]);
  });

  it.each([false, true])("keeps page-break borders inside the cells (split row: %s)", (split) => {
    const trPr = '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>';
    const rows = split
      ? row(`<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>${p("Content").repeat(20)}</w:tc>`, trPr)
      : Array.from({ length: 20 }, () => row(cell("Content", 3000), trPr)).join("");
    const table = `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/><w:gridCol w:w="1500"/></w:tblGrid>${rows}</w:tbl>`;
    const section = '<w:sectPr><w:pgSz w:w="12240" w:h="3600"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>';
    const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblBorders>
      ${["top", "bottom", "left", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="single" w:sz="6"/>`).join("")}
      </w:tblBorders></w:tblPr></w:style></w:styles>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(table + section), "word/styles.xml": styles }));
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    expect(result.totalPages).toBeGreaterThan(1);
    for (const page of result.pages) {
      const edges = page.items.filter((item) => item.kind === "edge").filter((item) => item.role === "table-rule");
      expect(edges.length).toBeGreaterThan(0);
      for (const edge of edges) {
        expect(Math.min(edge.x1, edge.x2)).toBeGreaterThanOrEqual(148);
        expect(Math.max(edge.x1, edge.x2)).toBeLessThanOrEqual(348);
      }
    }
  });

  it("matches vertical merges by grid column across different gridBefore values", () => {
    const { fills, x } = layout([1500, 3000, 1500],
      row(cell("Merged", 3000, '<w:vMerge w:val="restart"/>') + cell("Top", 1500), '<w:gridBefore w:val="1"/>') +
      row(cell("Left", 1500) + cell("", 3000, '<w:vMerge/>') + cell("Bottom", 1500)) +
      row(cell("Next", 3000) + cell("Last", 1500), '<w:gridBefore w:val="1"/>'),
    );
    const [merged, top, left, bottom, next] = fills;
    expect(merged.x).toBeCloseTo(x + 100);
    expect(merged.width).toBeCloseTo(200);
    expect(left.x).toBeCloseTo(x);
    expect(bottom.x).toBeCloseTo(x + 300);
    expect(merged.height).toBeCloseTo(top.height + bottom.height);
    expect(merged.y + merged.height).toBeCloseTo(next.y);
  });
});
