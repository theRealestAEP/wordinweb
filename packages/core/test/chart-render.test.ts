// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { insertChartAt } from "../src/edit/charts.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { ChartData, Paragraph, Run, TextContent } from "../src/model.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

function renderChart(data: ChartData): SVGSVGElement {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("Anchor")) }));
  const paragraph = doc.sections[0].blocks[0] as Paragraph;
  const run = paragraph.children[0] as Run;
  const text = run.content[0] as TextContent;
  insertChartAt(doc, text.srcT!, data);

  const container = document.createElement("div");
  renderToDom(doc, layoutDocument(doc, { measurer: new ApproxMeasurer() }), container);
  const chart = container.querySelector<SVGSVGElement>('svg[data-dxw-chart="1"]');
  if (!chart) throw new Error("chart missing");
  return chart;
}

describe("native chart DOM rendering", () => {
  it("matches Word's generated column-chart scale, palette, axes, and ticks", () => {
    const chart = renderChart({
      type: "column",
      title: "Architecture Decision Profile",
      categories: ["Control", "Delivery speed", "Operating effort"],
      series: [
        { name: "Managed platform", values: [8, 9, 4] },
        { name: "Custom stack", values: [10, 6, 9] },
      ],
    });

    expect(Array.from(chart.querySelectorAll("text")).map((node) => node.textContent)).toContain("12");
    expect(chart.querySelector('rect[fill="#156082"]')).toBeTruthy();
    expect(chart.querySelector('rect[fill="#e97132"]')).toBeTruthy();
    expect(chart.querySelectorAll('[data-dxw-chart-tick="major"]')).toHaveLength(7);
    expect(chart.querySelectorAll('[data-dxw-chart-tick="minor"]')).toHaveLength(24);
    expect(chart.querySelectorAll('[data-dxw-chart-tick="category"]')).toHaveLength(4);
    expect(chart.querySelector("text")?.getAttribute("font-size")).toBe("24");
    expect(Number(chart.querySelector('rect[fill="#156082"]')?.getAttribute("width"))).toBeCloseTo(28.84, 1);
  });

  it("paints a visible marker at each line-chart value", () => {
    const chart = renderChart({
      type: "line",
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "Revenue", values: [2, 4, 3] },
        { name: "Cost", values: [1, 2, 2] },
      ],
    });

    expect(chart.querySelectorAll('[data-dxw-chart-marker="1"]')).toHaveLength(6);
  });
});
