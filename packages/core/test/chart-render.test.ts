// @vitest-environment jsdom
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { insertChartAt } from "../src/edit/charts.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { ChartData, Paragraph, Run, TextContent } from "../src/model.js";
import { parseChartPart } from "../src/parse/chart.js";
import { renderToDom } from "../src/render/dom.js";
import { parseXml } from "../src/xml.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** A document whose only chart is the one insertChart just authored. */
function documentWithChart(data: ChartData): DocxDocument {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("Anchor")) }));
  const paragraph = doc.sections[0].blocks[0] as Paragraph;
  const run = paragraph.children[0] as Run;
  const text = run.content[0] as TextContent;
  insertChartAt(doc, text.srcT!, data);
  return doc;
}

function renderDoc(doc: DocxDocument): SVGSVGElement {
  const container = document.createElement("div");
  renderToDom(doc, layoutDocument(doc, { measurer: new ApproxMeasurer() }), container);
  const chart = container.querySelector<SVGSVGElement>('svg[data-dxw-chart="1"]');
  if (!chart) throw new Error("chart missing");
  return chart;
}

function renderChart(data: ChartData): SVGSVGElement {
  return renderDoc(documentWithChart(data));
}

/**
 * Render a hand-written c:chartSpace the way a Word-authored part reaches the
 * renderer. The part is parsed and its model replaces the one in the laid-out
 * page, so everything downstream of parsing is exercised on Word's own XML.
 */
function renderChartXml(xml: string): SVGSVGElement {
  const doc = documentWithChart({ type: "column", categories: ["a"], series: [{ name: "s", values: [1] }] });
  const parsed = parseChartPart(parseXml(xml), doc.theme);
  if (!parsed) throw new Error("chart XML did not parse");
  const container = document.createElement("div");
  const layout = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  for (const page of layout.pages) {
    for (const item of page.items) if (item.kind === "chart") item.data = parsed;
  }
  renderToDom(doc, layout, container);
  const chart = container.querySelector<SVGSVGElement>('svg[data-dxw-chart="1"]');
  if (!chart) throw new Error("chart missing");
  return chart;
}

const C_NS = `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
  ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;

const CAT_AXIS = `<c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>`;
const VAL_AXIS = `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>`;

const texts = (chart: SVGSVGElement): string[] =>
  Array.from(chart.querySelectorAll("text")).map((node) => node.textContent ?? "");

describe("native chart DOM rendering", () => {
  it("paints a column chart's scale, palette, ticks and legend", () => {
    const chart = renderChart({
      type: "column",
      title: "Architecture Decision Profile",
      categories: ["Control", "Delivery speed", "Operating effort"],
      series: [
        { name: "Managed platform", values: [8, 9, 4] },
        { name: "Custom stack", values: [10, 6, 9] },
      ],
    });

    // The data tops out at 10, so the scale runs 0..12 in steps of 2.
    expect(texts(chart)).toEqual(expect.arrayContaining(["0", "2", "4", "6", "8", "10", "12"]));
    expect(texts(chart)).toContain("Architecture Decision Profile");
    expect(texts(chart)).toEqual(expect.arrayContaining(["Managed platform", "Custom stack"]));
    // A document with no theme part falls back to the Office accent colours.
    expect(chart.querySelectorAll('[data-dxw-chart-bar][fill="#156082"]')).toHaveLength(3);
    expect(chart.querySelectorAll('[data-dxw-chart-bar][fill="#e97132"]')).toHaveLength(3);
    // insertChart writes majorTickMark="none" on the value axis and "out" on
    // the category axis, which ticks the three category boundaries plus the end.
    expect(chart.querySelectorAll('[data-dxw-chart-tick="major"]')).toHaveLength(0);
    expect(chart.querySelectorAll('[data-dxw-chart-tick="category"]')).toHaveLength(4);
    expect(chart.querySelector("text")?.getAttribute("font-size")).toBe(String(18 * (96 / 72)));
  });

  it("places clustered bars by the gap width and overlap in the file", () => {
    const chart = renderChart({
      type: "column",
      categories: ["A", "B"],
      series: [{ name: "One", values: [1, 2] }, { name: "Two", values: [2, 1] }],
    });
    const bars = Array.from(chart.querySelectorAll("[data-dxw-chart-bar]"));
    expect(bars).toHaveLength(4);
    expect(new Set(bars.map((bar) => bar.getAttribute("width"))).size).toBe(1);
    // Two series at gapWidth 150 with no overlap means 3.5 bar widths a band.
    const width = Number(bars[0].getAttribute("width"));
    expect(Number(bars[1].getAttribute("x")) - Number(bars[0].getAttribute("x"))).toBeCloseTo(width * 3.5, 6);
  });

  it("draws a bar chart horizontally from the value-zero baseline", () => {
    const chart = renderChart({
      type: "bar",
      categories: ["A", "B"],
      series: [{ name: "One", values: [3, 6] }],
    });
    const bars = Array.from(chart.querySelectorAll("[data-dxw-chart-bar]"));
    expect(bars).toHaveLength(2);
    expect(bars[0].getAttribute("x")).toBe(bars[1].getAttribute("x"));
    expect(Number(bars[1].getAttribute("width"))).toBeCloseTo(Number(bars[0].getAttribute("width")) * 2, 6);
  });

  it("runs a bar chart's categories up from the bottom, as Word orders them", () => {
    const chart = renderChart({
      type: "bar",
      categories: ["First", "Second", "Third"],
      series: [{ name: "One", values: [1, 2, 3] }],
    });
    const bars = Array.from(chart.querySelectorAll("[data-dxw-chart-bar]"));
    // The first category sits lowest on the page, so its y is the largest.
    const ys = bars.map((bar) => Number(bar.getAttribute("y")));
    expect(ys[0]).toBeGreaterThan(ys[1]);
    expect(ys[1]).toBeGreaterThan(ys[2]);
  });

  it("reverses a bar chart's legend to follow its reversed category axis", () => {
    // Word listed Beta above Alpha on the horizontal bar chart in
    // probe-charts-basic.docx; the column chart beside it kept Alpha first.
    const legend = (type: "bar" | "column"): string[] => {
      const chart = renderChart({
        type,
        legend: "b",
        categories: ["Q1", "Q2"],
        series: [
          { name: "Alpha", values: [4.2, 7.8] },
          { name: "Beta", values: [6.5, 2.4] },
        ],
      });
      return texts(chart).filter((text) => text === "Alpha" || text === "Beta");
    };
    expect(legend("bar")).toEqual(["Beta", "Alpha"]);
    expect(legend("column")).toEqual(["Alpha", "Beta"]);
  });

  it("sizes chart text at Word's defaults: an 18pt title over 10pt everywhere else", () => {
    const chart = renderChart({
      type: "column",
      title: "Sizes",
      legend: "b",
      categories: ["Q1"],
      series: [{ name: "Alpha", values: [4.2] }],
    });
    const sizes = Array.from(chart.querySelectorAll("text"))
      .map((node) => node.getAttribute("font-size"));
    expect(sizes[0]).toBe(String(18 * (96 / 72)));
    expect(new Set(sizes.slice(1))).toEqual(new Set([String(10 * (96 / 72))]));
  });

  it("seats a line's end points on the plot edges and a column's inside its band", () => {
    const categories = ["A", "B", "C"];
    const series = [{ name: "One", values: [1, 2, 3] }];
    const line = renderChart({ type: "line", categories, series });
    const column = renderChart({ type: "column", categories, series });

    // c:crossBetween is midCat for a line, so the first point is on the axis.
    const points = (line.querySelector("path")?.getAttribute("d") ?? "")
      .match(/-?\d+\.\d+/g)!.map(Number);
    const columnBars = Array.from(column.querySelectorAll("[data-dxw-chart-bar]"));
    const firstBar = Number(columnBars[0].getAttribute("x"));
    // The line starts left of where the first column bar does, because the
    // column's band insets it by half a category.
    expect(points[0]).toBeLessThan(firstBar);
  });

  it("paints a marker at each line value and breaks the line over a blank", () => {
    const chart = renderChartXml(
      `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:layout/>` +
      `<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/>` +
      `<c:cat><c:strRef><c:strCache><c:ptCount val="4"/>` +
      `<c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt>` +
      `<c:pt idx="2"><c:v>Q3</c:v></c:pt><c:pt idx="3"><c:v>Q4</c:v></c:pt>` +
      `</c:strCache></c:strRef></c:cat>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="4"/>` +
      `<c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt>` +
      `<c:pt idx="3"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>` +
      `<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>` +
      CAT_AXIS + VAL_AXIS + `</c:plotArea></c:chart></c:chartSpace>`,
    );
    expect(chart.querySelectorAll('[data-dxw-chart-marker="1"]')).toHaveLength(3);
    // Two subpaths: the lone first point, then the run over Q3 and Q4.
    expect((chart.querySelector("path")?.getAttribute("d") ?? "").match(/M /g)).toHaveLength(2);
  });

  it("paints pie wedges in proportion and labels them when the file asks", () => {
    const chart = renderChartXml(
      `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:layout/>` +
      `<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/>` +
      `<c:cat><c:strRef><c:strCache><c:ptCount val="2"/>` +
      `<c:pt idx="0"><c:v>Yes</c:v></c:pt><c:pt idx="1"><c:v>No</c:v></c:pt>` +
      `</c:strCache></c:strRef></c:cat>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="2"/>` +
      `<c:pt idx="0"><c:v>75</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt>` +
      `</c:numCache></c:numRef></c:val></c:ser>` +
      `<c:dLbls><c:showVal val="1"/></c:dLbls></c:pieChart></c:plotArea>` +
      `<c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`,
    );
    const wedges = Array.from(chart.querySelectorAll("path"));
    expect(wedges).toHaveLength(2);
    // A 75% sweep is more than half a turn; 25% is less.
    expect(wedges[0].getAttribute("d")).toContain(" 0 1 1 ");
    expect(wedges[1].getAttribute("d")).toContain(" 0 0 1 ");
    expect(chart.querySelectorAll('[data-dxw-chart-label="1"]')).toHaveLength(2);
    // A varied single series legends its points, not its one series name.
    expect(texts(chart)).toEqual(expect.arrayContaining(["75", "25", "Yes", "No"]));
  });

  it("cuts a hole in a doughnut", () => {
    const chart = renderChart({
      type: "doughnut",
      categories: ["A", "B"],
      series: [{ name: "Share", values: [1, 1] }],
    });
    const wedges = Array.from(chart.querySelectorAll("path"));
    expect(wedges).toHaveLength(2);
    // A ring segment traces the outer arc, steps in, and traces back.
    expect(wedges[0].getAttribute("d")).toMatch(/^M .* A .* L .* A .* Z$/);
  });

  it("plots a scatter series against its own x values", () => {
    const chart = renderChart({
      type: "scatter",
      categories: ["1", "2", "4"],
      series: [{ name: "Trial", values: [10, 14, 9] }],
    });
    const xs = Array.from(chart.querySelectorAll('[data-dxw-chart-marker="1"]'))
      .map((marker) => Number(marker.getAttribute("cx")));
    expect(xs).toHaveLength(3);
    // x = 1, 2, 4 puts the third point twice as far from the second as the
    // second is from the first.
    expect(xs[2] - xs[1]).toBeCloseTo((xs[1] - xs[0]) * 2, 6);
  });

  it("fills an area chart down to the baseline", () => {
    const chart = renderChart({
      type: "area",
      categories: ["A", "B", "C"],
      series: [{ name: "Load", values: [2, 5, 3] }],
    });
    const area = chart.querySelector("path");
    expect(area?.getAttribute("d")).toMatch(/Z$/);
    expect(area?.getAttribute("fill")).toBe("#156082");
  });

  it("stacks a stacked column and pins a percent-stacked axis at 100%", () => {
    const stackedXml = (grouping: string, format: string) =>
      `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:layout/>` +
      `<c:barChart><c:barDir val="col"/><c:grouping val="${grouping}"/>` +
      ([["A", 1], ["B", 3]] as Array<[string, number]>).map(([name, value], index) =>
        `<c:ser><c:idx val="${index}"/>` +
        `<c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt>` +
        `</c:strCache></c:strRef></c:tx>` +
        `<c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${value}</c:v></c:pt>` +
        `</c:numCache></c:numRef></c:val></c:ser>`).join("") +
      `<c:gapWidth val="150"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      CAT_AXIS +
      `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/>` +
      `<c:numFmt formatCode="${format}" sourceLinked="0"/><c:crossAx val="1"/></c:valAx>` +
      `</c:plotArea></c:chart></c:chartSpace>`;

    const stacked = renderChartXml(stackedXml("stacked", "General"));
    const bars = Array.from(stacked.querySelectorAll("[data-dxw-chart-bar]"));
    expect(bars).toHaveLength(2);
    // One slot: same x and width, with the second sitting on top of the first.
    expect(bars[0].getAttribute("x")).toBe(bars[1].getAttribute("x"));
    expect(bars[0].getAttribute("width")).toBe(bars[1].getAttribute("width"));
    expect(Number(bars[1].getAttribute("y")) + Number(bars[1].getAttribute("height")))
      .toBeCloseTo(Number(bars[0].getAttribute("y")), 6);

    const percent = renderChartXml(stackedXml("percentStacked", "0%"));
    expect(texts(percent)).toEqual(expect.arrayContaining(["0%", "20%", "40%", "60%", "80%", "100%"]));
    expect(texts(percent)).not.toContain("120%");
  });

  it("formats axis labels through the file's number format", () => {
    const chart = renderChartXml(
      `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:layout/>` +
      `<c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="2"/>` +
      `<c:pt idx="0"><c:v>1200000</c:v></c:pt><c:pt idx="1"><c:v>3000000</c:v></c:pt>` +
      `</c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      CAT_AXIS +
      `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/>` +
      `<c:numFmt formatCode="$#,##0" sourceLinked="0"/><c:crossAx val="1"/></c:valAx>` +
      `</c:plotArea></c:chart></c:chartSpace>`,
    );
    // The data tops out at 3,000,000, so the scale runs 0..3,500,000 by 500,000.
    expect(texts(chart)).toEqual(expect.arrayContaining(["$0", "$500,000", "$3,500,000"]));
  });

  it("honours a deleted axis and a suppressed tick label", () => {
    const chart = renderChartXml(
      `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:layout/>` +
      `<c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>` +
      `<c:cat><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Only</c:v></c:pt>` +
      `</c:strCache></c:strRef></c:cat>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>4</c:v></c:pt>` +
      `</c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:catAx><c:axId val="1"/><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>` +
      `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/>` +
      `<c:tickLblPos val="none"/><c:crossAx val="1"/></c:valAx>` +
      `</c:plotArea></c:chart></c:chartSpace>`,
    );
    expect(texts(chart)).toEqual([]);
    // The chart background and the single bar, and nothing else.
    expect(chart.querySelectorAll("rect")).toHaveLength(2);
  });

  it("reserves the box and names the kind for a chart it cannot paint", () => {
    const chart = renderChartXml(
      `<c:chartSpace ${C_NS}><c:chart>` +
      `<c:title><c:tx><c:rich><a:p><a:r><a:t>Terrain</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:plotArea><c:layout/><c:surface3DChart><c:ser><c:idx val="0"/></c:ser>` +
      `</c:surface3DChart></c:plotArea></c:chart></c:chartSpace>`,
    );
    expect(texts(chart)).toEqual(["Terrain", "surface3DChart is not rendered"]);
    const box = chart.querySelectorAll("rect")[1];
    expect(box.getAttribute("stroke-dasharray")).toBe("4 3");
    expect(Number(box.getAttribute("width"))).toBeCloseTo(479, 0);
  });

  it("colours series from the document theme when it has one", () => {
    const doc = DocxDocument.load(makeDocx({
      "word/document.xml": wrapDocument(p("Anchor")),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"` +
        ` Target="theme/theme1.xml"/></Relationships>`,
      "word/theme/theme1.xml": `<?xml version="1.0"?>` +
        `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>` +
        `<a:clrScheme name="Custom"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="ffffff"/></a:lt1>` +
        `<a:accent1><a:srgbClr val="4472c4"/></a:accent1><a:accent2><a:srgbClr val="ed7d31"/></a:accent2>` +
        `<a:accent3><a:srgbClr val="a5a5a5"/></a:accent3><a:accent4><a:srgbClr val="ffc000"/></a:accent4>` +
        `<a:accent5><a:srgbClr val="5b9bd5"/></a:accent5><a:accent6><a:srgbClr val="70ad47"/></a:accent6>` +
        `</a:clrScheme><a:fontScheme name="Custom">` +
        `<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>` +
        `<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>` +
        `</a:themeElements></a:theme>`,
    }));
    const paragraph = doc.sections[0].blocks[0] as Paragraph;
    const text = (paragraph.children[0] as Run).content[0] as TextContent;
    insertChartAt(doc, text.srcT!, {
      type: "column",
      categories: ["A"],
      series: [{ name: "One", values: [1] }, { name: "Two", values: [2] }],
    });
    const chart = renderDoc(doc);
    expect(chart.querySelector('[data-dxw-chart-bar][fill="#4472c4"]')).toBeTruthy();
    expect(chart.querySelector('[data-dxw-chart-bar][fill="#ed7d31"]')).toBeTruthy();
  });

  it("varies a seventh series so it stays apart from the first", () => {
    const chart = renderChart({
      type: "column",
      categories: ["A"],
      series: Array.from({ length: 7 }, (_, index) => ({ name: `S${index}`, values: [index + 1] })),
    });
    const bars = Array.from(chart.querySelectorAll("[data-dxw-chart-bar]"))
      .map((bar) => bar.getAttribute("fill"));
    expect(new Set(bars).size).toBe(7);
    expect(bars[6]).not.toBe(bars[0]);
  });
});

describe("chart round trip", () => {
  const sample: ChartData = {
    type: "column",
    title: "Insert and reopen",
    categories: ["A", "B", "C"],
    series: [{ name: "One", values: [3, 1, 4] }, { name: "Two", values: [1, 5, 9] }],
  };

  const chartPartOf = (doc: DocxDocument): string =>
    Object.keys(doc.pkg.raw()).find((name) => name.startsWith("word/charts/chart"))!;

  it("renders every kind insertChart can author, read back from the part it wrote", () => {
    for (const type of ["column", "bar", "line", "pie", "doughnut", "area", "scatter"] as const) {
      const doc = documentWithChart({ ...sample, type });
      const parsed = parseChartPart(parseXml(strFromU8(doc.pkg.raw()[chartPartOf(doc)])), doc.theme);

      // The part reads back as the kind it was asked for, with its data intact.
      // Nothing here sees the ChartData that went in.
      expect(parsed, type).not.toBeNull();
      expect(parsed!.type, type).toBe(type);
      expect(parsed!.unsupported, type).toBeUndefined();
      expect(parsed!.title, type).toBe("Insert and reopen");
      expect(parsed!.series.map((series) => series.name), type).toEqual(["One", "Two"]);
      expect(parsed!.series[1].values, type).toEqual([1, 5, 9]);
      if (type !== "scatter") expect(parsed!.categories, type).toEqual(["A", "B", "C"]);

      const chart = renderDoc(documentWithChart({ ...sample, type }));
      // Every kind paints its own marks, not just the background rectangle.
      expect(chart.childNodes.length, type).toBeGreaterThan(3);
      expect(texts(chart), type).toContain("Insert and reopen");
    }
  });

  it("saves an untouched chart part and its workbook byte for byte", () => {
    const doc = documentWithChart(sample);
    const first = unzipSync(doc.save());
    const chartPart = Object.keys(first).find((name) => name.startsWith("word/charts/chart"))!;
    const workbookPart = Object.keys(first).find((name) => name.endsWith(".xlsx"))!;

    const second = unzipSync(DocxDocument.load(doc.save()).save());

    expect(second[chartPart]).toEqual(first[chartPart]);
    expect(second[workbookPart]).toEqual(first[workbookPart]);
  });

  it("keeps a chart part the engine did not author byte-identical across a save", () => {
    // Carries elements the model does not keep (c:extLst, c:roundedCorners), so
    // a save that rebuilt the part from the model would be visible here.
    const authored = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<c:chartSpace ${C_NS}><c:roundedCorners val="0"/>` +
      `<c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>` +
      `<c:ser><c:idx val="0"/><c:extLst><c:ext uri="{C3380CC4}"/></c:extLst>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>7</c:v></c:pt>` +
      `</c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      CAT_AXIS + VAL_AXIS + `</c:plotArea></c:chart></c:chartSpace>`;
    const doc = documentWithChart(sample);
    const part = chartPartOf(doc);
    doc.pkg.raw()[part] = new TextEncoder().encode(authored);
    doc.markPackageResourceChanged();
    doc.refresh();

    const before = unzipSync(doc.save())[part];
    // Edit the text around the chart, then save again.
    const paragraph = doc.sections[0].blocks[0] as Paragraph;
    ((paragraph.children[0] as Run).content[0] as TextContent).srcT!.text = "Anchor edited";
    const after = unzipSync(doc.save())[part];

    expect(strFromU8(after)).toBe(authored);
    expect(after).toEqual(before);
  });
});
