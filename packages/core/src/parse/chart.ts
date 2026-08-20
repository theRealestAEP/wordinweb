import type { ChartAxis, ChartData, ChartSeries, ChartTickMark, ChartType, Theme } from "../model.js";
import { XmlElement, attr, child, children, intAttr, localName } from "../xml.js";
import { solidColorOf } from "./color.js";

/** c:plotArea children that name a plot, mapped to the type the renderer paints.
 * Everything else in that position (bar3DChart, radarChart, surfaceChart,
 * stockChart, bubbleChart, ofPieChart, …) is reported as unsupported. */
const PLOT_TYPES: Record<string, ChartType> = {
  barChart: "column",
  lineChart: "line",
  pieChart: "pie",
  doughnutChart: "doughnut",
  areaChart: "area",
  scatterChart: "scatter",
};

/** Any c:plotArea child whose name ends in "Chart" is a plot. */
function isPlot(name: string): boolean {
  return name.endsWith("Chart");
}

function descendant(node: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!node) return undefined;
  for (const item of node.children) {
    if (localName(item.name) === name) return item;
    const found = descendant(item, name);
    if (found) return found;
  }
  return undefined;
}

function descendants(node: XmlElement | undefined, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (item: XmlElement | undefined): void => {
    if (!item) return;
    for (const current of item.children) {
      if (localName(current.name) === name) out.push(current);
      walk(current);
    }
  };
  walk(node);
  return out;
}

function val(node: XmlElement | undefined): string | undefined {
  return node ? attr(node, "val") : undefined;
}

function numberVal(node: XmlElement | undefined): number | undefined {
  const raw = val(node);
  if (raw === undefined) return undefined;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : undefined;
}

/**
 * Read a c:strCache/c:numCache into a dense array indexed by c:pt/@idx.
 * Word omits the point for a blank cell, so the gaps have to be preserved:
 * the array is sized by c:ptCount and missing slots stay empty.
 */
function cachePoints(cache: XmlElement | undefined): string[] {
  if (!cache) return [];
  const points = children(cache, "pt");
  const declared = intAttr(child(cache, "ptCount"), "val") ?? 0;
  const highest = points.reduce((max, point) => Math.max(max, (intAttr(point, "idx") ?? 0) + 1), 0);
  const out = new Array<string>(Math.max(declared, highest)).fill("");
  for (const point of points) {
    const index = intAttr(point, "idx") ?? 0;
    if (index >= 0 && index < out.length) out[index] = child(point, "v")?.text ?? "";
  }
  return out;
}

/** A numeric reference (c:val, c:xVal, c:yVal): blanks become NaN so the
 * renderer can leave a gap where Word leaves one. */
function numberPoints(holder: XmlElement | undefined): number[] {
  const cache = descendant(holder, "numCache") ?? descendant(holder, "numLit");
  return cachePoints(cache).map((text) => (text === "" ? NaN : Number(text)));
}

/** The literal text of a c:tx, whether cached from a cell or typed inline. */
function referenceText(tx: XmlElement | undefined): string | undefined {
  if (!tx) return undefined;
  const cached = descendant(tx, "strCache") ?? descendant(tx, "numCache");
  if (cached) {
    const points = cachePoints(cached);
    if (points.some((point) => point !== "")) return points.filter(Boolean).join(" ");
  }
  const literal = descendants(tx, "t").map((node) => node.text).join("").trim();
  if (literal) return literal;
  const direct = child(tx, "v")?.text?.trim();
  return direct || undefined;
}

/** Series colour from c:spPr: the fill for area charts, the line for line and
 * scatter charts (which carry a:ln rather than a solid fill). */
function shapeColor(shape: XmlElement | undefined, theme: Theme | undefined): string | undefined {
  if (!shape) return undefined;
  const fill = child(shape, "solidFill");
  if (fill) return solidColorOf(fill, theme);
  const lineFill = child(child(shape, "ln"), "solidFill");
  return lineFill ? solidColorOf(lineFill, theme) : undefined;
}

/** EMU to px, the ratio every DrawingML length in this file converts by. */
const EMU_PER_PX = 9525;

/** The a:ln/@w of a shape, in px: how heavy Word draws that shape's stroke. */
function shapeLineWidth(shape: XmlElement | undefined): number | undefined {
  const width = intAttr(child(shape, "ln"), "w");
  return width ? width / EMU_PER_PX : undefined;
}

/** True when a c:dLbls block turns value labels on. */
function showsValues(labels: XmlElement | undefined): boolean {
  return !!labels && val(child(labels, "showVal")) === "1";
}

function parseSeries(element: XmlElement, theme: Theme | undefined): ChartSeries {
  const shape = child(element, "spPr");
  const explicit = shapeColor(shape, theme);
  const lineWidth = shapeLineWidth(shape);
  const marker = child(element, "marker");
  // c:size is in points, like every other ST_MarkerSize.
  const markerSize = intAttr(child(marker, "size"), "val");
  const markerSymbol = val(child(marker, "symbol"));
  const pointColors: Record<number, string> = {};
  for (const point of children(element, "dPt")) {
    const index = intAttr(child(point, "idx"), "val");
    const color = shapeColor(child(point, "spPr"), theme);
    if (index !== undefined && color) pointColors[index] = color;
  }
  // Scatter keeps its y values in c:yVal; every other type uses c:val.
  const yValues = child(element, "val") ?? child(element, "yVal");
  const xValues = child(element, "xVal");
  return {
    name: referenceText(child(element, "tx")) ?? "",
    values: numberPoints(yValues),
    ...(xValues ? { xValues: numberPoints(xValues) } : {}),
    ...(explicit ? { color: explicit } : {}),
    ...(Object.keys(pointColors).length ? { pointColors } : {}),
    ...(val(child(element, "smooth")) === "1" ? { smooth: true } : {}),
    ...(lineWidth !== undefined ? { lineWidth } : {}),
    ...(markerSymbol ? { markerSymbol } : {}),
    ...(markerSize ? { markerSize: markerSize * (96 / 72) } : {}),
  };
}

function tickMark(node: XmlElement | undefined, fallback: ChartTickMark): ChartTickMark {
  const raw = val(node);
  return raw === "none" || raw === "in" || raw === "out" || raw === "cross" ? raw : fallback;
}

/** The a:defRPr/@sz (hundredths of a point) inside a text-property holder, in px. */
function defaultTextPx(holder: XmlElement | undefined): number | undefined {
  const size = intAttr(descendant(holder, "defRPr"), "sz");
  return size ? (size / 100) * (96 / 72) : undefined;
}

function parseAxis(element: XmlElement, seriesFormat: string | undefined): ChartAxis {
  const scaling = child(element, "scaling");
  const numFmt = child(element, "numFmt");
  // sourceLinked means "use the number format of the cells the data came from",
  // which the c:numCache/c:formatCode of the series records.
  const format = numFmt && attr(numFmt, "sourceLinked") !== "1"
    ? attr(numFmt, "formatCode")
    : seriesFormat ?? (numFmt ? attr(numFmt, "formatCode") : undefined);
  const min = numberVal(child(scaling, "min"));
  const max = numberVal(child(scaling, "max"));
  const majorUnit = numberVal(child(element, "majorUnit"));
  const title = referenceText(child(child(element, "title"), "tx"));
  return {
    ...(val(child(element, "delete")) === "1" ? { hidden: true } : {}),
    ...(child(element, "majorGridlines") ? { gridlines: true } : {}),
    ...(format && format !== "General" ? { format } : {}),
    ...(title ? { title } : {}),
    ...(min !== undefined ? { min } : {}),
    ...(max !== undefined ? { max } : {}),
    ...(majorUnit !== undefined ? { majorUnit } : {}),
    ...(val(child(scaling, "orientation")) === "maxMin" ? { reversed: true } : {}),
    majorTick: tickMark(child(element, "majorTickMark"), "out"),
    minorTick: tickMark(child(element, "minorTickMark"), "none"),
    labels: val(child(element, "tickLblPos")) !== "none",
    ...(val(child(element, "crossBetween")) === "midCat"
      ? { crossBetween: "midCat" as const }
      : val(child(element, "crossBetween")) === "between"
        ? { crossBetween: "between" as const }
        : {}),
  };
}

/**
 * Split c:plotArea's axes into the category axis and the value axis.
 *
 * c:catAx/c:dateAx names the category axis outright. A scatter chart has two
 * c:valAx instead, so fall back to position: the category axis of a bar chart
 * runs up the left edge, and of every other type along the bottom.
 */
function splitAxes(plotArea: XmlElement, type: ChartType): { category?: XmlElement; value?: XmlElement } {
  const axes = plotArea.children.filter((item) => ["catAx", "valAx", "dateAx"].includes(localName(item.name)));
  const categoryEdge = type === "bar" ? "l" : "b";
  const category = axes.find((axis) => localName(axis.name) !== "valAx")
    ?? axes.find((axis) => val(child(axis, "axPos")) === categoryEdge);
  const value = axes.find((axis) => axis !== category);
  return { ...(category ? { category } : {}), ...(value ? { value } : {}) };
}

/** Resolve the cached display data from a native ChartML part. */
export function parseChartPart(root: XmlElement, theme?: Theme): ChartData | null {
  const chart = child(root, "chart");
  const plotArea = child(chart, "plotArea") ?? descendant(root, "plotArea");
  if (!plotArea) return null;
  const plot = plotArea.children.find((item) => isPlot(localName(item.name)));
  if (!plot) return null;

  const plotName = localName(plot.name);
  const title = val(child(chart, "autoTitleDeleted")) === "1"
    ? undefined
    : referenceText(child(child(chart, "title"), "tx"));
  const legendPos = child(chart, "legend") ? val(child(child(chart, "legend"), "legendPos")) ?? "r" : undefined;
  const legend = ["l", "r", "t", "b", "tr"].includes(legendPos ?? "")
    ? (legendPos as ChartData["legend"])
    : undefined;

  const mapped = PLOT_TYPES[plotName];
  if (!mapped) {
    // Keep the title and legend so the placeholder can name the chart.
    return {
      type: "column",
      unsupported: plotName,
      ...(title ? { title } : {}),
      categories: [],
      series: [],
      ...(legend ? { legend } : {}),
    };
  }

  const type: ChartType = mapped === "column" && val(child(plot, "barDir")) === "bar" ? "bar" : mapped;
  const seriesEls = children(plot, "ser");
  const series = seriesEls.map((element) => parseSeries(element, theme));
  series.forEach((item, index) => {
    if (!item.name) item.name = `Series ${index + 1}`;
  });

  // Every non-scatter type shares one category list; take it from the first
  // series that has one (Word writes it on all of them).
  const categorySource = seriesEls.map((element) => child(element, "cat")).find((cat) => cat !== undefined);
  const categoryCache = descendant(categorySource, "strCache")
    ?? descendant(categorySource, "numCache")
    ?? descendant(categorySource, "strLit")
    ?? descendant(categorySource, "numLit");
  const categories = cachePoints(categoryCache);

  const grouping = val(child(plot, "grouping"));
  const { category: categoryAxisEl, value: valueAxisEl } = splitAxes(plotArea, type);
  const seriesFormat = descendant(seriesEls[0], "formatCode")?.text;
  const dataLabels = showsValues(child(plot, "dLbls")) || seriesEls.some((el) => showsValues(child(el, "dLbls")));
  const holeSize = intAttr(child(plot, "holeSize"), "val");
  const gapWidth = intAttr(child(plot, "gapWidth"), "val");
  const overlap = intAttr(child(plot, "overlap"), "val");
  const titleSize = defaultTextPx(child(chart, "title"));
  const textSize = defaultTextPx(child(root, "txPr"));

  return {
    type,
    ...(title ? { title } : {}),
    categories,
    series,
    ...(grouping === "clustered" || grouping === "stacked" || grouping === "percentStacked" ? { grouping } : {}),
    ...(legend ? { legend } : {}),
    ...(dataLabels ? { dataLabels: true } : {}),
    ...(val(child(plot, "varyColors")) === "1" ? { varyColors: true } : {}),
    ...(categoryAxisEl ? { categoryAxis: parseAxis(categoryAxisEl, undefined) } : {}),
    ...(valueAxisEl ? { valueAxis: parseAxis(valueAxisEl, seriesFormat) } : {}),
    ...(holeSize !== undefined ? { holeSize } : {}),
    ...(gapWidth !== undefined ? { gapWidth } : {}),
    ...(overlap !== undefined ? { overlap } : {}),
    ...(val(child(plot, "marker")) === "1" ? { markers: true } : {}),
    ...(titleSize ? { titleSize } : {}),
    ...(textSize ? { textSize } : {}),
  };
}
