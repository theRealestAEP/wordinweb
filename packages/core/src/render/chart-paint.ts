import type { ChartAxis, ChartData, Theme } from "../model.js";
import type { ChartItem } from "../layout/types.js";
import {
  type AxisScale,
  type MarkerShape,
  type Point,
  type Rect,
  LEGEND_LINE_KEY,
  LEGEND_LINE_SPACING,
  LEGEND_SAMPLE,
  LEGEND_SWATCH,
  LEGEND_SWATCH_KEY,
  MARKER_SHAPES,
  axisScale,
  barSlots,
  chartFrame,
  defaultOverlap,
  formatChartNumber,
  linePath,
  markerFilled,
  markerPath,
  markerShapeFor,
  stackPoints,
  textWidth,
  wedgePath,
} from "./chart-geometry.js";

/**
 * How a chart is painted, free of any drawing surface: the DOM renderer
 * paints through an SVG pen and the PDF renderer through one that writes
 * operators, so both draw the same chart from the same code.
 */

/** accent1..6 of the Office (Aptos) theme, for a document with no theme part. */
const DEFAULT_CHART_COLORS = ["#156082", "#e97132", "#196b24", "#0f9ed5", "#a02b93", "#4ea72e"];

/** Word's default chart text colour and rules: tx1 at 65% luminance for
 * labels, 15% for the axis lines and gridlines. */
export const CHART_LABEL_COLOR = "#595959";
const CHART_RULE_COLOR = "#d9d9d9";
/** Word's default chart title is 18pt and the rest of its text — axis labels,
 * legend entries, category labels — is 10pt. Both measured off the desktop
 * Word render of probe-charts-basic.docx. */
const CHART_TITLE_PX = 18 * (96 / 72);
export const CHART_TEXT_PX = 10 * (96 / 72);

/**
 * What a line or scatter series looks like when it states nothing: a 3pt
 * stroke, a 9pt marker, and a 1pt marker outline in the series' own colour.
 *
 * All three come off the PDF Word exports of probe-charts-basic, where the
 * series carried neither a c:spPr nor a c:marker and Word supplied these. A
 * series that does state them wins — insertChartAt states all three now — but
 * a chart from elsewhere leaves Word's choice standing, so draw what Word
 * draws rather than the 2.25px line and 6px marker this used to.
 */
const SERIES_LINE_PX = 3 * (96 / 72);
const MARKER_SIZE_PX = 9 * (96 / 72);
const MARKER_LINE_PX = 96 / 72;

/** Darken a hex colour toward black by `amount` (DrawingML's a:shade). */
function shade(hex: string, amount: number): string {
  const channel = (offset: number) =>
    Math.round(parseInt(hex.slice(offset, offset + 2), 16) * amount).toString(16).padStart(2, "0");
  return `#${channel(1)}${channel(3)}${channel(5)}`;
}

/**
 * Word colours chart series from the theme's accent1..6 in order. Past the
 * sixth it reuses the accents with a luminance variation rather than repeating
 * them exactly, so a seventh series stays distinguishable from the first.
 */
function chartColors(theme: Theme, count: number): string[] {
  const themed = Array.from({ length: 6 }, (_, index) => theme.colors.get(`accent${index + 1}`));
  const base = themed.every((color): color is string => !!color) ? themed : DEFAULT_CHART_COLORS;
  const out: string[] = [];
  for (let index = 0; index < Math.max(count, base.length); index++) {
    const cycle = Math.floor(index / base.length);
    out.push(cycle === 0 ? base[index] : shade(base[index % base.length], 1 - Math.min(cycle, 3) * 0.22));
  }
  return out;
}

/** Drawing surface for one chart: the palette, font and data every part of
 * the chart needs, and two ways to put marks on it. */
export interface ChartPen {
  data: ChartData;
  colors: string[];
  size: number;
  /** A CSS font-family list; the first family is what a PDF uses. */
  font: string;
  /** An SVG-shaped element: rect, line or path, with its attributes. */
  add: (name: "rect" | "line" | "path", attrs: Record<string, string | number>) => void;
  /** Text with its baseline's left end at (x, y) unless text-anchor says otherwise. */
  label: (x: number, y: number, text: string, attrs?: Record<string, string | number>) => void;
}

function pointColor(pen: ChartPen, seriesIndex: number, pointIndex: number): string {
  const series = pen.data.series[seriesIndex];
  const explicit = series?.pointColors?.[pointIndex];
  if (explicit) return explicit;
  const vary = pen.data.varyColors ?? (pen.data.type === "pie" || pen.data.type === "doughnut");
  if (vary && pen.data.series.length <= 1) return pen.colors[pointIndex % pen.colors.length];
  return series?.color ?? pen.colors[seriesIndex % pen.colors.length];
}

/**
 * The marker shape of one series: its own c:symbol when it names one this
 * renderer can draw, else its slot in Word's shape sequence. A series that
 * asks for "none" — or for a symbol with no shape behind it, like a picture —
 * gets no marker.
 */
function markerOf(pen: ChartPen, seriesIndex: number): MarkerShape | undefined {
  const symbol = pen.data.series[seriesIndex]?.markerSymbol;
  if (symbol === "none") return undefined;
  if (!symbol || symbol === "auto") return markerShapeFor(seriesIndex);
  return (MARKER_SHAPES as readonly string[]).includes(symbol) ? (symbol as MarkerShape) : undefined;
}

/** Half the width across a series' marker, which is what markerPath takes. */
function markerRadius(pen: ChartPen, seriesIndex: number): number {
  return (pen.data.series[seriesIndex]?.markerSize ?? MARKER_SIZE_PX) / 2;
}

/** A legend row. `line` and `shape` carry what the series itself draws, which
 * is what Word puts in the key for a line or scatter series. */
interface LegendEntry {
  text: string;
  color: string;
  line?: boolean;
  shape?: MarkerShape;
}

/** Legend rows: one per point for a varied single series, else one per series. */
function legendEntries(pen: ChartPen): LegendEntry[] {
  const type = pen.data.type;
  const vary = pen.data.varyColors ?? (type === "pie" || type === "doughnut");
  if (vary && pen.data.series.length <= 1) {
    return pen.data.categories.map((category, index) => ({
      text: category || `Point ${index + 1}`,
      color: pointColor(pen, 0, index),
    }));
  }
  const entries = pen.data.series.map((series, index) => ({
    text: series.name,
    color: pointColor(pen, index, -1),
    ...(type === "line" || (type === "scatter" && series.smooth) ? { line: true } : {}),
    ...(type === "scatter" || (type === "line" && pen.data.markers)
      ? { shape: markerOf(pen, index) }
      : {}),
  }));
  // A horizontal bar chart runs its categories up the left edge, so the first
  // category sits at the bottom. Word reverses the legend to match: on the
  // probe chart it lists Beta above Alpha.
  return type === "bar" ? entries.reverse() : entries;
}

/**
 * Word fills a marker with the series colour and outlines it in that same
 * colour, 1pt wide — read off its export of probe-charts-basic, where every
 * marker's stroke and fill are the one accent. An earlier reading of that
 * render called the outline white; the PDF's own path colours say otherwise.
 *
 * `tag` says where this marker sits — over a plotted point, or in a legend key
 * — because the two carry the same shape and only their place differs.
 */
function paintMarker(
  pen: ChartPen,
  at: Point,
  r: number,
  shape: MarkerShape,
  color: string,
  tag: "marker" | "legend-key" = "marker",
): void {
  const filled = markerFilled(shape);
  pen.add("path", {
    d: markerPath(shape, at, r),
    fill: filled ? color : "none",
    stroke: color,
    "stroke-width": MARKER_LINE_PX,
    [`data-dxw-chart-${tag}`]: "1",
  });
}

/**
 * One legend key at the left of its row.
 *
 * Word keys an area or a bar with a filled swatch, and a line or scatter series
 * with the line segment and marker that series draws — measured on the line
 * page of probe-charts-basic, where our filled swatch was one of the residuals.
 */
function paintLegendKey(pen: ChartPen, entry: LegendEntry, x: number, y: number, size: number): void {
  if (!entry.line && !entry.shape) {
    pen.add("rect", {
      x, y: y - size / 2, width: size, height: size, fill: entry.color,
      "data-dxw-chart-legend-key": "1",
    });
    return;
  }
  if (entry.line) {
    pen.add("line", {
      x1: x, y1: y, x2: x + size, y2: y,
      stroke: entry.color, "stroke-width": 2.25, "stroke-linecap": "round",
      "data-dxw-chart-legend-key": "1",
    });
  }
  if (entry.shape) {
    paintMarker(pen, { x: x + size / 2, y }, Math.min(size, 11) * 0.35, entry.shape, entry.color, "legend-key");
  }
}

function paintLegend(pen: ChartPen, box: Rect & { vertical: boolean }): void {
  const entries = legendEntries(pen);
  const swatch = Math.min(pen.size * 0.85, 11);
  if (box.vertical) {
    // Word's measured band (probe-legendedge): a 5.49pt swatch or a 19.2pt
    // line sample at the band's left, labels in a column 8.05pt (line: 21.42)
    // right of it.
    const step = pen.size * LEGEND_LINE_SPACING;
    entries.forEach((entry, index) => {
      const y = box.y + step * index + step / 2;
      paintLegendKey(pen, entry, box.x, y, entry.line ? LEGEND_SAMPLE : LEGEND_SWATCH);
      pen.label(box.x + (entry.line ? LEGEND_LINE_KEY : LEGEND_SWATCH_KEY), y + pen.size * 0.35, entry.text);
    });
    return;
  }
  // A horizontal legend spreads its entries evenly across the chart's width,
  // which is how Word centres a bottom legend under the plot.
  const widths = entries.map((entry) => swatch + 5 + textWidth(entry.text, pen.size));
  const total = widths.reduce((sum, width) => sum + width, 0) + (entries.length - 1) * 14;
  let x = box.x + Math.max((box.width - total) / 2, 0);
  const y = box.y + box.height / 2;
  entries.forEach((entry, index) => {
    paintLegendKey(pen, entry, x, y, swatch);
    pen.label(x + swatch + 5, y + pen.size * 0.35, entry.text);
    x += widths[index] + 14;
  });
}

/** A chart kind the engine does not paint still occupies its exact box, with
 * the ChartML's own name for it, so the page keeps its shape and the reason is
 * legible instead of blank. */
function paintUnsupported(pen: ChartPen, item: Pick<ChartItem, "width" | "height">): void {
  pen.add("rect", {
    x: 0.5, y: 0.5, width: Math.max(item.width - 1, 1), height: Math.max(item.height - 1, 1),
    fill: "#f5f6f7", stroke: "#c6cbd1", "stroke-width": 1, "stroke-dasharray": "4 3", rx: 4,
  });
  const name = pen.data.unsupported ?? "chart";
  const middle = item.height / 2;
  pen.label(item.width / 2, middle - 2, pen.data.title || "Chart", {
    "text-anchor": "middle", fill: "#404040", "font-size": Math.min(pen.size * 1.2, 15),
  });
  pen.label(item.width / 2, middle + pen.size * 1.4, `${name} is not rendered`, { "text-anchor": "middle" });
}

/** Pie and doughnut: one ring of wedges from the first series. */
function paintPie(pen: ChartPen, plot: Rect): void {
  const values = pen.data.series[0]?.values ?? [];
  const total = values.reduce((sum, value) => sum + (Number.isFinite(value) ? Math.abs(value) : 0), 0);
  if (!total) return;
  const radius = Math.max(Math.min(plot.width, plot.height) * 0.45, 8);
  const inner = pen.data.type === "doughnut" ? radius * Math.min(pen.data.holeSize ?? 75, 90) / 100 : 0;
  const center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 };
  // Word starts the first wedge at twelve o'clock and sweeps clockwise.
  let angle = -Math.PI / 2;
  values.forEach((value, index) => {
    const share = (Number.isFinite(value) ? Math.abs(value) : 0) / total;
    if (share <= 0) return;
    const next = angle + share * Math.PI * 2;
    pen.add("path", {
      d: wedgePath(center, radius, inner, angle, next),
      fill: pointColor(pen, 0, index),
      stroke: "#fff",
      "stroke-width": 1,
    });
    if (pen.data.dataLabels) {
      const mid = (angle + next) / 2;
      const distance = inner ? (radius + inner) / 2 : radius * 0.7;
      pen.label(center.x + Math.cos(mid) * distance, center.y + Math.sin(mid) * distance + pen.size * 0.35,
        formatChartNumber(Math.abs(value), pen.data.valueAxis?.format), {
          "text-anchor": "middle", fill: "#fff", "data-dxw-chart-label": "1",
        });
    }
    angle = next;
  });
}

interface Axes {
  /** Value axis mapping, along the plot's value direction. */
  value: AxisScale;
  /** Pixel position of a value on the value axis. */
  valuePos: (value: number) => number;
  /** Pixel position of a category's label, along the category axis. */
  categoryPos: (index: number) => number;
  /** Pixel positions of the category axis tick marks. Word ticks the category
   * boundaries, not the label positions, whenever c:crossBetween is "between". */
  categoryTicks: number[];
  /** The value axis after the chart's own defaults are folded in. */
  valueAxis: ChartAxis;
}

/** Draw gridlines, axis lines, ticks and tick labels around the plot.
 *
 * `horizontal` means the value axis runs left-to-right, which is the bar
 * chart's layout; every other type runs it bottom-to-top. */
function paintAxes(
  pen: ChartPen,
  plot: Rect,
  axes: Axes,
  categoryLabels: string[],
  horizontal: boolean,
): void {
  const valueAxis = axes.valueAxis;
  const categoryAxis = pen.data.categoryAxis ?? {};
  const format = valueAxis.format;
  const tickLength = 4;

  if (!valueAxis.hidden) {
    for (const tick of axes.value.ticks) {
      const at = axes.valuePos(tick);
      if (valueAxis.gridlines) {
        pen.add("line", horizontal
          ? { x1: at, y1: plot.y, x2: at, y2: plot.y + plot.height, stroke: CHART_RULE_COLOR, "stroke-width": 1 }
          : { x1: plot.x, y1: at, x2: plot.x + plot.width, y2: at, stroke: CHART_RULE_COLOR, "stroke-width": 1 });
      }
      if (valueAxis.majorTick && valueAxis.majorTick !== "none") {
        const out = valueAxis.majorTick === "in" ? 0 : tickLength;
        const into = valueAxis.majorTick === "out" ? 0 : tickLength;
        pen.add("line", horizontal
          ? { x1: at, y1: plot.y + plot.height + out, x2: at, y2: plot.y + plot.height - into, stroke: CHART_RULE_COLOR, "stroke-width": 1, "data-dxw-chart-tick": "major" }
          : { x1: plot.x - out, y1: at, x2: plot.x + into, y2: at, stroke: CHART_RULE_COLOR, "stroke-width": 1, "data-dxw-chart-tick": "major" });
      }
      if (valueAxis.labels !== false) {
        pen.label(
          horizontal ? at : plot.x - tickLength - 3,
          horizontal ? plot.y + plot.height + pen.size * 1.15 : at + pen.size * 0.35,
          formatChartNumber(tick, format),
          { "text-anchor": horizontal ? "middle" : "end" },
        );
      }
    }
    // The value axis line sits where the category axis crosses it.
    const base = horizontal ? plot.y + plot.height : plot.x;
    pen.add("line", horizontal
      ? { x1: plot.x, y1: base, x2: plot.x + plot.width, y2: base, stroke: CHART_RULE_COLOR, "stroke-width": 1 }
      : { x1: base, y1: plot.y, x2: base, y2: plot.y + plot.height, stroke: CHART_RULE_COLOR, "stroke-width": 1 });
  }

  if (!categoryAxis.hidden) {
    // Word crosses the category axis at value zero (c:crosses autoZero), so a
    // chart with negative values gets its baseline inside the plot.
    const zero = axes.valuePos(Math.min(Math.max(0, axes.value.low), axes.value.high));
    pen.add("line", horizontal
      ? { x1: zero, y1: plot.y, x2: zero, y2: plot.y + plot.height, stroke: CHART_RULE_COLOR, "stroke-width": 1 }
      : { x1: plot.x, y1: zero, x2: plot.x + plot.width, y2: zero, stroke: CHART_RULE_COLOR, "stroke-width": 1 });
    if (categoryAxis.majorTick && categoryAxis.majorTick !== "none") {
      const out = categoryAxis.majorTick === "in" ? 0 : tickLength;
      const into = categoryAxis.majorTick === "out" ? 0 : tickLength;
      for (const at of axes.categoryTicks) {
        pen.add("line", horizontal
          ? { x1: zero - out, y1: at, x2: zero + into, y2: at, stroke: CHART_RULE_COLOR, "stroke-width": 1, "data-dxw-chart-tick": "category" }
          : { x1: at, y1: zero + out, x2: at, y2: zero - into, stroke: CHART_RULE_COLOR, "stroke-width": 1, "data-dxw-chart-tick": "category" });
      }
    }
    if (categoryAxis.labels !== false) {
      // Labels stay at the low edge of the plot rather than riding the
      // baseline, so they never land on top of the bars.
      const every = Math.max(1, Math.ceil(categoryLabels.length / (horizontal ? 12 : 10)));
      categoryLabels.forEach((text, index) => {
        if (index % every !== 0 || !text) return;
        const at = axes.categoryPos(index);
        pen.label(
          horizontal ? plot.x - tickLength - 3 : at,
          horizontal ? at + pen.size * 0.35 : plot.y + plot.height + pen.size * 1.15,
          text,
          { "text-anchor": horizontal ? "end" : "middle" },
        );
      });
    }
  }

  if (valueAxis.title) {
    const x = horizontal ? plot.x + plot.width / 2 : plot.x - pen.size * 1.2;
    const y = horizontal ? plot.y + plot.height + pen.size * 2.6 : plot.y + plot.height / 2;
    pen.label(x, y, valueAxis.title, {
      "text-anchor": "middle",
      ...(horizontal ? {} : { transform: `rotate(-90 ${x} ${y})` }),
    });
  }
  if (categoryAxis.title) {
    const x = horizontal ? plot.x - pen.size * 1.2 : plot.x + plot.width / 2;
    const y = horizontal ? plot.y + plot.height / 2 : plot.y + plot.height + pen.size * 2.6;
    pen.label(x, y, categoryAxis.title, {
      "text-anchor": "middle",
      ...(horizontal ? { transform: `rotate(-90 ${x} ${y})` } : {}),
    });
  }
}

/** What Word writes for c:crossBetween when it authors this kind of chart, for
 * the rare file that leaves the element out. */
function defaultCrossBetween(type: ChartData["type"]): "between" | "midCat" {
  return type === "line" || type === "area" ? "midCat" : "between";
}

/**
 * A percent-stacked chart always fills its category, so Word pins the value
 * axis at 100% rather than opening an interval above the data the way it does
 * for an ordinary scale.
 */
function percentAxis(data: ChartData): ChartAxis | undefined {
  if (data.grouping !== "percentStacked") return data.valueAxis;
  const negative = data.series.some((series) => series.values.some((value) => value < 0));
  return {
    ...data.valueAxis,
    max: data.valueAxis?.max ?? 1,
    min: data.valueAxis?.min ?? (negative ? -1 : 0),
    format: data.valueAxis?.format ?? "0%",
  };
}

/** Column, bar, line and area: series over a shared category axis. */
function paintCategoryChart(pen: ChartPen, plot: Rect): void {
  const data = pen.data;
  const horizontal = data.type === "bar";
  const stacks = stackPoints(data.series, data.grouping);
  const scale = axisScale(stacks.flat().flatMap((point) => [point.base, point.top]), percentAxis(data));
  const range = scale.high - scale.low || 1;
  const reversed = data.valueAxis?.reversed ?? false;
  const valuePos = (value: number): number => {
    const share = (value - scale.low) / range;
    const along = reversed ? 1 - share : share;
    return horizontal ? plot.x + along * plot.width : plot.y + plot.height - along * plot.height;
  };
  const categoryCount = Math.max(data.categories.length, ...data.series.map((series) => series.values.length), 1);
  const length = horizontal ? plot.height : plot.width;
  const bars = barSlots(
    length,
    categoryCount,
    data.grouping === "stacked" || data.grouping === "percentStacked" ? 1 : data.series.length,
    data.gapWidth ?? 150,
    data.overlap ?? defaultOverlap(data.grouping),
  );
  // A bar chart's categories run up from the bottom of the plot, so its
  // category direction is the reverse of a column chart's left-to-right.
  const toCategory = (offset: number, extent = 0): number =>
    horizontal ? plot.y + plot.height - offset - extent : plot.x + offset;
  // "midCat" seats the first and last point on the plot's edges; "between"
  // insets every point to the middle of its own category band.
  const midCat = (data.valueAxis?.crossBetween ?? defaultCrossBetween(data.type)) === "midCat";
  const stride = midCat && categoryCount > 1 ? length / (categoryCount - 1) : bars.band;
  const categoryPos = (index: number): number =>
    toCategory(midCat ? index * stride : bars.band * (index + 0.5));
  const categoryTicks = midCat
    ? Array.from({ length: categoryCount }, (_, index) => toCategory(index * stride))
    : Array.from({ length: categoryCount + 1 }, (_, index) => toCategory(bars.band * index));

  paintAxes(
    pen,
    plot,
    { value: scale, valuePos, categoryPos, categoryTicks, valueAxis: percentAxis(data) ?? {} },
    data.categories,
    horizontal,
  );

  const labelValue = (seriesIndex: number, index: number): string =>
    formatChartNumber(data.series[seriesIndex].values[index], data.valueAxis?.format);

  if (data.type === "column" || data.type === "bar") {
    // Stacked series share one slot; clustered ones each get their own.
    const stacked = data.grouping === "stacked" || data.grouping === "percentStacked";
    data.series.forEach((series, seriesIndex) => {
      series.values.forEach((value, index) => {
        if (!Number.isFinite(value)) return;
        const { base, top } = stacks[seriesIndex][index];
        const near = valuePos(base);
        const far = valuePos(top);
        const thickness = Math.max(bars.size, 1);
        const offset = toCategory(bars.start(index, stacked ? 0 : seriesIndex), thickness);
        pen.add("rect", {
          x: horizontal ? Math.min(near, far) : offset,
          y: horizontal ? offset : Math.min(near, far),
          width: horizontal ? Math.max(Math.abs(far - near), 0.5) : thickness,
          height: horizontal ? thickness : Math.max(Math.abs(far - near), 0.5),
          fill: pointColor(pen, seriesIndex, index),
          "data-dxw-chart-bar": "1",
        });
        if (data.dataLabels) {
          const outside = far + (far >= near ? pen.size : -pen.size * 0.4) * (horizontal ? 1 : -1);
          pen.label(
            horizontal ? outside + (far >= near ? 2 : -2) : offset + thickness / 2,
            horizontal ? offset + thickness / 2 + pen.size * 0.35 : outside,
            labelValue(seriesIndex, index),
            { "text-anchor": horizontal ? (far >= near ? "start" : "end") : "middle", "data-dxw-chart-label": "1" },
          );
        }
      });
    });
    return;
  }

  const pointAt = (seriesIndex: number, index: number): Point | null => {
    const value = data.series[seriesIndex].values[index];
    if (!Number.isFinite(value)) return null;
    const along = categoryPos(index);
    const across = valuePos(stacks[seriesIndex][index].top);
    return horizontal ? { x: across, y: along } : { x: along, y: across };
  };

  if (data.type === "area") {
    const baseline = valuePos(Math.min(Math.max(0, scale.low), scale.high));
    data.series.forEach((series, seriesIndex) => {
      const tops = series.values.map((_, index) => pointAt(seriesIndex, index));
      const present = tops.filter((point): point is Point => !!point);
      if (present.length < 2) return;
      const stacked = data.grouping === "stacked" || data.grouping === "percentStacked";
      const under = stacked
        ? series.values.map((_, index) => ({
            x: categoryPos(index),
            y: valuePos(stacks[seriesIndex][index].base),
          })).filter((_, index) => !!tops[index]).reverse()
        : [{ x: present[present.length - 1].x, y: baseline }, { x: present[0].x, y: baseline }];
      const outline = `${linePath(tops)} L ${under.map((point) => `${point.x.toFixed(3)} ${point.y.toFixed(3)}`).join(" L ")} Z`;
      pen.add("path", { d: outline, fill: pointColor(pen, seriesIndex, -1) });
    });
    return;
  }

  // Line chart.
  data.series.forEach((series, seriesIndex) => {
    const points = series.values.map((_, index) => pointAt(seriesIndex, index));
    const color = pointColor(pen, seriesIndex, -1);
    pen.add("path", {
      d: linePath(points, series.smooth ?? false),
      fill: "none", stroke: color, "stroke-width": series.lineWidth ?? SERIES_LINE_PX,
      "stroke-linecap": "round", "stroke-linejoin": "round",
    });
    const shape = data.markers ? markerOf(pen, seriesIndex) : undefined;
    if (shape) {
      const r = markerRadius(pen, seriesIndex);
      points.forEach((point) => {
        if (point) paintMarker(pen, point, r, shape, color);
      });
    }
    if (data.dataLabels) {
      points.forEach((point, index) => {
        if (!point) return;
        pen.label(point.x, point.y - pen.size * 0.6, labelValue(seriesIndex, index), {
          "text-anchor": "middle", "data-dxw-chart-label": "1",
        });
      });
    }
  });
}

/** Scatter: both axes are value axes, and each series carries its own x list. */
function paintScatter(pen: ChartPen, plot: Rect): void {
  const data = pen.data;
  const xValues = data.series.flatMap((series) => series.xValues ?? series.values.map((_, index) => index + 1));
  const yValues = data.series.flatMap((series) => series.values);
  const xScale = axisScale(xValues, data.categoryAxis);
  const yScale = axisScale(yValues, data.valueAxis);
  const xRange = xScale.high - xScale.low || 1;
  const yRange = yScale.high - yScale.low || 1;
  const toX = (value: number): number => plot.x + ((value - xScale.low) / xRange) * plot.width;
  const toY = (value: number): number => plot.y + plot.height - ((value - yScale.low) / yRange) * plot.height;

  paintAxes(
    pen,
    plot,
    {
      value: yScale,
      valuePos: toY,
      categoryPos: (index) => toX(xScale.ticks[index] ?? xScale.low),
      categoryTicks: xScale.ticks.map(toX),
      valueAxis: data.valueAxis ?? {},
    },
    xScale.ticks.map((tick) => formatChartNumber(tick, data.categoryAxis?.format)),
    false,
  );

  data.series.forEach((series, seriesIndex) => {
    const color = pointColor(pen, seriesIndex, -1);
    const points = series.values.map((value, index) => {
      const x = series.xValues?.[index] ?? index + 1;
      return Number.isFinite(value) && Number.isFinite(x) ? { x: toX(x), y: toY(value) } : null;
    });
    // A scatter series connects its points when it carries a stroke of its own
    // — the c:spPr/a:ln insertChartAt writes — or when c:smooth bends one
    // through them. With neither stated, the markers stand alone.
    if (series.smooth || series.lineWidth !== undefined) {
      pen.add("path", {
        d: linePath(points, series.smooth ?? false),
        fill: "none", stroke: color, "stroke-width": series.lineWidth ?? SERIES_LINE_PX,
        "stroke-linecap": "round", "stroke-linejoin": "round",
      });
    }
    const shape = markerOf(pen, seriesIndex);
    if (shape) {
      const r = markerRadius(pen, seriesIndex);
      points.forEach((point) => {
        if (point) paintMarker(pen, point, r, shape, color);
      });
    }
  });
}

export function paintChart(pen: ChartPen, item: Pick<ChartItem, "width" | "height">, theme: Theme): void {
  const data = pen.data;
  pen.add("rect", { x: 0, y: 0, width: item.width, height: item.height, fill: "#fff" });

  if (data.unsupported) {
    paintUnsupported(pen, item);
    return;
  }

  const round = data.type === "pie" || data.type === "doughnut";
  pen.colors = chartColors(theme, round ? data.categories.length : data.series.length);

  const titleSize = data.titleSize ?? CHART_TITLE_PX;
  const stacks = round ? [] : stackPoints(data.series, data.grouping);
  const valueLabels = round
    ? []
    : axisScale(stacks.flat().flatMap((point) => [point.base, point.top]), percentAxis(data))
        .ticks.map((tick) => formatChartNumber(tick, percentAxis(data)?.format));
  const legendKeys = legendEntries(pen);
  // The bottom-axis label that centres on the plot's right edge: a bar
  // chart's last value tick, a line or area chart's last category. A column
  // chart centres its categories inside their bands, so nothing overhangs.
  const rightOverhangLabel =
    data.type === "bar"
      ? valueLabels[valueLabels.length - 1]
      : data.type === "line" || data.type === "area"
        ? data.categories[data.categories.length - 1]
        : undefined;
  const frame = chartFrame({
    width: item.width,
    height: item.height,
    titleSize,
    textSize: pen.size,
    ...(data.title ? { title: data.title } : {}),
    ...(data.legend ? { legend: data.legend } : {}),
    legendLabels: legendKeys.map((entry) => entry.text),
    valueLabels,
    categoryLabels: data.type === "bar" ? data.categories : [],
    horizontalValues: data.type === "bar",
    axes: !round,
    legendLineKeys: legendKeys.some((entry) => entry.line),
    ...(rightOverhangLabel !== undefined ? { rightOverhangLabel } : {}),
    ...(data.valueAxis?.title ? { valueAxisTitle: data.valueAxis.title } : {}),
    ...(data.categoryAxis?.title ? { categoryAxisTitle: data.categoryAxis.title } : {}),
  });

  if (data.title && frame.titleBaseline !== undefined) {
    pen.label(item.width / 2, frame.titleBaseline, data.title, {
      "text-anchor": "middle", "font-size": titleSize, fill: CHART_LABEL_COLOR,
    });
  }

  if (round) paintPie(pen, frame.plot);
  else if (data.type === "scatter") paintScatter(pen, frame.plot);
  else paintCategoryChart(pen, frame.plot);

  if (frame.legend) paintLegend(pen, frame.legend);
}
