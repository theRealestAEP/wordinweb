import type { ChartAxis, ChartData, ChartSeries } from "../model.js";

/**
 * Chart geometry and label formatting, kept free of the DOM so the numbers a
 * chart is painted from can be asserted directly.
 */

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Point {
  x: number;
  y: number;
}

// ---------------------------------------------------------------- number format

/**
 * Reduce a format code to the literal-and-digits text of its positive section.
 *
 * A `[$€-407]` block carries a currency symbol to print; every other bracketed
 * block ([Red], [h], a condition) directs colour or units and prints nothing.
 * `_x` reserves a character's width, `*x` fills, and `\x` escapes a literal —
 * all three affect spacing only.
 */
function positiveSection(code: string): string {
  const cleaned = code
    .replace(/\[\$([^\]\-]*)[^\]]*\]/g, "$1")
    .replace(/\[[^\]]*\]/g, "");
  let out = "";
  let quoted = false;
  for (let i = 0; i < cleaned.length; i++) {
    const ch = cleaned[i];
    if (ch === '"') { quoted = !quoted; continue; }
    if (quoted) { out += ch; continue; }
    if (ch === ";") break;
    if (ch === "_" || ch === "*") { i++; continue; }
    if (ch === "\\") { out += cleaned[++i] ?? ""; continue; }
    out += ch;
  }
  return out;
}

function groupThousands(digits: string): string {
  return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/** Word's "General" rendering: as many decimals as the value needs, no more. */
function generalNumber(value: number): string {
  if (!Number.isFinite(value)) return "";
  const rounded = Number(value.toPrecision(10));
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

/**
 * Render a value through an Excel/Word number format code.
 *
 * This covers the codes Word writes for chart axes — digit patterns with an
 * optional thousands separator, a percent scale, and literal prefixes or
 * suffixes such as a currency symbol. Anything it cannot read falls back to
 * the General rendering rather than showing the raw code.
 */
export function formatChartNumber(value: number, format?: string): string {
  if (!Number.isFinite(value)) return "";
  if (!format || format === "General") return generalNumber(value);
  const section = positiveSection(format);
  const pattern = section.match(/[#0][#0,]*(?:\.[#0]+)?/);
  if (!pattern) return generalNumber(value);
  const body = pattern[0];
  const percent = section.includes("%");
  const scaled = percent ? value * 100 : value;
  const [intPattern, fracPattern = ""] = body.split(".");
  const decimals = (fracPattern.match(/0/g) ?? []).length;
  const optional = (fracPattern.match(/#/g) ?? []).length;
  let text = Math.abs(scaled).toFixed(decimals + optional);
  if (optional) text = text.replace(/0+$/, "").replace(/\.$/, "");
  let [whole, fraction = ""] = text.split(".");
  const minimumDigits = (intPattern.replace(/,/g, "").match(/0/g) ?? []).length;
  whole = whole.padStart(minimumDigits, "0");
  if (intPattern.includes(",")) whole = groupThousands(whole);
  const number = (scaled < 0 ? "-" : "") + whole + (fraction ? `.${fraction}` : "");
  // Whatever surrounds the digit pattern is literal text: a currency symbol
  // before, a percent sign or a unit after.
  const prefix = section.slice(0, pattern.index);
  const suffix = section.slice((pattern.index ?? 0) + body.length);
  return prefix + number + suffix;
}

// ---------------------------------------------------------------- axis scale

export interface AxisScale {
  low: number;
  high: number;
  step: number;
  ticks: number[];
}

/** Round a rough interval up to the 1 / 2 / 2.5 / 5 / 10 sequence Word steps by. */
function niceStep(rough: number): number {
  const magnitude = 10 ** Math.floor(Math.log10(rough));
  const fraction = rough / magnitude;
  const nice = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 2.5 ? 2.5 : fraction <= 5 ? 5 : 10;
  return nice * magnitude;
}

/**
 * How far past the data Word's automatic scale reaches before it rounds out to
 * a major unit: 5% of the data range, added at whichever end the data runs
 * past zero. Measured over six single-series column charts that desktop Word
 * rendered at its defaults (probe-charts-autoscale.docx). Data maxima of 9.6,
 * 10, 10.4, 3.7, 47 and 0.85 gave axis maxima of 12, 12, 12, 4, 50 and 0.9.
 * The rule pins all six.
 */
const AXIS_HEADROOM = 0.05;

/**
 * Interval count the automatic major unit divides the padded span by.
 *
 * APPROXIMATE — pending probes. The measurements pin the resulting major units
 * but not the rule that produces them. The six charts above land between 6 and
 * 10 intervals, and a divisor of 10 ahead of the 1/2/2.5/5 ladder reproduces
 * every measured (axis max, major unit) pair: 12 by 2 three times, then 4 by
 * 0.5, 50 by 5 and 0.9 by 0.1. Because niceStep rounds up, 10 acts as a
 * ceiling on the interval count rather than a target — those six charts come
 * out at 6, 6, 6, 8, 10 and 9 intervals.
 */
const AXIS_TARGET_INTERVALS = 10;

/**
 * Choose the value-axis scale.
 *
 * The automatic branch keeps zero on the axis, pads the data by AXIS_HEADROOM
 * and rounds out to a nice interval, which is what Word does for the ordinary
 * case of a chart whose data spans its own magnitude. Explicit
 * c:min/c:max/c:majorUnit always win — that is how a Word user pins a scale,
 * and it is the only part of this that the file states outright.
 */
export function axisScale(values: number[], axis?: ChartAxis): AxisScale {
  const finite = values.filter((value) => Number.isFinite(value));
  const rawLow = Math.min(0, ...finite);
  const rawHigh = Math.max(0, ...finite);
  const headroom = (rawHigh - rawLow) * AXIS_HEADROOM;
  const paddedLow = rawLow < 0 ? rawLow - headroom : rawLow;
  const paddedHigh = rawHigh > 0 ? rawHigh + headroom : rawHigh;
  const span = paddedHigh - paddedLow || 1;
  const step = axis?.majorUnit && axis.majorUnit > 0
    ? axis.majorUnit
    : niceStep(span / AXIS_TARGET_INTERVALS);
  // Multiplying a step back out reintroduces binary-float dust — 0.2 × 6 is
  // 1.2000000000000002 — which would show up in a tick label.
  const clean = (value: number): number => Number(value.toPrecision(12));
  const low = clean(axis?.min ?? Math.floor(paddedLow / step) * step);
  let high = clean(axis?.max ?? Math.ceil(paddedHigh / step) * step);
  if (high <= low) high = clean(low + step);
  const ticks: number[] = [];
  for (let value = low, guard = 0; value <= high + step / 1000 && guard < 1000; value += step, guard++) {
    ticks.push(clean(value));
  }
  return { low, high, step, ticks };
}

// ---------------------------------------------------------------- stacking

export interface StackPoint {
  base: number;
  top: number;
}

/**
 * Resolve each point to the span it occupies on the value axis.
 *
 * Clustered series all start at zero. Stacked series accumulate, with negative
 * values stacking downward from zero independently of the positive run.
 * Percent-stacked series are normalised to a 0..1 share of the category, which
 * is the scale Word labels with a percent format.
 */
export function stackPoints(series: ChartSeries[], grouping?: ChartData["grouping"]): StackPoint[][] {
  const points = series.map((item) => item.values.map(() => ({ base: 0, top: 0 })));
  const categoryCount = series.reduce((max, item) => Math.max(max, item.values.length), 0);
  if (grouping !== "stacked" && grouping !== "percentStacked") {
    series.forEach((item, s) => item.values.forEach((value, c) => {
      points[s][c] = { base: 0, top: Number.isFinite(value) ? value : 0 };
    }));
    return points;
  }
  for (let c = 0; c < categoryCount; c++) {
    const total = grouping === "percentStacked"
      ? series.reduce((sum, item) => sum + Math.abs(Number.isFinite(item.values[c]) ? item.values[c] : 0), 0) || 1
      : 1;
    let up = 0;
    let down = 0;
    series.forEach((item, s) => {
      if (c >= item.values.length) return;
      const raw = Number.isFinite(item.values[c]) ? item.values[c] : 0;
      const value = raw / total;
      if (value < 0) {
        points[s][c] = { base: down, top: down + value };
        down += value;
      } else {
        points[s][c] = { base: up, top: up + value };
        up += value;
      }
    });
  }
  return points;
}

// ---------------------------------------------------------------- bar bands

export interface BarSlots {
  /** Thickness of one bar along the category axis. */
  size: number;
  /** Width of one category band. */
  band: number;
  /** Offset from the plot's category-axis origin to a bar's near edge. */
  start: (categoryIndex: number, seriesIndex: number) => number;
}

/** Word writes overlap=100 for a stacked plot, which puts every series in one
 * slot. A stacked file that omits it still means the same thing. */
export function defaultOverlap(grouping?: ChartData["grouping"]): number {
  return grouping === "stacked" || grouping === "percentStacked" ? 100 : 0;
}

/**
 * Place clustered or stacked bars inside their category bands.
 *
 * ChartML states the spacing relative to the bar itself: c:gapWidth is the gap
 * between category groups as a percent of one bar, and c:overlap is how far
 * adjacent bars in a group overlap, also as a percent of one bar. So a band of
 * n bars measures n − (n−1)·overlap + gapWidth bar-widths, which fixes the bar
 * width for a given band.
 */
export function barSlots(
  length: number,
  categoryCount: number,
  seriesCount: number,
  gapWidth = 150,
  overlap = 0,
): BarSlots {
  const band = length / Math.max(categoryCount, 1);
  const n = Math.max(seriesCount, 1);
  const step = 1 - overlap / 100;
  const bars = n - (n - 1) * (overlap / 100) + gapWidth / 100;
  const size = bars > 0 ? band / bars : band;
  const groupWidth = size * (n - (n - 1) * (overlap / 100));
  return {
    size,
    band,
    start: (categoryIndex, seriesIndex) =>
      categoryIndex * band + (band - groupWidth) / 2 + seriesIndex * size * step,
  };
}

// ---------------------------------------------------------------- frame layout

export interface ChartFrameSpec {
  width: number;
  height: number;
  titleSize: number;
  textSize: number;
  /** Present when the chart draws a title band above the plot. */
  title?: string;
  legend?: ChartData["legend"];
  legendLabels: string[];
  /** Formatted value-axis tick labels; they set the gutter on their edge. */
  valueLabels: string[];
  categoryLabels: string[];
  /** A bar chart runs its categories up the left edge and its values along the
   * bottom; every other type is the other way round. */
  horizontalValues: boolean;
  /** Charts without axes (pie, doughnut) reserve no label gutters. */
  axes: boolean;
  /** True when the legend keys its entries with a line sample rather than a
   * filled swatch (line and scatter series); the sample band is wider. */
  legendLineKeys?: boolean;
  /** The bottom-axis label that centres on the plot's right edge, when the
   * chart has one: a bar chart's last value label, a line or area chart's
   * last category. Half its width is reserved right of the plot. A column
   * chart centres its category labels inside their bands, so it passes
   * nothing. */
  rightOverhangLabel?: string;
  valueAxisTitle?: string;
  categoryAxisTitle?: string;
}

export interface ChartFrame {
  plot: Rect;
  titleBaseline?: number;
  legend?: Rect & { vertical: boolean };
}

/** Approximate advance width of a label. Chart gutters only need to be wide
 * enough not to clip, so a per-character average beats a real measurement pass
 * that the layout would have to thread a font through. */
export function textWidth(text: string, size: number): number {
  return text.length * size * 0.52;
}

const EDGE_PAD = 6;
const TICK_GAP = 5;

/**
 * Clear space Word keeps between the plot's right edge and a side legend,
 * beyond the tick allowance and edge inset that edge already gives up.
 *
 * probe-charts-basic's LINE page, read off the Word PDF's gridlines with
 * fitz get_drawings() (parity commit a5b5383). The chart space agrees
 * exactly, so everything below is in its local px:
 *
 *                    plot right      legend key      Alpha text
 *   Word                 375.62          410.02          433.48
 *   ours, before         396.71          415.33          431.33
 *
 * Word leaves 34.4 px between plot and legend where we left 18.62, and 18.62
 * is exactly what the tick allowance (textSize/2) and the edge inset already
 * spend. So the missing clearance is this 16, and it lands the plot's right
 * edge at 375.38 against Word's 375.62.
 */
const LEGEND_GAP = 16;

/**
 * A right-side legend's band and the plot edge beside it, measured on
 * probe-legendedge.docx (parity #84): bar, line and pie, each at 360x216pt
 * AND 240x144pt, read from Word's own PDF vectors with fitz get_drawings().
 * Every quantity below is identical at both box sizes, so the band is fixed
 * in pt at Word's default 10pt chart text — nothing about this edge scales
 * with the box, which retires the fractional inset and the flat LEGEND_GAP
 * on it. In CSS px (pt x 4/3):
 *
 *   swatch key, left edge to label left   8.05pt  (5.49pt swatch + 2.56 gap)
 *   line-sample key, same span           21.42pt  (19.2pt sample + 2.2 gap)
 *   label right edge to chart right       9.87pt
 *   plot right edge to key left          16.00pt + the overhanging label
 *
 * The per-type gaps decompose as 16pt plus half the width of the bottom-axis
 * label that centres on the plot's right edge: bar 16 + 5.30 ("12") = 21.30
 * against 21.33 measured, line 16 + 6.31 ("Q4") = 22.31 against 21.99, area
 * (probe-charts-basic) 16 + 6.31 = 22.31 against 22.33, column 16 + 0 =
 * 16.00 against 16.00 — a column's category labels sit inside their bands
 * and overhang nothing.
 */
export const LEGEND_SWATCH_KEY = 10.73;
export const LEGEND_LINE_KEY = 28.56;
const LEGEND_TRAILING = 13.16;
const LEGEND_CLEARANCE = 21.33;
export const LEGEND_SWATCH = 7.32;
export const LEGEND_SAMPLE = 25.6;

/**
 * Baseline-to-baseline distance in a legend, as a multiple of the chart text
 * size. Word set 18.1pt between legend rows at its default 10pt chart text
 * (probe-charts-basic.docx), which is the 1.81 here. Scaling by the text size
 * carries the measurement to charts that set their own size.
 */
export const LEGEND_LINE_SPACING = 1.81;

/**
 * Inset Word keeps on each edge of the plot rectangle, past the title, label
 * and legend bands, as a fraction of the chart box along the edge's own axis.
 *
 * Two pages of probe-charts-basic measure it, each on a different axis of the
 * same 360 x 216pt chart box. Word's value-axis gridline span beside ours
 * before this inset (parity commit 6669f9e):
 *
 *   column page   plot height  ours 161.0pt   Word 150.2pt   10.8pt too much
 *   bar page      plot width   ours 282.7pt   Word 264.8pt   17.9pt too much
 *
 * The two DIFFERENCES are what this fits, not the two absolute spans. The
 * extraction reads our own bar page 2.15pt narrower than this function
 * computes it (284.85pt), where it agrees with the column page exactly, so on
 * that page it reads both PDFs low and only the difference survives.
 *
 * One constant per edge cannot fit both differences — the column asks 5.4pt
 * and the bar 8.95pt. A fraction of the box can: 10.8 / (2 x 216) is 0.02500
 * and 17.9 / (2 x 360) is 0.02486. The least-squares fit of the two is the
 * 0.0249 here. It gives up 10.76pt on the column page against the 10.8
 * measured and 17.93pt on the bar page against 17.9 — under 0.05pt of residual
 * on each.
 *
 * A fraction is also the shape ChartML states plot geometry in: an explicit
 * c:plotArea/c:layout gives x, y, w and h as fractions of the chart space, so
 * Word's automatic layout works in the same units.
 */
const PLOT_EDGE_INSET = 0.0249;

/** Carve the chart box into title, legend and plot rectangles. */
export function chartFrame(spec: ChartFrameSpec): ChartFrame {
  let left = EDGE_PAD;
  let top = EDGE_PAD;
  let right = spec.width - EDGE_PAD;
  let bottom = spec.height - EDGE_PAD;

  let titleBaseline: number | undefined;
  if (spec.title) {
    titleBaseline = top + spec.titleSize;
    top = titleBaseline + spec.titleSize * 0.5;
  }

  let legend: (Rect & { vertical: boolean }) | undefined;
  /** The edge a vertical legend took, so the plot can clear it below. */
  let legendSide: "l" | "r" | undefined;
  if (spec.legend && spec.legendLabels.length) {
    const vertical = spec.legend === "l" || spec.legend === "r" || spec.legend === "tr";
    if (vertical) {
      // Word's band, measured on probe-legendedge (see the constants above):
      // the key segment, the longest label, then 9.87pt of trailing space to
      // the chart's own edge. The trailing constant already covers EDGE_PAD.
      const keySegment = spec.legendLineKeys ? LEGEND_LINE_KEY : LEGEND_SWATCH_KEY;
      const width =
        keySegment + Math.max(...spec.legendLabels.map((label) => textWidth(label, spec.textSize))) +
        (LEGEND_TRAILING - EDGE_PAD);
      const height = spec.legendLabels.length * spec.textSize * LEGEND_LINE_SPACING;
      const x = spec.legend === "l" ? left : right - width;
      const y = spec.legend === "tr" ? top : top + Math.max((bottom - top - height) / 2, 0);
      legend = { x, y, width, height, vertical };
      legendSide = spec.legend === "l" ? "l" : "r";
      if (spec.legend === "l") left += width; else right -= width;
    } else {
      const height = spec.textSize * LEGEND_LINE_SPACING;
      const y = spec.legend === "t" ? top : bottom - height;
      legend = { x: left, y, width: right - left, height, vertical };
      if (spec.legend === "t") top += height; else bottom -= height;
    }
  }

  if (spec.axes) {
    const valueGutter = Math.max(...spec.valueLabels.map((label) => textWidth(label, spec.textSize)), 0) + TICK_GAP + 4;
    const categoryGutter = Math.max(...spec.categoryLabels.map((label) => textWidth(label, spec.textSize)), 0) + TICK_GAP + 4;
    const lineHeight = spec.textSize * 1.4;
    if (spec.horizontalValues) {
      left += categoryGutter;
      bottom -= lineHeight;
    } else {
      left += valueGutter;
      bottom -= lineHeight;
    }
    if (spec.valueAxisTitle) {
      if (spec.horizontalValues) bottom -= lineHeight; else left += lineHeight;
    }
    if (spec.categoryAxisTitle) {
      if (spec.horizontalValues) left += lineHeight; else bottom -= lineHeight;
    }
    // Word leaves room for the topmost tick label to sit beside the axis.
    top += spec.textSize * 0.5;
    if (legendSide !== "r") right -= spec.textSize * 0.5;
    // Both probe pages are axis charts, so the inset stays where it was
    // measured; a pie keeps the whole box its title and legend leave it.
    left += spec.width * PLOT_EDGE_INSET;
    if (legendSide !== "r") right -= spec.width * PLOT_EDGE_INSET;
    top += spec.height * PLOT_EDGE_INSET;
    bottom -= spec.height * PLOT_EDGE_INSET;
    if (legendSide === "r" && legend) {
      // Beside a right legend the edge is measured directly (probe-legendedge,
      // both box sizes): the plot stops 16pt short of the legend key, plus
      // half the bottom-axis label that centres on the edge. This replaces the
      // tick allowance, the fractional inset and the flat gap on this edge —
      // all three were fitted at one box size, and the two-size sweep shows
      // the reservation is fixed in pt, not fractional.
      const overhang = spec.rightOverhangLabel
        ? textWidth(spec.rightOverhangLabel, spec.textSize) / 2
        : 0;
      right = legend.x - LEGEND_CLEARANCE - overhang;
    } else if (legendSide === "l" && !spec.horizontalValues) {
      // The left-legend clearance keeps its old flat value; no measurement
      // covers it yet.
      left += LEGEND_GAP;
    }
  }

  if (!spec.axes && legendSide === "r" && legend) {
    // Word centres a pie in the band left of a right legend, measured from
    // the chart's own left edge to 4.7pt short of the legend key — centre
    // x = (keyLeft - 4.7pt)/2 at both probe-legendedge box sizes.
    left = 0;
    right = legend.x - 6.27;
  }

  return {
    plot: { x: left, y: top, width: Math.max(right - left, 8), height: Math.max(bottom - top, 8) },
    ...(titleBaseline !== undefined ? { titleBaseline } : {}),
    ...(legend ? { legend } : {}),
  };
}

// ---------------------------------------------------------------- path helpers

/** A pie or doughnut wedge. `inner` of 0 gives a solid slice. */
export function wedgePath(
  center: Point,
  radius: number,
  inner: number,
  startAngle: number,
  endAngle: number,
): string {
  const at = (angle: number, r: number): string =>
    `${(center.x + Math.cos(angle) * r).toFixed(3)} ${(center.y + Math.sin(angle) * r).toFixed(3)}`;
  const sweep = endAngle - startAngle;
  // A lone 100% point closes on itself, where a single arc is ambiguous; two
  // half turns draw the full ring.
  if (sweep >= Math.PI * 2 - 1e-6) {
    const half = startAngle + Math.PI;
    const outer = `M ${at(startAngle, radius)} A ${radius} ${radius} 0 1 1 ${at(half, radius)}` +
      ` A ${radius} ${radius} 0 1 1 ${at(startAngle, radius)} Z`;
    if (inner <= 0) return outer;
    return `${outer} M ${at(startAngle, inner)} A ${inner} ${inner} 0 1 0 ${at(half, inner)}` +
      ` A ${inner} ${inner} 0 1 0 ${at(startAngle, inner)} Z`;
  }
  const large = sweep > Math.PI ? 1 : 0;
  if (inner <= 0) {
    return `M ${center.x} ${center.y} L ${at(startAngle, radius)}` +
      ` A ${radius} ${radius} 0 ${large} 1 ${at(endAngle, radius)} Z`;
  }
  return `M ${at(startAngle, radius)} A ${radius} ${radius} 0 ${large} 1 ${at(endAngle, radius)}` +
    ` L ${at(endAngle, inner)} A ${inner} ${inner} 0 ${large} 0 ${at(startAngle, inner)} Z`;
}

/**
 * Marker shapes in the order Word hands them out, one per series.
 *
 * The Word render of probe-charts-basic pins the first two: a diamond for
 * Alpha and a square for Beta (parity commit 6669f9e). The rest follow the
 * order Office cycles ST_MarkerStyle in. A tenth series starts over at the
 * diamond.
 */
export const MARKER_SHAPES = [
  "diamond", "square", "triangle", "x", "star", "dot", "dash", "circle", "plus",
] as const;

export type MarkerShape = (typeof MARKER_SHAPES)[number];

export function markerShapeFor(seriesIndex: number): MarkerShape {
  return MARKER_SHAPES[seriesIndex % MARKER_SHAPES.length];
}

/** Word fills the closed marker shapes; the open ones are strokes only. */
export function markerFilled(shape: MarkerShape): boolean {
  return shape !== "x" && shape !== "plus" && shape !== "star" && shape !== "dash";
}

/**
 * One marker centred on `at`, spanning 2 x `r` across whatever its shape. A
 * "dot" is Word's small filled circle, half the width of the rest.
 */
export function markerPath(shape: MarkerShape, at: Point, r: number): string {
  const n = (value: number): string => value.toFixed(3);
  const { x, y } = at;
  switch (shape) {
    case "square":
      return `M ${n(x - r)} ${n(y - r)} H ${n(x + r)} V ${n(y + r)} H ${n(x - r)} Z`;
    case "triangle":
      return `M ${n(x)} ${n(y - r)} L ${n(x + r)} ${n(y + r)} L ${n(x - r)} ${n(y + r)} Z`;
    case "x":
      return `M ${n(x - r)} ${n(y - r)} L ${n(x + r)} ${n(y + r)}` +
        ` M ${n(x + r)} ${n(y - r)} L ${n(x - r)} ${n(y + r)}`;
    case "plus":
      return `M ${n(x - r)} ${n(y)} L ${n(x + r)} ${n(y)} M ${n(x)} ${n(y - r)} L ${n(x)} ${n(y + r)}`;
    case "star":
      return `${markerPath("x", at, r)} ${markerPath("plus", at, r)}`;
    case "dash":
      return `M ${n(x - r)} ${n(y)} L ${n(x + r)} ${n(y)}`;
    case "circle":
    case "dot": {
      const radius = shape === "dot" ? r / 2 : r;
      return `M ${n(x - radius)} ${n(y)} a ${n(radius)} ${n(radius)} 0 1 0 ${n(radius * 2)} 0` +
        ` a ${n(radius)} ${n(radius)} 0 1 0 ${n(-radius * 2)} 0 Z`;
    }
  }
  return `M ${n(x)} ${n(y - r)} L ${n(x + r)} ${n(y)} L ${n(x)} ${n(y + r)} L ${n(x - r)} ${n(y)} Z`;
}

/**
 * Connect a series' points, leaving a gap wherever the data has one — which is
 * what Word's default c:dispBlanksAs of "gap" asks for. `smooth` bends the
 * joins through a Catmull-Rom spline, the curve c:smooth selects.
 */
export function linePath(points: Array<Point | null>, smooth = false): string {
  const runs: Point[][] = [];
  let run: Point[] = [];
  for (const point of points) {
    if (point) run.push(point);
    else if (run.length) { runs.push(run); run = []; }
  }
  if (run.length) runs.push(run);

  return runs.map((segment) => {
    // A point with a gap on both sides still has to show; a zero-length segment
    // under a round cap is the dot Word leaves there.
    if (segment.length === 1) {
      const only = `${segment[0].x.toFixed(3)} ${segment[0].y.toFixed(3)}`;
      return `M ${only} L ${only}`;
    }
    let d = `M ${segment[0].x.toFixed(3)} ${segment[0].y.toFixed(3)}`;
    for (let i = 0; i < segment.length - 1; i++) {
      const p1 = segment[i];
      const p2 = segment[i + 1];
      if (!smooth) { d += ` L ${p2.x.toFixed(3)} ${p2.y.toFixed(3)}`; continue; }
      const p0 = segment[i - 1] ?? p1;
      const p3 = segment[i + 2] ?? p2;
      const c1 = { x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 };
      const c2 = { x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 };
      d += ` C ${c1.x.toFixed(3)} ${c1.y.toFixed(3)} ${c2.x.toFixed(3)} ${c2.y.toFixed(3)}` +
        ` ${p2.x.toFixed(3)} ${p2.y.toFixed(3)}`;
    }
    return d;
  }).join(" ");
}
