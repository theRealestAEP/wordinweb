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
const SWATCH = 10;
const LEGEND_GAP = 16;
const TICK_GAP = 5;

/**
 * Baseline-to-baseline distance in a legend, as a multiple of the chart text
 * size. Word set 18.1pt between legend rows at its default 10pt chart text
 * (probe-charts-basic.docx), which is the 1.81 here. Scaling by the text size
 * carries the measurement to charts that set their own size.
 */
export const LEGEND_LINE_SPACING = 1.81;

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
  if (spec.legend && spec.legendLabels.length) {
    const vertical = spec.legend === "l" || spec.legend === "r" || spec.legend === "tr";
    if (vertical) {
      const width = SWATCH + 6 + Math.max(...spec.legendLabels.map((label) => textWidth(label, spec.textSize))) + 8;
      const height = spec.legendLabels.length * spec.textSize * LEGEND_LINE_SPACING;
      const x = spec.legend === "l" ? left : right - width;
      const y = spec.legend === "tr" ? top : top + Math.max((bottom - top - height) / 2, 0);
      legend = { x, y, width, height, vertical };
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
    right -= spec.textSize * 0.5;
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
