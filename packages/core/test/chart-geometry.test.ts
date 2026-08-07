import { describe, expect, it } from "vitest";
import {
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
} from "../src/render/chart-geometry.js";

describe("chart value axis scale", () => {
  it("keeps zero on the axis and opens an interval above the top value", () => {
    // 8/9/4/10/6/9 spans 10, so the nice step is 2. Word does not let the
    // tallest bar touch the top of the plot, so the scale runs one step past.
    expect(axisScale([8, 9, 4, 10, 6, 9])).toEqual({
      low: 0,
      high: 12,
      step: 2,
      ticks: [0, 2, 4, 6, 8, 10, 12],
    });
  });

  it("extends below zero for negative data", () => {
    const scale = axisScale([-4, 6, 2]);
    expect(scale.low).toBe(-6);
    expect(scale.high).toBe(8);
    expect(scale.ticks).toEqual([-6, -4, -2, 0, 2, 4, 6, 8]);
  });

  it("obeys an explicit c:min, c:max and c:majorUnit", () => {
    expect(axisScale([3, 77], { min: 0, max: 100, majorUnit: 25 })).toEqual({
      low: 0,
      high: 100,
      step: 25,
      ticks: [0, 25, 50, 75, 100],
    });
  });

  it("pins a percent-stacked scale at one when the caller caps it", () => {
    // The automatic rule would open an interval past 100%, which a chart whose
    // categories always fill has no room for.
    expect(axisScale([0, 0.25, 1]).high).toBe(1.2);
    expect(axisScale([0, 0.25, 1], { min: 0, max: 1 }).ticks).toEqual([0, 0.2, 0.4, 0.6, 0.8, 1]);
  });

  // Six single-series column charts that desktop Word rendered at its
  // defaults, read off probe-charts-autoscale.docx against
  // parity/probe-charts-autoscale-word.pdf. Word's automatic maximum is the
  // data maximum plus 5% of the range, rounded up to the next major unit.
  it.each([
    { data: [4.2, 7.8, 3.1, 9.6], high: 12, step: 2 },
    { data: [4.2, 7.8, 3.1, 10], high: 12, step: 2 },
    { data: [4.2, 7.8, 3.1, 10.4], high: 12, step: 2 },
    { data: [1.2, 2.8, 1.1, 3.7], high: 4, step: 0.5 },
    { data: [12, 31, 19, 47], high: 50, step: 5 },
    { data: [0.2, 0.6, 0.31, 0.85], high: 0.9, step: 0.1 },
  ])("matches Word's automatic scale for a maximum of $data.3", ({ data, high, step }) => {
    const scale = axisScale(data);
    expect(scale.low).toBe(0);
    expect(scale.high).toBe(high);
    expect(scale.step).toBe(step);
  });

  it("survives a series with no finite values", () => {
    const scale = axisScale([NaN, NaN]);
    expect(scale.high).toBeGreaterThan(scale.low);
    expect(scale.ticks.length).toBeGreaterThan(1);
  });
});

describe("bar placement", () => {
  it("splits a band into clustered bars by gap width and overlap", () => {
    // A 400px axis over 4 categories gives 100px bands. Two bars at the default
    // gapWidth of 150 and no overlap measure 2 + 1.5 = 3.5 bar widths per band.
    const slots = barSlots(400, 4, 2, 150, 0);
    expect(slots.band).toBe(100);
    expect(slots.size).toBeCloseTo(28.5714, 4);
    expect(slots.start(0, 0)).toBeCloseTo(21.4286, 4);
    expect(slots.start(0, 1)).toBeCloseTo(50, 4);
    expect(slots.start(3, 0)).toBeCloseTo(321.4286, 4);
  });

  it("negative overlap pushes clustered bars apart", () => {
    // Word 2016 writes gapWidth 219 / overlap -27 for a clustered column.
    const slots = barSlots(300, 3, 2, 219, -27);
    expect(slots.size).toBeCloseTo(100 / (2 + 0.27 + 2.19), 6);
    expect(slots.start(0, 1) - slots.start(0, 0)).toBeCloseTo(slots.size * 1.27, 6);
  });

  it("puts every stacked series in one slot", () => {
    const slots = barSlots(300, 3, 1, 150, defaultOverlap("stacked"));
    expect(slots.size).toBeCloseTo(40, 6);
    expect(slots.start(1, 0)).toBeCloseTo(130, 6);
  });
});

describe("series stacking", () => {
  const series = [
    { name: "A", values: [1, 2] },
    { name: "B", values: [3, 4] },
  ];

  it("leaves clustered series on the zero baseline", () => {
    expect(stackPoints(series)).toEqual([
      [{ base: 0, top: 1 }, { base: 0, top: 2 }],
      [{ base: 0, top: 3 }, { base: 0, top: 4 }],
    ]);
  });

  it("accumulates stacked series per category", () => {
    expect(stackPoints(series, "stacked")).toEqual([
      [{ base: 0, top: 1 }, { base: 0, top: 2 }],
      [{ base: 1, top: 4 }, { base: 2, top: 6 }],
    ]);
  });

  it("normalises a percent-stacked category to one", () => {
    expect(stackPoints(series, "percentStacked")[1][0]).toEqual({ base: 0.25, top: 1 });
  });

  it("stacks negative values downward, independent of the positive run", () => {
    const mixed = [{ name: "A", values: [-2] }, { name: "B", values: [3] }, { name: "C", values: [-1] }];
    expect(stackPoints(mixed, "stacked")).toEqual([
      [{ base: 0, top: -2 }],
      [{ base: 0, top: 3 }],
      [{ base: -2, top: -3 }],
    ]);
  });
});

describe("axis label formatting", () => {
  it("renders General with no trailing zeros", () => {
    expect(formatChartNumber(1234.5)).toBe("1234.5");
    expect(formatChartNumber(0.1 + 0.2)).toBe("0.3");
    expect(formatChartNumber(-0)).toBe("0");
  });

  it("scales and marks a percent format", () => {
    expect(formatChartNumber(0.25, "0%")).toBe("25%");
    expect(formatChartNumber(0.256, "0.0%")).toBe("25.6%");
  });

  it("groups thousands and pads fixed decimals", () => {
    expect(formatChartNumber(1234567, "#,##0")).toBe("1,234,567");
    expect(formatChartNumber(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(formatChartNumber(5, "0.00")).toBe("5.00");
  });

  it("keeps a literal currency prefix, including a bracketed locale code", () => {
    expect(formatChartNumber(1234.5, "$#,##0.00")).toBe("$1,234.50");
    expect(formatChartNumber(1234.5, '[$$-409]#,##0.00')).toBe("$1,234.50");
    expect(formatChartNumber(12, '#,##0" kg"')).toBe("12 kg");
  });

  it("reads only the positive section of a multi-part code", () => {
    expect(formatChartNumber(-3.2, "#,##0.0;(#,##0.0)")).toBe("-3.2");
  });

  it("falls back to General for a code with no digit pattern", () => {
    expect(formatChartNumber(7.25, "yyyy-mm-dd")).toBe("7.25");
  });

  it("renders a blank for a missing point", () => {
    expect(formatChartNumber(NaN, "0.0")).toBe("");
  });
});

describe("chart frame", () => {
  const base = {
    width: 480,
    height: 288,
    titleSize: 18.666,
    textSize: 12,
    legendLabels: [],
    valueLabels: [],
    categoryLabels: [],
    horizontalValues: false,
    axes: true,
  };

  it("reserves a title band and a value-label gutter", () => {
    const bare = chartFrame({ ...base, valueLabels: ["0", "5", "10"] });
    const titled = chartFrame({ ...base, title: "Quarterly revenue", valueLabels: ["0", "5", "10"] });
    expect(titled.titleBaseline).toBeCloseTo(24.666, 3);
    expect(titled.plot.y - bare.plot.y).toBeCloseTo(18.666 * 1.5, 3);
    expect(titled.plot.width).toBe(bare.plot.width);
  });

  it("widens the value gutter for wider tick labels", () => {
    const narrow = chartFrame({ ...base, valueLabels: ["0", "5"] });
    const wide = chartFrame({ ...base, valueLabels: ["0", "1,000,000"] });
    expect(narrow.plot.x).toBeLessThan(wide.plot.x);
    expect(narrow.plot.width).toBeGreaterThan(wide.plot.width);
  });

  it("takes a side legend out of the plot's width and a bottom one out of its height", () => {
    const right = chartFrame({ ...base, legend: "r", legendLabels: ["Managed", "Custom"] });
    const bottom = chartFrame({ ...base, legend: "b", legendLabels: ["Managed", "Custom"] });
    expect(right.legend?.vertical).toBe(true);
    expect(right.plot.x + right.plot.width).toBeLessThanOrEqual(right.legend!.x + 0.001);
    expect(bottom.legend?.vertical).toBe(false);
    expect(bottom.plot.height).toBeLessThan(right.plot.height);
    expect(bottom.plot.width).toBeGreaterThan(right.plot.width);
  });

  it("spaces legend rows at the 18.1pt Word measured for 10pt chart text", () => {
    const pt = 96 / 72;
    const legend = chartFrame({
      ...base,
      textSize: 10 * pt,
      legend: "r",
      legendLabels: ["Alpha", "Beta", "Gamma"],
    }).legend!;
    expect(legend.height / 3 / pt).toBeCloseTo(18.1, 3);
  });

  it("gives a pie no axis gutters", () => {
    const pie = chartFrame({ ...base, axes: false });
    expect(pie.plot).toEqual({ x: 6, y: 6, width: 468, height: 276 });
  });

  it("gives up the plot room Word gives up on the two measured probe pages", () => {
    // Pages 1 and 4 of probe-charts-basic: the same 480 x 288px (360 x 216pt)
    // chart box, an 18pt title over 10pt text, ticks 0..12 by 2 and a right
    // legend of Alpha and Beta. Word's value-axis gridline span was 150.2pt on
    // the column page and 264.8pt on the bar page, ours 161.0 and 282.7
    // (parity commit 6669f9e).
    const pt = 96 / 72;
    const probe = {
      width: 480,
      height: 288,
      titleSize: 18 * pt,
      textSize: 10 * pt,
      legend: "r" as const,
      valueLabels: ["0", "2", "4", "6", "8", "10", "12"],
      axes: true,
    };
    const column = chartFrame({
      ...probe,
      title: "Clustered column",
      legendLabels: ["Alpha", "Beta"],
      categoryLabels: [],
      horizontalValues: false,
    });
    const bar = chartFrame({
      ...probe,
      title: "Bar",
      legendLabels: ["Beta", "Alpha"],
      categoryLabels: ["Q1", "Q2", "Q3", "Q4"],
      horizontalValues: true,
      // A bar chart's last value tick centres on the plot's right edge.
      rightOverhangLabel: "12",
    });

    // The column page went 161.0 -> 150.24pt against Word's 150.2.
    expect(column.plot.height / pt).toBeCloseTo(150.24, 2);
    // probe-legendedge (parity #84) re-read the bar page's Word plot as
    // 28.40..295.74pt local — width 267.34pt, confirming the 2.54pt
    // extraction offset the 264.8 of 6669f9e carried. The measured right
    // edge lands within 0.9pt of Word's; the remaining ~2.2pt sits in the
    // left gutter, which no probe has decomposed yet.
    expect(bar.plot.width / pt).toBeCloseTo(264.27, 2);
    expect(Math.abs(bar.plot.width / pt - 267.34)).toBeLessThan(3.1);
  });

  it("clears a side legend by the width Word left on the probe's line page", () => {
    // probe-charts-basic's line page, re-measured by probe-legendedge
    // (parity #84) at TWO box sizes: Word's plot right sits 16pt plus half
    // the overhanging "Q4" label left of the legend's line-sample key, and
    // every quantity is identical at 360x216 and 240x144pt. In the box's
    // local CSS px Word puts the sample at 404.95 and the plot right at
    // 375.63. Our textWidth approximation carries ~1.4px of label-width
    // error into the band, so the tolerances here are its, not the rule's.
    const pt = 96 / 72;
    const ox = 134.99;
    const oy = 96;
    const line = chartFrame({
      width: 480,
      height: 288,
      titleSize: 18 * pt,
      textSize: 10 * pt,
      title: "Line",
      legend: "r" as const,
      legendLabels: ["Alpha", "Beta"],
      valueLabels: ["0", "2", "4", "6", "8", "10", "12"],
      categoryLabels: ["Q1", "Q2", "Q3", "Q4"],
      horizontalValues: false,
      axes: true,
      legendLineKeys: true,
      rightOverhangLabel: "Q4",
    });
    const plotRight = line.plot.x + line.plot.width;
    expect(Math.abs(plotRight - 375.63)).toBeLessThan(0.5);
    expect(Math.abs(line.legend!.x - 404.95)).toBeLessThan(1.5);
    // The model identity: 16pt of clearance plus half the "Q4" label.
    expect(line.legend!.x - plotRight).toBeCloseTo(21.33 + textWidth("Q4", 10 * pt) / 2, 5);
    // The probe's vertical is closed and this must not disturb it: Word's
    // 150.40 .. 350.71 against our measured 1.44 / 1.45 below, height to 0.01.
    expect(line.plot.y - (150.4 - oy)).toBeCloseTo(1.44, 1);
    expect(line.plot.y + line.plot.height - (350.71 - oy)).toBeCloseTo(1.45, 1);
    expect(line.plot.height).toBeCloseTo(200.31, 1);
    // The left edge keeps its own +5.60. Nothing here addresses it.
    expect(line.plot.x - (170.2 - ox)).toBeCloseTo(5.6, 1);
  });

  it("leaves a pie's plot rect to its title and legend alone", () => {
    // A pie sets axes: false, so it takes neither the edge inset nor the
    // legend clearance - only the band itself. No probe measures Word's pie
    // plot rect, so this pins the shape rather than a measurement.
    const bare = chartFrame({ ...base, axes: false });
    expect(bare.plot).toEqual({ x: 6, y: 6, width: 468, height: 276 });
    const withLegend = chartFrame({
      ...base,
      axes: false,
      legend: "r",
      legendLabels: ["Alpha", "Beta"],
    });
    expect(withLegend.plot.x).toBe(6);
    expect(withLegend.plot.x + withLegend.plot.width).toBeCloseTo(withLegend.legend!.x, 6);
  });
});

describe("chart markers", () => {
  it("hands shapes out in Word's order, diamond then square", () => {
    expect(markerShapeFor(0)).toBe("diamond");
    expect(markerShapeFor(1)).toBe("square");
    expect(markerShapeFor(2)).toBe("triangle");
    // A tenth series starts the nine-shape sequence over.
    expect(MARKER_SHAPES).toHaveLength(9);
    expect(markerShapeFor(9)).toBe("diamond");
  });

  it("centres each shape on its point and spans twice the radius", () => {
    expect(markerPath("diamond", { x: 10, y: 20 }, 3))
      .toBe("M 10.000 17.000 L 13.000 20.000 L 10.000 23.000 L 7.000 20.000 Z");
    expect(markerPath("square", { x: 10, y: 20 }, 3))
      .toBe("M 7.000 17.000 H 13.000 V 23.000 H 7.000 Z");
    expect(markerPath("triangle", { x: 0, y: 0 }, 2))
      .toBe("M 0.000 -2.000 L 2.000 2.000 L -2.000 2.000 Z");
    expect(markerPath("dash", { x: 0, y: 0 }, 2)).toBe("M -2.000 0.000 L 2.000 0.000");
    // A star is an x and a plus over each other: four strokes.
    expect(markerPath("star", { x: 0, y: 0 }, 1).match(/M /g)).toHaveLength(4);
    // A dot is Word's small filled circle, half the width of the rest.
    expect(markerPath("dot", { x: 0, y: 0 }, 4)).toContain("a 2.000 2.000");
  });

  it("fills the closed shapes and strokes the open ones", () => {
    expect(MARKER_SHAPES.filter(markerFilled))
      .toEqual(["diamond", "square", "triangle", "dot", "circle"]);
  });
});

describe("path helpers", () => {
  it("draws a solid pie wedge from the centre", () => {
    const d = wedgePath({ x: 100, y: 100 }, 50, 0, -Math.PI / 2, 0);
    expect(d).toBe("M 100 100 L 100.000 50.000 A 50 50 0 0 1 150.000 100.000 Z");
  });

  it("draws a doughnut wedge as a ring segment with no centre point", () => {
    const d = wedgePath({ x: 0, y: 0 }, 10, 5, 0, Math.PI);
    expect(d.startsWith("M 10.000 0.000")).toBe(true);
    expect(d).toContain("A 5 5 0 0 0");
    expect(d).not.toContain("L 0 0");
  });

  it("closes a lone full-circle point with two half turns", () => {
    const d = wedgePath({ x: 0, y: 0 }, 10, 0, -Math.PI / 2, Math.PI * 1.5);
    expect(d.match(/A 10 10/g)).toHaveLength(2);
  });

  it("breaks the line where the data has a gap", () => {
    const d = linePath([{ x: 0, y: 0 }, null, { x: 2, y: 2 }, { x: 3, y: 1 }]);
    // The lone leading point keeps a zero-length segment so it still shows.
    expect(d).toBe("M 0.000 0.000 L 0.000 0.000 M 2.000 2.000 L 3.000 1.000");
  });

  it("bends a smooth series through cubic segments", () => {
    const d = linePath([{ x: 0, y: 0 }, { x: 10, y: 10 }, { x: 20, y: 0 }], true);
    expect(d.match(/ C /g)).toHaveLength(2);
    expect(d).not.toContain(" L ");
  });
});
