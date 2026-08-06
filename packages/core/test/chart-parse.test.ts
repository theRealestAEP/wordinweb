import { describe, expect, it } from "vitest";
import type { Theme } from "../src/model.js";
import { parseChartPart } from "../src/parse/chart.js";
import { parseXml } from "../src/xml.js";

const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

const theme: Theme = {
  majorFont: "Calibri Light",
  minorFont: "Calibri",
  colors: new Map([
    ["accent1", "#4472c4"],
    ["accent2", "#ed7d31"],
    ["accent3", "#a5a5a5"],
    ["tx1", "#000000"],
  ]),
};

/** Wrap plot-area content the way Word writes a c:chartSpace. */
function chartSpace(inner: string, extra = ""): string {
  return `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart>${inner}</c:chart>${extra}</c:chartSpace>`;
}

function parse(xml: string) {
  const result = parseChartPart(parseXml(xml), theme);
  if (!result) throw new Error("chart did not parse");
  return result;
}

/** A c:ser with cached categories and values, as Word caches them. */
function series(index: number, name: string, categories: string[], values: Array<number | null>, extra = ""): string {
  const points = (list: Array<string | number | null>) =>
    list.map((value, i) => (value === null ? "" : `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`)).join("") +
    `<c:ptCount val="${list.length}"/>`;
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    `<c:tx><c:strRef><c:f>Sheet1!$${String.fromCharCode(66 + index)}$1</c:f>` +
    `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    extra +
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${categories.length + 1}</c:f>` +
    `<c:strCache>${points(categories)}</c:strCache></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$${values.length + 1}</c:f>` +
    `<c:numCache><c:formatCode>General</c:formatCode>${points(values)}</c:numCache></c:numRef></c:val>` +
    `</c:ser>`;
}

const CAT_AXIS = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
  `<c:tickLblPos val="nextTo"/><c:crossAx val="2"/></c:catAx>`;
const VAL_AXIS = `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>` +
  `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>` +
  `<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/></c:valAx>`;

describe("ChartML parsing", () => {
  it("reads a clustered column chart with its axes, legend and spacing", () => {
    const data = parse(chartSpace(
      `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarterly revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:autoTitleDeleted val="0"/>` +
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
      `<c:varyColors val="0"/>` +
      series(0, "Managed", ["Q1", "Q2", "Q3"], [8, 9, 4]) +
      series(1, "Custom", ["Q1", "Q2", "Q3"], [10, 6, 9]) +
      `<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      CAT_AXIS + VAL_AXIS + `</c:plotArea>` +
      `<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/>`,
    ));

    expect(data.type).toBe("column");
    expect(data.title).toBe("Quarterly revenue");
    expect(data.categories).toEqual(["Q1", "Q2", "Q3"]);
    expect(data.series.map((item) => item.name)).toEqual(["Managed", "Custom"]);
    expect(data.series[1].values).toEqual([10, 6, 9]);
    expect(data.grouping).toBe("clustered");
    expect(data.legend).toBe("b");
    expect(data.gapWidth).toBe(219);
    expect(data.overlap).toBe(-27);
    expect(data.valueAxis).toMatchObject({ gridlines: true, majorTick: "none", labels: true });
    expect(data.categoryAxis).toMatchObject({ majorTick: "out", minorTick: "none" });
    expect(data.dataLabels).toBeUndefined();
  });

  it("reads a bar chart's direction from c:barDir", () => {
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>` +
      series(0, "A", ["x"], [1]) +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx>` +
      `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="1"/></c:valAx>` +
      `</c:plotArea>`,
    ));
    expect(data.type).toBe("bar");
    expect(data.categoryAxis?.majorTick).toBe("out");
  });

  it("carries a percent-stacked grouping and its axis number format", () => {
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>` +
      series(0, "A", ["x", "y"], [1, 2]) + series(1, "B", ["x", "y"], [3, 4]) +
      `<c:gapWidth val="150"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      CAT_AXIS +
      `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/>` +
      `<c:numFmt formatCode="0%" sourceLinked="0"/><c:crossAx val="1"/></c:valAx></c:plotArea>`,
    ));
    expect(data.grouping).toBe("percentStacked");
    expect(data.overlap).toBe(100);
    expect(data.valueAxis?.format).toBe("0%");
  });

  it("leaves a gap where a cached point is missing", () => {
    // Word omits c:pt for a blank cell but still counts it in c:ptCount.
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/>` +
      series(0, "Revenue", ["Q1", "Q2", "Q3"], [2, null, 5]) +
      `<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>` +
      CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
    ));
    expect(data.type).toBe("line");
    expect(data.markers).toBe(true);
    expect(data.series[0].values[0]).toBe(2);
    expect(Number.isNaN(data.series[0].values[1])).toBe(true);
    expect(data.series[0].values[2]).toBe(5);
    expect(data.categories).toEqual(["Q1", "Q2", "Q3"]);
  });

  it("resolves an explicit series colour through the theme", () => {
    const spPr = `<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></c:spPr>`;
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
      series(0, "A", ["x"], [1], spPr) +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>` + CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
    ));
    // accent2 #ed7d31 at 75% luminance.
    expect(data.series[0].color).toBe("#b25e25");
  });

  it("reads a line series' own colour from its a:ln rather than a fill", () => {
    const spPr = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="70ad47"/></a:solidFill></a:ln></c:spPr>`;
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/>` +
      series(0, "A", ["x"], [1], spPr) +
      `<c:axId val="1"/><c:axId val="2"/></c:lineChart>` + CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
    ));
    expect(data.series[0].color).toBe("#70ad47");
  });

  it("reads a pie's varied point colours and data labels", () => {
    const points = `<c:dPt><c:idx val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:dPt>` +
      `<c:dPt><c:idx val="1"/><c:bubble3D val="0"/>` +
      `<c:spPr><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></c:spPr></c:dPt>`;
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>` +
      series(0, "Share", ["A", "B", "C"], [30, 45, 25], points) +
      `<c:dLbls><c:showVal val="1"/><c:showCatName val="0"/></c:dLbls>` +
      `<c:firstSliceAng val="0"/></c:pieChart></c:plotArea>`,
    ));
    expect(data.type).toBe("pie");
    expect(data.varyColors).toBe(true);
    expect(data.dataLabels).toBe(true);
    expect(data.series[0].pointColors).toEqual({ 0: "#4472c4", 1: "#ff0000" });
    expect(data.categoryAxis).toBeUndefined();
  });

  it("reads a doughnut's hole size", () => {
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:doughnutChart><c:varyColors val="1"/>` +
      series(0, "Share", ["A", "B"], [1, 3]) +
      `<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart></c:plotArea>`,
    ));
    expect(data.type).toBe("doughnut");
    expect(data.holeSize).toBe(50);
  });

  it("reads a scatter series' x values and both value axes", () => {
    const numeric = (values: number[]) =>
      `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
      values.map((value, i) => `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`).join("") + `</c:numCache>`;
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/>` +
      `<c:ser><c:idx val="0"/><c:order val="0"/>` +
      `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/>` +
      `<c:pt idx="0"><c:v>Trial</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
      `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>${numeric([1, 2, 4])}</c:numRef></c:xVal>` +
      `<c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>${numeric([10, 14, 9])}</c:numRef></c:yVal>` +
      `<c:smooth val="0"/></c:ser>` +
      `<c:axId val="1"/><c:axId val="2"/></c:scatterChart>` +
      `<c:valAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>` +
      `<c:valAx><c:axId val="2"/><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>` +
      `</c:plotArea>`,
    ));
    expect(data.type).toBe("scatter");
    expect(data.series[0].name).toBe("Trial");
    expect(data.series[0].xValues).toEqual([1, 2, 4]);
    expect(data.series[0].values).toEqual([10, 14, 9]);
    // The bottom axis is the categories even though both are c:valAx.
    expect(data.categoryAxis?.gridlines).toBeUndefined();
    expect(data.valueAxis?.gridlines).toBe(true);
  });

  it("reads an area chart", () => {
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:areaChart><c:grouping val="stacked"/>` +
      series(0, "A", ["x", "y"], [1, 2]) +
      `<c:axId val="1"/><c:axId val="2"/></c:areaChart>` + CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
    ));
    expect(data.type).toBe("area");
    expect(data.grouping).toBe("stacked");
  });

  it("keeps axis titles, an explicit scale and a hidden axis", () => {
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
      series(0, "A", ["x"], [1]) +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:catAx><c:axId val="1"/><c:delete val="1"/><c:axPos val="b"/>` +
      `<c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:crossAx val="2"/></c:catAx>` +
      `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="maxMin"/>` +
      `<c:max val="50"/><c:min val="10"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
      `<c:title><c:tx><c:rich><a:p><a:r><a:t>Units</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:majorUnit val="5"/><c:tickLblPos val="none"/><c:crossAx val="1"/></c:valAx></c:plotArea>`,
    ));
    expect(data.categoryAxis).toMatchObject({ hidden: true, title: "Quarter" });
    expect(data.valueAxis).toMatchObject({
      title: "Units", min: 10, max: 50, majorUnit: 5, reversed: true, labels: false,
    });
  });

  it("does not mistake an axis title for the chart title", () => {
    const data = parse(chartSpace(
      `<c:autoTitleDeleted val="1"/>` +
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>` +
      series(0, "A", ["x"], [1]) +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:catAx><c:axId val="1"/><c:delete val="0"/><c:axPos val="b"/>` +
      `<c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:crossAx val="2"/></c:catAx>${VAL_AXIS}</c:plotArea>`,
    ));
    expect(data.title).toBeUndefined();
    expect(data.categoryAxis?.title).toBe("Quarter");
  });

  it("takes explicit text sizes from c:txPr and the title's run properties", () => {
    const data = parse(chartSpace(
      `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="2000"/></a:pPr>` +
      `<a:r><a:t>Big</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>` +
      series(0, "A", ["x"], [1]) + `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
      `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`,
    ));
    expect(data.title).toBe("Big");
    expect(data.titleSize).toBeCloseTo(20 * (96 / 72), 4);
    expect(data.textSize).toBeCloseTo(10 * (96 / 72), 4);
  });

  it("reports a plot kind it cannot paint instead of guessing at one", () => {
    const data = parse(chartSpace(
      `<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales in 3D</a:t></a:r></a:p></c:rich></c:tx></c:title>` +
      `<c:plotArea><c:layout/><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
      series(0, "A", ["x"], [1]) +
      `<c:axId val="1"/><c:axId val="2"/><c:axId val="3"/></c:bar3DChart>` +
      CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
    ));
    expect(data.unsupported).toBe("bar3DChart");
    expect(data.title).toBe("Sales in 3D");
    expect(data.series).toEqual([]);
  });

  it("returns nothing for a chart part with no plot", () => {
    expect(parseChartPart(parseXml(chartSpace(`<c:plotArea><c:layout/></c:plotArea>`)), theme)).toBeNull();
  });

  it("names an unnamed series by its position", () => {
    const data = parse(chartSpace(
      `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>` +
      `<c:ser><c:idx val="0"/><c:order val="0"/>` +
      `<c:val><c:numRef><c:numCache><c:ptCount val="2"/>` +
      `<c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>` +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>` + CAT_AXIS + VAL_AXIS + `</c:plotArea>`,
    ));
    expect(data.series[0].name).toBe("Series 1");
    expect(data.series[0].values).toEqual([3, 4]);
    expect(data.categories).toEqual([]);
  });
});
