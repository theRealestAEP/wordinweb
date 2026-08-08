import { strToU8, zipSync } from "fflate";
import { FIXED_ZIP_MTIME } from "../zip.js";
import { DocxDocument } from "../docx.js";
import type { ChartData, ChartSeries } from "../model.js";
import { parseChartPart } from "../parse/chart.js";
import { parseRelationships, relsPathFor } from "../parse/rels.js";
import { parseXml, type XmlElement, child, localName } from "../xml.js";

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function descendant(node: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!node) return undefined;
  if (localName(node.name) === name) return node;
  for (const item of node.children) {
    const found = descendant(item, name);
    if (found) return found;
  }
  return undefined;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function columnName(index: number): string {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    value--;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name;
}

/** The chart types this writer can author, and which of them stack. */
export const AUTHORABLE_CHART_TYPES = ["column", "bar", "line", "pie", "doughnut", "area", "scatter"] as const;
export const STACKABLE_CHART_TYPES = ["column", "bar", "area"] as const;

/** Make the cached chart data and its workbook use the same rectangular range. */
export function normalizeChartData(data: ChartData): ChartData {
  const categories = data.categories.length ? data.categories.map(String) : ["Category 1"];
  const inputSeries: ChartSeries[] = data.series.length ? data.series : [{ name: "Series 1", values: [] }];
  const stackable = (STACKABLE_CHART_TYPES as readonly string[]).includes(data.type);
  const grouping = stackable && (data.grouping === "stacked" || data.grouping === "percentStacked")
    ? data.grouping
    : undefined;
  return {
    type: data.type,
    ...(data.title?.trim() ? { title: data.title.trim() } : {}),
    categories,
    series: inputSeries.map((series, index) => ({
      name: series.name.trim() || `Series ${index + 1}`,
      values: categories.map((_, valueIndex) => {
        const value = Number(series.values[valueIndex] ?? 0);
        return Number.isFinite(value) ? value : 0;
      }),
    })),
    ...(grouping ? { grouping } : {}),
  };
}

function workbookSheetXml(data: ChartData): string {
  const rows: string[] = [];
  const stringCell = (ref: string, value: string) =>
    `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
  rows.push(
    `<row r="1">${stringCell("A1", "Category")}${data.series
      .map((series, index) => stringCell(`${columnName(index + 1)}1`, series.name))
      .join("")}</row>`,
  );
  data.categories.forEach((category, rowIndex) => {
    const row = rowIndex + 2;
    rows.push(
      `<row r="${row}">${stringCell(`A${row}`, category)}${data.series
        .map((series, seriesIndex) => `<c r="${columnName(seriesIndex + 1)}${row}"><v>${series.values[rowIndex]}</v></c>`)
        .join("")}</row>`,
    );
  });
  const lastColumn = columnName(data.series.length);
  const lastRow = data.categories.length + 1;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<dimension ref="A1:${lastColumn}${lastRow}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join("")}</sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
    `</worksheet>`;
}

/** Build the editable data workbook related from a native ChartML part. */
export function buildChartWorkbook(input: ChartData): Uint8Array {
  const data = normalizeChartData(input);
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `</Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`,
    ),
    "xl/workbook.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_R}">` +
      `<bookViews><workbookView/></bookViews><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>` +
      `<calcPr calcId="191029"/></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `</Relationships>`,
    ),
    "xl/worksheets/sheet1.xml": strToU8(workbookSheetXml(data)),
  };
  return zipSync(files, { mtime: FIXED_ZIP_MTIME });
}

function chartTitle(title: string | undefined): string {
  if (!title) return "";
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>` +
    `<a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>`;
}

function stringCache(values: string[]): string {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values
    .map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`)
    .join("")}</c:strCache>`;
}

function numberCache(values: number[]): string {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`)
    .join("")}</c:numCache>`;
}

/** Categories double as the x values of an authored scatter series; a category
 * that is not a number falls back to its 1-based position. */
function scatterX(categories: string[]): number[] {
  return categories.map((category, index) => (Number.isFinite(Number(category)) ? Number(category) : index + 1));
}

/**
 * The stroke and the marker of one line or scatter series, as Word paints them
 * when the series says nothing.
 *
 * A c:ser with no c:spPr and no c:marker leaves both to Word, and the PDF Word
 * exports of probe-charts-basic says what it picks: a 3.0pt stroke in the
 * series' accent, and a 9pt marker filled in that accent under a 1.0pt outline
 * of the same colour. Stating those outright is the same move the gridlines
 * needed (parity commit 241f276) one level down — this renderer drew a 2.25px
 * line and a 6px marker against them, which was the whole 30.03% the line probe
 * had left.
 *
 * Both numbers are measured, not assumed: a calibration DOCX carrying exactly
 * this XML exports from Word as a 3.000pt stroke and a 9.000 x 9.000pt marker
 * path, which is what its fallback already painted. 28575 EMU (2.25pt), the
 * weight Word's UI writes into some templates, would move Word's own render.
 *
 * Word's shape sequence, one per series, is in MARKER_SHAPES beside the
 * renderer that draws them; the authored symbol and the drawn one have to stay
 * the same list.
 */
const SERIES_LINE_EMU = 38100;
const MARKER_LINE_EMU = 12700;
const MARKER_SIZE_PT = 9;
const MARKER_SYMBOLS = ["diamond", "square", "triangle", "x", "star", "dot", "dash", "circle", "plus"];

/** Word colours series from the theme's accent1..6 and starts over at the
 * seventh, which is the slot the renderer paints from too. */
function seriesAccent(index: number): string {
  return `<a:schemeClr val="accent${(index % 6) + 1}"/>`;
}

function seriesShapeXml(index: number): string {
  const accent = seriesAccent(index);
  const symbol = MARKER_SYMBOLS[index % MARKER_SYMBOLS.length];
  return `<c:spPr><a:ln w="${SERIES_LINE_EMU}" cap="rnd"><a:solidFill>${accent}</a:solidFill>` +
    `<a:round/></a:ln><a:effectLst/></c:spPr>` +
    `<c:marker><c:symbol val="${symbol}"/><c:size val="${MARKER_SIZE_PT}"/>` +
    `<c:spPr><a:solidFill>${accent}</a:solidFill>` +
    `<a:ln w="${MARKER_LINE_EMU}"><a:solidFill>${accent}</a:solidFill></a:ln>` +
    `<a:effectLst/></c:spPr></c:marker>`;
}

function chartSeries(data: ChartData): string {
  const lastRow = data.categories.length + 1;
  const stroked = data.type === "line" || data.type === "scatter";
  return data.series.map((series, index) => {
    const column = columnName(index + 1);
    const name = `<c:tx><c:strRef><c:f>Data!$${column}$1</c:f>${stringCache([series.name])}</c:strRef></c:tx>`;
    const shape = stroked ? seriesShapeXml(index) : "";
    const values = `<c:numRef><c:f>Data!$${column}$2:$${column}$${lastRow}</c:f>${numberCache(series.values)}</c:numRef>`;
    if (data.type === "scatter") {
      return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${shape}` +
        `<c:xVal><c:numRef><c:f>Data!$A$2:$A$${lastRow}</c:f>${numberCache(scatterX(data.categories))}</c:numRef></c:xVal>` +
        `<c:yVal>${values}</c:yVal><c:smooth val="0"/></c:ser>`;
    }
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${shape}` +
      `<c:cat><c:strRef><c:f>Data!$A$2:$A$${lastRow}</c:f>${stringCache(data.categories)}</c:strRef></c:cat>` +
      `<c:val>${values}</c:val>` +
      (data.type === "line" ? `<c:smooth val="0"/>` : "") +
      `</c:ser>`;
  }).join("");
}

const CATEGORY_AXIS_ID = "48650112";
const VALUE_AXIS_ID = "48672768";

/**
 * The line Word's own UI puts on a chart's gridlines and axes: tx1 at 15%
 * luminance, which resolves to #D9D9D9, drawn 9525 EMU (0.75pt) wide.
 *
 * Authoring it matters. An axis or a c:majorGridlines with no c:spPr makes Word
 * fall back to solid black at 1.0pt, which is what it painted on
 * probe-charts-basic and most of that probe's remaining ink weight (parity
 * commit 6669f9e). No chart inserted through Word's own UI looks like that, and
 * #D9D9D9 at 0.75pt is what this engine's renderer already draws.
 */
const RULE_SPPR = `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">` +
  `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill>` +
  `<a:round/></a:ln><a:effectLst/></c:spPr>`;

/**
 * The chart space: white, with no frame around it.
 *
 * Word's fallback for a c:chartSpace with no c:spPr is a black hairline
 * rectangle, which the Word render of probe-charts-basic shows as a #898989
 * frame. A chart inserted through Word's UI has no frame, and this renderer
 * draws none, so state that outright rather than leaving Word to guess.
 */
const CHART_SPACE_SPPR = `<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>` +
  `<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>`;

/** The c:catAx (or, for scatter, the second c:valAx) along the chart's
 * category direction. */
function categoryAxisXml(horizontal: boolean, numeric: boolean): string {
  const shared = `<c:axId val="${CATEGORY_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${horizontal ? "l" : "b"}"/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${RULE_SPPR}` +
    `<c:crossAx val="${VALUE_AXIS_ID}"/><c:crosses val="autoZero"/>`;
  return numeric
    ? `<c:valAx>${shared}<c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx>${shared}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
}

/** Word seats a line's or area's end points on the plot edges (midCat) and
 * insets a bar's or column's inside their bands (between). A percent-stacked
 * plot gets Word's "0%" axis format. */
function valueAxisXml(horizontal: boolean, crossBetween: "between" | "midCat", percent = false): string {
  return `<c:valAx><c:axId val="${VALUE_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${horizontal ? "b" : "l"}"/>` +
    `<c:majorGridlines>${RULE_SPPR}</c:majorGridlines>` +
    (percent
      ? `<c:numFmt formatCode="0%" sourceLinked="0"/>`
      : `<c:numFmt formatCode="General" sourceLinked="1"/>`) +
    `<c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${RULE_SPPR}` +
    `<c:crossAx val="${CATEGORY_AXIS_ID}"/><c:crosses val="autoZero"/>` +
    `<c:crossBetween val="${crossBetween}"/></c:valAx>`;
}

function chartPlot(data: ChartData): string {
  const series = chartSeries(data);
  const axIds = `<c:axId val="${CATEGORY_AXIS_ID}"/><c:axId val="${VALUE_AXIS_ID}"/>`;
  if (data.type === "pie") {
    return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  if (data.type === "doughnut") {
    return `<c:doughnutChart><c:varyColors val="1"/>${series}` +
      `<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>`;
  }
  const horizontal = data.type === "bar";
  const stacked = data.grouping === "stacked" || data.grouping === "percentStacked";
  const plot = data.type === "line"
    ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}` +
      `<c:marker val="1"/><c:smooth val="0"/>${axIds}</c:lineChart>`
    : data.type === "area"
      ? `<c:areaChart><c:grouping val="${stacked ? data.grouping : "standard"}"/>` +
        `<c:varyColors val="0"/>${series}${axIds}</c:areaChart>`
      : data.type === "scatter"
        ? `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${axIds}</c:scatterChart>`
        : `<c:barChart><c:barDir val="${horizontal ? "bar" : "col"}"/>` +
          `<c:grouping val="${stacked ? data.grouping : "clustered"}"/>` +
          `<c:varyColors val="0"/>${series}<c:gapWidth val="150"/>` +
          // Word writes overlap=100 for its stacked bar/column presets: every
          // series shares one slot per category.
          (stacked ? `<c:overlap val="100"/>` : "") +
          `${axIds}</c:barChart>`;
  const crossBetween = data.type === "line" || data.type === "area" ? "midCat" : "between";
  return plot + categoryAxisXml(horizontal, data.type === "scatter") +
    valueAxisXml(horizontal, crossBetween, data.grouping === "percentStacked");
}

/** Build native ChartML with cached display data and an editable workbook link. */
export function buildChartXml(input: ChartData, workbookRelId = "rId1"): string {
  const data = normalizeChartData(input);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:style val="2"/>` +
    `<c:chart>${chartTitle(data.title)}<c:autoTitleDeleted val="${data.title ? "0" : "1"}"/>` +
    `<c:plotArea><c:layout/>${chartPlot(data)}</c:plotArea>` +
    `<c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart>` +
    CHART_SPACE_SPPR +
    `<c:externalData r:id="${escapeXml(workbookRelId)}"><c:autoUpdate val="0"/></c:externalData>` +
    `</c:chartSpace>`;
}

/** Insert a native inline chart after the run containing the caret. */
export function insertChartAt(doc: DocxDocument, caretT: XmlElement, input: ChartData): XmlElement | null {
  const caretRun = doc.findParentOf(caretT);
  const parent = caretRun && doc.findParentOf(caretRun);
  if (!caretRun || !parent || localName(caretRun.name) !== "r") return null;
  const data = normalizeChartData(input);
  const { relId } = doc.addChartResource(buildChartXml(data), buildChartWorkbook(data));
  const w = prefixOf(caretRun);
  const id = String(doc.nextDrawingId());
  const cx = String(Math.round(480 * EMU_PER_PX));
  const cy = String(Math.round(288 * EMU_PER_PX));
  const drawing = el(`${w}drawing`, {}, [
    el("wp:inline", { "xmlns:wp": NS_WP, distT: "0", distB: "0", distL: "0", distR: "0" }, [
      el("wp:extent", { cx, cy }),
      el("wp:docPr", { id, name: `Chart ${id}` }),
      el("wp:cNvGraphicFramePr"),
      el("a:graphic", { "xmlns:a": NS_A }, [
        el("a:graphicData", { uri: NS_C }, [
          el("c:chart", { "xmlns:c": NS_C, "xmlns:r": NS_R, "r:id": relId }),
        ]),
      ]),
    ]),
  ]);
  parent.children.splice(parent.children.indexOf(caretRun) + 1, 0, el(`${w}r`, {}, [drawing]));
  doc.refresh();
  return drawing;
}

/** Replace the data for a selected native chart and its embedded workbook. */
export function setChartData(doc: DocxDocument, drawing: XmlElement, input: ChartData): boolean {
  const chartRef = descendant(drawing, "chart");
  const relKey = chartRef && Object.keys(chartRef.attrs).find((key) => localName(key) === "id");
  const chartRel = relKey ? doc.documentRels.get(chartRef!.attrs[relKey]) : undefined;
  if (!chartRel || chartRel.external) return false;
  const relsXml = doc.pkg.text(relsPathFor(chartRel.target));
  if (!relsXml) return false;
  const rels = parseRelationships(parseXml(relsXml), chartRel.target);
  const packageRel = [...rels.values()].find((rel) => rel.type.endsWith("/package") && !rel.external);
  if (!packageRel) return false;
  // Writing the new data rebuilds the whole part, which would turn a plot this
  // writer cannot express (a 3-D column, a radar) into a flat one. Refuse
  // instead of silently replacing the user's chart with a different chart.
  const existing = doc.pkg.text(chartRel.target);
  if (existing && parseChartPart(parseXml(existing))?.unsupported) return false;
  const data = normalizeChartData(input);
  doc.pkg.raw()[chartRel.target] = strToU8(buildChartXml(data, packageRel.id));
  doc.pkg.raw()[packageRel.target] = buildChartWorkbook(data);
  doc.markPackageResourceChanged();
  doc.refresh();
  return true;
}
