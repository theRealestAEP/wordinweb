import { strToU8, zipSync } from "fflate";
import { DocxDocument } from "../docx.js";
import type { ChartData, ChartSeries } from "../model.js";
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

/** Make the cached chart data and its workbook use the same rectangular range. */
export function normalizeChartData(data: ChartData): ChartData {
  const categories = data.categories.length ? data.categories.map(String) : ["Category 1"];
  const inputSeries: ChartSeries[] = data.series.length ? data.series : [{ name: "Series 1", values: [] }];
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
  return zipSync(files);
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

function chartSeries(data: ChartData): string {
  const lastRow = data.categories.length + 1;
  return data.series.map((series, index) => {
    const column = columnName(index + 1);
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
      `<c:tx><c:strRef><c:f>Data!$${column}$1</c:f>${stringCache([series.name])}</c:strRef></c:tx>` +
      `<c:cat><c:strRef><c:f>Data!$A$2:$A$${lastRow}</c:f>${stringCache(data.categories)}</c:strRef></c:cat>` +
      `<c:val><c:numRef><c:f>Data!$${column}$2:$${column}$${lastRow}</c:f>${numberCache(series.values)}</c:numRef></c:val>` +
      (data.type === "line" ? `<c:smooth val="0"/>` : "") +
      `</c:ser>`;
  }).join("");
}

function chartPlot(data: ChartData): string {
  const series = chartSeries(data);
  if (data.type === "pie") {
    return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  const categoryAxisId = "48650112";
  const valueAxisId = "48672768";
  const horizontal = data.type === "bar";
  const chart = data.type === "line"
    ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}` +
      `<c:marker val="1"/><c:smooth val="0"/><c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/></c:lineChart>`
    : `<c:barChart><c:barDir val="${horizontal ? "bar" : "col"}"/><c:grouping val="clustered"/>` +
      `<c:varyColors val="0"/>${series}<c:gapWidth val="150"/>` +
      `<c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/></c:barChart>`;
  const categoryAxis = `<c:catAx><c:axId val="${categoryAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${horizontal ? "l" : "b"}"/><c:tickLblPos val="nextTo"/>` +
    `<c:crossAx val="${valueAxisId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>` +
    `<c:lblOffset val="100"/></c:catAx>`;
  const valueAxis = `<c:valAx><c:axId val="${valueAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${horizontal ? "b" : "l"}"/><c:majorGridlines/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/>` +
    `<c:crossAx val="${categoryAxisId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  return chart + categoryAxis + valueAxis;
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
  const data = normalizeChartData(input);
  doc.pkg.raw()[chartRel.target] = strToU8(buildChartXml(data, packageRel.id));
  doc.pkg.raw()[packageRel.target] = buildChartWorkbook(data);
  doc.markPackageResourceChanged();
  doc.refresh();
  return true;
}
