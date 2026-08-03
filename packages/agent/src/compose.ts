import { FIXED_ZIP_MTIME } from "@wordinweb/core";
import { strToU8, zipSync } from "fflate";
import { agentOperationSchema, validateJsonSchema } from "./capabilities.js";
import type { AgentComposeBlock, AgentComposeCell, AgentComposeRequest, AgentComposeRun } from "./types.js";

type JsonSchema = Record<string, unknown>;

const string = (maxLength: number, minLength = 0): JsonSchema => ({ type: "string", maxLength, ...(minLength ? { minLength } : {}) });
const number = (minimum: number, maximum: number): JsonSchema => ({ type: "number", minimum, maximum });
const closed = (properties: Record<string, JsonSchema>, required: string[] = []): JsonSchema => ({ type: "object", properties, required, additionalProperties: false });
const field = (kind: Parameters<typeof agentOperationSchema>[0], name: string): JsonSchema =>
  ((agentOperationSchema(kind).properties as Record<string, JsonSchema>)[name]);

const runSchema = closed({
  text: string(100_000),
  bold: { type: "boolean" },
  italic: { type: "boolean" },
  underline: { type: "boolean" },
  color: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
  highlight: string(40, 1),
  fontSizePt: number(1, 1638),
  fontFamily: string(100, 1),
}, ["text"]);

const textStyleSchema = closed({
  bold: { type: "boolean" },
  italic: { type: "boolean" },
  underline: { type: "boolean" },
  color: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
  fontSizePt: number(1, 1638),
  fontFamily: string(100, 1),
  align: { enum: ["left", "center", "right"] },
});

const cellSchema = closed({
  text: string(100_000),
  runs: { type: "array", minItems: 1, maxItems: 1000, items: runSchema },
  fill: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
  align: { enum: ["left", "center", "right"] },
  verticalAlign: { enum: ["top", "center", "bottom"] },
});

const placement = {
  position: closed({ xPx: number(-5000, 5000), yPx: number(-5000, 5000) }, ["xPx", "yPx"]),
  widthPx: number(1, 5000),
  heightPx: number(1, 5000),
};
const composeColor = { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" };
const composeNullableColor = { anyOf: [composeColor, { type: "null" }] };
const composeAlignment = { enum: ["left", "center", "right"] };

const blockSchemas: JsonSchema[] = [
  closed({
    type: { const: "paragraph" },
    text: string(100_000),
    runs: { type: "array", minItems: 1, maxItems: 1000, items: runSchema },
    styleId: { enum: ["Normal", "Title", "Subtitle", "Heading1", "Heading2", "Heading3"] },
    align: { enum: ["left", "center", "right", "both"] },
    spacing: closed({ beforePt: number(0, 2000), afterPt: number(0, 2000), lineMultiple: number(0.1, 20) }),
    list: { enum: ["bullet", "number"] },
  }, ["type"]),
  closed({ type: { const: "heading" }, level: { type: "integer", minimum: 1, maximum: 3 }, text: string(100_000, 1) }, ["type", "level", "text"]),
  closed({
    type: { const: "table" },
    rows: { type: "array", minItems: 1, maxItems: 50, items: { type: "array", minItems: 1, maxItems: 50, items: { anyOf: [string(100_000), cellSchema] } } },
    headerRows: { type: "integer", minimum: 0, maximum: 50 },
    headerFill: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
    headerTextColor: { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" },
    columnWidths: { type: "array", minItems: 1, maxItems: 50, items: number(1, 5000) },
  }, ["type", "rows"]),
  closed({ type: { const: "equation" }, mathText: field("insertMath", "mathText"), align: composeAlignment }, ["type", "mathText"]),
  closed({ type: { const: "chart" }, chart: field("insertChart", "chart"), widthPx: placement.widthPx, heightPx: placement.heightPx, align: composeAlignment }, ["type", "chart"]),
  closed({ type: { const: "smartArt" }, smartArt: field("insertSmartArt", "smartArt"), widthPx: placement.widthPx, heightPx: placement.heightPx, align: composeAlignment }, ["type", "smartArt"]),
  closed({
    type: { const: "image" },
    assetRef: field("insertImage", "assetRef"),
    widthPx: placement.widthPx,
    heightPx: placement.heightPx,
    alt: field("setImageAltText", "alt"),
    align: composeAlignment,
    wrap: field("setImageWrap", "mode"),
    position: placement.position,
  }, ["type", "assetRef", "widthPx", "heightPx"]),
  closed({
    type: { const: "shape" },
    preset: field("insertShape", "preset"),
    text: string(20_000),
    textStyle: textStyleSchema,
    widthPx: placement.widthPx,
    heightPx: placement.heightPx,
    position: placement.position,
    fill: composeNullableColor,
    line: { anyOf: [{ type: "null" }, closed({ color: composeColor, widthPx: field("setDrawingLineStyle", "widthPx"), dash: field("setDrawingLineStyle", "dash") }, ["color", "widthPx", "dash"])] },
    wrap: field("setImageWrap", "mode"),
    order: field("setDrawingOrder", "order"),
  }, ["type", "preset"]),
  closed({
    type: { const: "wordArt" },
    text: field("insertWordArt", "text"),
    preset: field("insertWordArt", "preset"),
    widthPx: placement.widthPx,
    heightPx: placement.heightPx,
    position: placement.position,
    rotation: field("setDrawingRotation", "degrees"),
    fill: composeColor,
    opacity: field("setDrawingWordArtStyle", "opacity"),
    wrap: field("setImageWrap", "mode"),
    order: field("setDrawingOrder", "order"),
  }, ["type", "text", "preset"]),
  closed({
    type: { const: "pageNumber" },
    fieldKind: field("insertPageField", "fieldKind"),
    align: composeAlignment,
    color: composeColor,
    fontSizePt: number(1, 1638),
    fontFamily: string(100, 1),
    bold: { type: "boolean" },
  }, ["type", "fieldKind"]),
  closed({ type: { const: "pageBreak" } }, ["type"]),
];

const blockArray = (maximum: number): JsonSchema => ({ type: "array", minItems: 1, maxItems: maximum, items: { anyOf: blockSchemas } });

export const AGENT_COMPOSE_SCHEMA: JsonSchema = closed({
  revision: string(100, 1),
  body: blockArray(500),
  header: blockArray(100),
  footer: blockArray(100),
  firstHeader: blockArray(100),
  firstFooter: blockArray(100),
  page: field("setPageLayout", "patch"),
}, ["revision", "body"]);

export function validateAgentComposeRequest(input: unknown): string | null {
  if (input && typeof input === "object" && !Array.isArray(input)) {
    const record = input as Record<string, unknown>;
    for (const story of ["body", "header", "footer", "firstHeader", "firstFooter"] as const) {
      const blocks = record[story];
      if (!Array.isArray(blocks)) continue;
      for (let index = 0; index < blocks.length; index++) {
        const block = blocks[index];
        const type = block && typeof block === "object" && !Array.isArray(block) ? (block as { type?: unknown }).type : undefined;
        const schema = blockSchemas.find((candidate) => ((candidate.properties as Record<string, JsonSchema>).type as { const?: unknown }).const === type);
        if (!schema) return `compose.${story}[${index}] has an unknown block type`;
        const blockError = validateJsonSchema(block, schema, `compose.${story}[${index}]`);
        if (blockError) return blockError;
      }
    }
  }
  const error = validateJsonSchema(input, AGENT_COMPOSE_SCHEMA, "compose");
  if (error) return error;
  const request = input as AgentComposeRequest;
  let textLength = 0;
  for (const blocks of [request.body, request.header ?? [], request.footer ?? [], request.firstHeader ?? [], request.firstFooter ?? []]) {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        if (block.text !== undefined && block.runs !== undefined) return "compose paragraph must use text or runs";
        textLength += block.text?.length ?? block.runs?.reduce((sum, run) => sum + run.text.length, 0) ?? 0;
      } else if (block.type === "heading" || block.type === "wordArt") textLength += block.text.length;
      else if (block.type === "shape") textLength += block.text?.length ?? 0;
      else if (block.type === "table") {
        const columns = block.rows[0].length;
        if (block.rows.some((row) => row.length !== columns)) return "compose table rows must have equal lengths";
        if (block.headerRows !== undefined && block.headerRows > block.rows.length) return "compose table headerRows exceeds its row count";
        if (block.columnWidths && block.columnWidths.length !== columns) return "compose table columnWidths must match its column count";
        for (const row of block.rows) for (const cell of row) {
          if (typeof cell === "string") textLength += cell.length;
          else {
            if (cell.text !== undefined && cell.runs !== undefined) return "compose table cell must use text or runs";
            textLength += cell.text?.length ?? cell.runs?.reduce((sum, run) => sum + run.text.length, 0) ?? 0;
          }
        }
      }
      if ((block.type === "shape" || block.type === "wordArt" || block.type === "chart" || block.type === "smartArt") &&
        ((block.widthPx === undefined) !== (block.heightPx === undefined))) {
        return "compose drawing widthPx and heightPx must be supplied together";
      }
    }
  }
  return textLength > 2_000_000 ? "compose text is too large" : null;
}

function escapeText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(value: string): string {
  return escapeText(value).replace(/"/g, "&quot;");
}

function color(value: string): string {
  return value.replace(/^#/, "").toUpperCase();
}

function runXml(run: AgentComposeRun, defaults?: { bold?: boolean; color?: string }): string {
  const props = [
    run.bold ?? defaults?.bold ? "<w:b/>" : "",
    run.italic ? "<w:i/>" : "",
    run.underline ? '<w:u w:val="single"/>' : "",
    run.color ?? defaults?.color ? `<w:color w:val="${color(run.color ?? defaults!.color!)}"/>` : "",
    run.highlight ? `<w:highlight w:val="${escapeAttr(run.highlight)}"/>` : "",
    run.fontSizePt ? `<w:sz w:val="${Math.round(run.fontSizePt * 2)}"/><w:szCs w:val="${Math.round(run.fontSizePt * 2)}"/>` : "",
    run.fontFamily ? `<w:rFonts w:ascii="${escapeAttr(run.fontFamily)}" w:hAnsi="${escapeAttr(run.fontFamily)}"/>` : "",
  ].join("");
  const text = run.text.split("\n").map((part, index) => `${index ? '<w:br/>' : ""}<w:t xml:space="preserve">${escapeText(part)}</w:t>`).join("");
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}${text}</w:r>`;
}

function paragraphXml(block: Extract<AgentComposeBlock, { type: "paragraph" | "heading" }>, listNumId?: number): string {
  const heading = block.type === "heading";
  const styleId = heading ? `Heading${block.level}` : block.styleId;
  const align = heading ? undefined : block.align;
  const spacing = heading ? undefined : block.spacing;
  const pPr = [
    styleId ? `<w:pStyle w:val="${styleId}"/>` : "",
    listNumId !== undefined ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${listNumId}"/></w:numPr>` : "",
    align ? `<w:jc w:val="${align}"/>` : "",
    spacing ? `<w:spacing${spacing.beforePt !== undefined ? ` w:before="${Math.round(spacing.beforePt * 20)}"` : ""}${spacing.afterPt !== undefined ? ` w:after="${Math.round(spacing.afterPt * 20)}"` : ""}${spacing.lineMultiple !== undefined ? ` w:line="${Math.round(spacing.lineMultiple * 240)}" w:lineRule="auto"` : ""}/>` : "",
  ].join("");
  const runs = heading
    ? [{ text: block.text }]
    : block.runs ?? [{ text: block.text ?? "" }];
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${runs.map((run) => runXml(run)).join("")}</w:p>`;
}

function cellValue(cell: string | AgentComposeCell): AgentComposeCell {
  return typeof cell === "string" ? { text: cell } : cell;
}

function tableXml(block: Extract<AgentComposeBlock, { type: "table" }>, contentWidthTwips: number): string {
  const columns = block.rows[0].length;
  const widths = block.columnWidths
    ? block.columnWidths.map((width) => Math.round(width * 15))
    : Array.from({ length: columns }, () => Math.floor(contentWidthTwips / columns));
  const headerRows = block.headerRows ?? 0;
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="B7C9DB"/>`;
  const rows = block.rows.map((row, rowIndex) => {
    const isHeader = rowIndex < headerRows;
    const cells = row.map((raw, columnIndex) => {
      const cell = cellValue(raw);
      const fill = cell.fill ?? (isHeader ? block.headerFill ?? "1F4E79" : undefined);
      const runs = cell.runs ?? [{ text: cell.text ?? "" }];
      const tcPr = `<w:tcW w:w="${widths[columnIndex]}" w:type="dxa"/>${fill ? `<w:shd w:val="clear" w:fill="${color(fill)}"/>` : ""}${cell.verticalAlign ? `<w:vAlign w:val="${cell.verticalAlign}"/>` : ""}`;
      const pPr = cell.align ? `<w:pPr><w:jc w:val="${cell.align}"/></w:pPr>` : "";
      return `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p>${pPr}${runs.map((run) => runXml(run, { bold: isHeader, color: isHeader ? block.headerTextColor ?? "FFFFFF" : undefined })).join("")}</w:p></w:tc>`;
    }).join("");
    return `<w:tr>${isHeader ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${cells}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="${contentWidthTwips}" w:type="dxa"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(border).join("")}</w:tblBorders></w:tblPr><w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>${rows}</w:tbl>`;
}

function blockXml(block: AgentComposeBlock, contentWidthTwips: number, listNumId?: number): string {
  if (block.type === "paragraph" || block.type === "heading") return paragraphXml(block, listNumId);
  if (block.type === "table") return tableXml(block, contentWidthTwips);
  if (block.type === "pageBreak") return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const align = "align" in block ? block.align : undefined;
  const pPr = align ? `<w:pPr><w:jc w:val="${align}"/></w:pPr>` : "";
  const run = block.type === "pageNumber"
    ? runXml({
        text: "",
        bold: block.bold,
        color: block.color,
        fontSizePt: block.fontSizePt,
        fontFamily: block.fontFamily,
      })
    : '<w:r><w:t xml:space="preserve"></w:t></w:r>';
  return `<w:p>${pPr}${run}</w:p>`;
}

const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Subtitle"/><w:qFormat/><w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:color w:val="5B6573"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F75B5"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="180" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="404040"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>` +
  `</w:styles>`;

interface ListPlan {
  ids: WeakMap<AgentComposeBlock, number>;
  kinds: Map<number, "bullet" | "number">;
}

function planLists(stories: AgentComposeBlock[][]): ListPlan {
  const ids = new WeakMap<AgentComposeBlock, number>();
  const kinds = new Map<number, "bullet" | "number">();
  let nextId = 1;
  for (const blocks of stories) {
    let previousKind: "bullet" | "number" | undefined;
    let previousId: number | undefined;
    for (const block of blocks) {
      const kind = block.type === "paragraph" ? block.list : undefined;
      if (!kind) {
        previousKind = undefined;
        previousId = undefined;
        continue;
      }
      const id = kind === previousKind && previousId !== undefined ? previousId : nextId++;
      ids.set(block, id);
      kinds.set(id, kind);
      previousKind = kind;
      previousId = id;
    }
  }
  return { ids, kinds };
}

function numberingXml(kinds: Map<number, "bullet" | "number">): string {
  const used = new Set(kinds.values());
  const abstract = (kind: "bullet" | "number", id: number) =>
    `<w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${kind === "bullet" ? "bullet" : "decimal"}"/>` +
    `<w:lvlText w:val="${kind === "bullet" ? "•" : "%1."}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>`;
  return `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `${used.has("bullet") ? abstract("bullet", 0) : ""}${used.has("number") ? abstract("number", 1) : ""}` +
    [...kinds].map(([numId, kind]) => `<w:num w:numId="${numId}"><w:abstractNumId w:val="${kind === "bullet" ? 0 : 1}"/></w:num>`).join("") +
    `</w:numbering>`;
}

export function composeDocxBytes(request: AgentComposeRequest): Uint8Array {
  const error = validateAgentComposeRequest(request);
  if (error) throw new Error(error);
  const pageWidth = request.page?.size?.width ?? 8.5;
  const margins = request.page?.margins;
  const contentWidthTwips = Math.max(1440, Math.round((pageWidth - (margins?.left ?? 1) - (margins?.right ?? 1)) * 1440));
  const hasHeader = Boolean(request.header?.length);
  const hasFooter = Boolean(request.footer?.length);
  const hasFirstHeader = Boolean(request.firstHeader?.length);
  const hasFirstFooter = Boolean(request.firstFooter?.length);
  const hasFirstPage = hasFirstHeader || hasFirstFooter;
  const stories = [request.body, request.header ?? [], request.footer ?? [], request.firstHeader ?? [], request.firstFooter ?? []];
  const lists = planLists(stories);
  const hasNumbering = lists.kinds.size > 0;
  const storyXml = (blocks: AgentComposeBlock[]) => blocks.map((block) => blockXml(block, contentWidthTwips, lists.ids.get(block))).join("");
  const sectionRefs =
    `${hasHeader ? '<w:headerReference w:type="default" r:id="rId2"/>' : ""}` +
    `${hasFooter ? '<w:footerReference w:type="default" r:id="rId3"/>' : ""}` +
    `${hasFirstHeader ? '<w:headerReference w:type="first" r:id="rId4"/>' : ""}` +
    `${hasFirstFooter ? '<w:footerReference w:type="first" r:id="rId5"/>' : ""}`;
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body>${storyXml(request.body)}<w:sectPr>${sectionRefs}${hasFirstPage ? "<w:titlePg/>" : ""}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
  const relationships = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    ...(hasHeader ? ['<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'] : []),
    ...(hasFooter ? ['<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'] : []),
    ...(hasFirstHeader ? ['<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>'] : []),
    ...(hasFirstFooter ? ['<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>'] : []),
    ...(hasNumbering ? ['<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'] : []),
  ].join("");
  const contentTypes =
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `${hasHeader ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ""}` +
    `${hasFooter ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ""}` +
    `${hasFirstHeader ? '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ""}` +
    `${hasFirstFooter ? '<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ""}` +
    `${hasNumbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ""}</Types>`;
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(documentXml),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`),
    "word/styles.xml": strToU8(STYLES_XML),
  };
  if (hasHeader) files["word/header1.xml"] = strToU8(`<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${storyXml(request.header!)}</w:hdr>`);
  if (hasFooter) files["word/footer1.xml"] = strToU8(`<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${storyXml(request.footer!)}</w:ftr>`);
  if (hasFirstHeader) files["word/header2.xml"] = strToU8(`<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${storyXml(request.firstHeader!)}</w:hdr>`);
  if (hasFirstFooter) files["word/footer2.xml"] = strToU8(`<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${storyXml(request.firstFooter!)}</w:ftr>`);
  if (hasNumbering) files["word/numbering.xml"] = strToU8(numberingXml(lists.kinds));
  return zipSync(files, { mtime: FIXED_ZIP_MTIME });
}
