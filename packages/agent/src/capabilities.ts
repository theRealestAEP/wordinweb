import { registeredOperationCapabilities, type RegisteredOperationKind } from "@wordinweb/core";
import { INTENT_KINDS, type Intent } from "@wordinweb/collab/client";

export interface AgentEditCapability {
  category: "text" | "paragraph" | "review" | "table" | "insert" | "drawing" | "math" | "document";
  description: string;
  required: string[];
  optional?: string[];
}

/**
 * The agent-facing coverage map for every edit operation that is NOT declared
 * in the core operation registry.
 *
 * The Record is still the coverage gate it always was: a new hand-written
 * intent fails the agent build until its tool contract is declared here.
 * Registered operations are excluded because their category, description, and
 * agent-facing fields already come from one declaration — see
 * registeredOperationCapabilities.
 */
const HAND_WRITTEN_CAPABILITIES: Record<
  Exclude<Intent["kind"], RegisteredOperationKind>,
  AgentEditCapability
> = {
  insertText: { category: "text", description: "Insert text at a run offset.", required: ["at", "text"], optional: ["suggest"] },
  deleteText: { category: "text", description: "Delete a range within one run.", required: ["blockRef", "runRef", "start", "end"] },
  splitParagraph: { category: "paragraph", description: "Split a paragraph at a run offset.", required: ["at"], optional: ["suggest"] },
  formatRun: { category: "text", description: "Format a complete run.", required: ["blockRef", "runRef", "patch"], optional: ["suggest"] },
  formatParagraph: { category: "paragraph", description: "Set paragraph alignment or style.", required: ["blockRef"], optional: ["align", "styleId", "suggest"] },
  formatRange: { category: "text", description: "Format a range within one run.", required: ["blockRef", "runRef", "start", "end", "patch"], optional: ["suggest"] },
  tableOp: { category: "table", description: "Apply a row, column, cell, or table operation.", required: ["cellRef", "op"] },
  mergeParagraph: { category: "paragraph", description: "Merge a paragraph into its predecessor.", required: ["blockRef"], optional: ["suggest"] },
  commentRun: { category: "review", description: "Add a review comment to a run.", required: ["runRef", "text"], optional: ["initials"] },
  pasteBlocks: { category: "insert", description: "Insert validated OOXML paragraph blocks.", required: ["afterBlockRef", "blocksXml"] },
  insertImage: { category: "insert", description: "Insert a registered image asset.", required: ["runRef", "assetRef", "widthPx", "heightPx"] },
  suggestRevision: { category: "review", description: "Mark text ranges or paragraph glyphs as suggested deletions.", required: [], optional: ["ranges", "marks"] },
  insertBreak: { category: "insert", description: "Insert a page or column break.", required: ["runRef", "breakKind"] },
  insertMath: { category: "math", description: "Insert an equation from linear math text.", required: ["runRef", "mathText"] },
  insertShape: { category: "insert", description: "Insert a shape or text box.", required: ["runRef", "preset"], optional: ["text"] },
  replyComment: { category: "review", description: "Reply to a review comment.", required: ["parentId", "text"], optional: ["initials"] },
  adjustIndent: { category: "paragraph", description: "Increase or decrease a paragraph indent.", required: ["blockRef", "direction"], optional: ["suggest"] },
  setSpacing: { category: "paragraph", description: "Set paragraph spacing.", required: ["blockRef", "patch"], optional: ["suggest"] },
  insertPageField: { category: "insert", description: "Insert a page-number field.", required: ["runRef", "fieldKind"] },
  setLink: { category: "text", description: "Wrap a run in a hyperlink.", required: ["runRef", "url"] },
  insertFootnote: { category: "insert", description: "Insert a footnote.", required: ["runRef", "text"] },
  setDropCap: { category: "paragraph", description: "Set or clear a drop cap.", required: ["blockRef", "mode"] },
  setDivider: { category: "paragraph", description: "Set or clear a paragraph divider.", required: ["blockRef", "divider"] },
  insertBookmark: { category: "insert", description: "Insert a bookmark at a run.", required: ["runRef", "name"] },
  insertBlankPage: { category: "insert", description: "Insert a blank page.", required: ["runRef"] },
  insertSectionBreak: { category: "insert", description: "Insert a section break.", required: ["runRef", "breakType"] },
  insertCrossRef: { category: "insert", description: "Insert a bookmark cross-reference.", required: ["runRef", "bookmark", "refKind"] },
  insertCoverPage: { category: "insert", description: "Insert a cover page.", required: ["content"] },
  setPageLayout: { category: "document", description: "Set page layout properties.", required: ["patch"] },
  setListLevel: { category: "paragraph", description: "Change a list nesting level.", required: ["blockRef", "delta"] },
  insertWordArt: { category: "insert", description: "Insert WordArt.", required: ["runRef", "text", "preset"] },
  insertChart: { category: "insert", description: "Insert a chart.", required: ["runRef", "chart"] },
  insertSmartArt: { category: "insert", description: "Insert SmartArt.", required: ["runRef", "smartArt"] },
  setLineNumbering: { category: "document", description: "Configure margin line numbering.", required: ["patch"] },
  insertDateTimeField: { category: "insert", description: "Insert a date or time field.", required: ["runRef", "dtKind", "picture"] },
  insertField: { category: "insert", description: "Insert a safe Word field.", required: ["runRef", "instruction"], optional: ["cachedResult"] },
  setDrawingRotation: { category: "drawing", description: "Rotate a drawing.", required: ["objectRef", "degrees"] },
  setDrawingFill: { category: "drawing", description: "Set or clear a drawing fill.", required: ["objectRef", "color"] },
  setChartData: { category: "drawing", description: "Replace chart data.", required: ["objectRef", "chart"] },
  setSmartArtNodeText: { category: "drawing", description: "Set SmartArt node text.", required: ["objectRef", "index", "text"] },
  setDrawingWordArtText: { category: "drawing", description: "Set WordArt text.", required: ["objectRef", "text"] },
  setDrawingWordArtStyle: { category: "drawing", description: "Set WordArt glyph color and opacity.", required: ["objectRef", "color", "opacity"] },
  setDrawingLineStyle: { category: "drawing", description: "Set or clear a drawing outline.", required: ["objectRef", "color"], optional: ["widthPx", "dash"] },
  setImageAltText: { category: "drawing", description: "Set image alternative text.", required: ["objectRef", "alt"] },
  removeLink: { category: "text", description: "Remove a hyperlink.", required: ["runRef"] },
  setImageWrap: { category: "drawing", description: "Set drawing text wrapping.", required: ["objectRef", "mode"] },
  setDrawingOrder: { category: "drawing", description: "Move a drawing to the front or back.", required: ["objectRef", "order"] },
  setSmartArtData: { category: "drawing", description: "Replace SmartArt data.", required: ["objectRef", "smartArt"] },
  setSmartArtFill: { category: "drawing", description: "Set or clear SmartArt fill.", required: ["objectRef", "color"], optional: ["nodeIndex"] },
  setSmartArtTextFormat: { category: "drawing", description: "Set SmartArt text formatting.", required: ["objectRef", "format"], optional: ["nodeIndex"] },
  setFloatingPagePosition: { category: "drawing", description: "Position a floating drawing on the page.", required: ["objectRef", "xPx", "yPx"] },
  resizeDrawing: { category: "drawing", description: "Resize a drawing.", required: ["objectRef", "widthPx", "heightPx"] },
  resizeTableColumn: { category: "table", description: "Resize a table column.", required: ["cellRef", "boundary", "deltaPx"], optional: ["renderedWidths"] },
  moveTable: { category: "table", description: "Move a floating table.", required: ["cellRef", "xPx", "yPx", "preservePageStart", "pageDelta"] },
  removeDrawing: { category: "drawing", description: "Remove a drawing.", required: ["objectRef"] },
  setMathLinear: { category: "math", description: "Replace an equation from linear math text.", required: ["blockRef", "mathText"] },
  deleteMath: { category: "math", description: "Delete an equation.", required: ["blockRef"] },
  moveMath: { category: "math", description: "Move an equation to a text position.", required: ["blockRef", "at"] },
  ensureHeaderFooter: { category: "document", description: "Create a header or footer story.", required: ["hfKind"] },
  deleteComment: { category: "review", description: "Delete a comment thread.", required: ["commentId"] },
  insertBookmarkRange: { category: "insert", description: "Wrap a run range in a bookmark.", required: ["runRef", "name", "start", "end"] },
  acceptRevision: { category: "review", description: "Accept one tracked revision.", required: ["index"] },
  rejectRevision: { category: "review", description: "Reject one tracked revision.", required: ["index"] },
  acceptAllRevisions: { category: "review", description: "Accept all tracked revisions.", required: [] },
  rejectAllRevisions: { category: "review", description: "Reject all tracked revisions.", required: [] },
};

/** The complete agent-facing coverage map: hand-written rows plus the rows
 * the core operation registry declares. */
export const AGENT_EDIT_CAPABILITIES: Record<Intent["kind"], AgentEditCapability> = {
  ...HAND_WRITTEN_CAPABILITIES,
  ...registeredOperationCapabilities(),
};

type JsonSchema = Record<string, unknown>;

const closedObject = (
  properties: Record<string, JsonSchema>,
  required: string[] = [],
): JsonSchema => ({ type: "object", properties, required, additionalProperties: false });

const boolean = { type: "boolean" };
const number = (minimum?: number, maximum?: number): JsonSchema => ({
  type: "number",
  ...(minimum !== undefined ? { minimum } : {}),
  ...(maximum !== undefined ? { maximum } : {}),
});
const integer = (minimum?: number, maximum?: number): JsonSchema => ({
  type: "integer",
  ...(minimum !== undefined ? { minimum } : {}),
  ...(maximum !== undefined ? { maximum } : {}),
});
const string = (maxLength?: number, minLength = 0): JsonSchema => ({
  type: "string",
  ...(minLength ? { minLength } : {}),
  ...(maxLength !== undefined ? { maxLength } : {}),
});
const rgb = { type: "string", pattern: "^#?[0-9A-Fa-f]{6}$" };
const wirePosition = closedObject({
  blockRef: { type: "string", pattern: "^block:[0-9]+$" },
  runRef: { type: "string", pattern: "^run:[0-9]+$" },
  offset: integer(0),
}, ["blockRef", "runRef", "offset"]);

const styleId = { type: "string", pattern: "^[A-Za-z0-9\\-_]{1,253}$" };

const runFormatPatch = closedObject({
  bold: boolean,
  italic: boolean,
  underline: boolean,
  strike: boolean,
  color: { anyOf: [rgb, { type: "null" }] },
  highlight: { type: ["string", "null"] },
  fontSizePt: number(1, 1638),
  fontFamily: string(100, 1),
  verticalAlign: { enum: ["superscript", "subscript", null] },
  // Applying a character style is a run-patch property rather than an
  // operation of its own, because it splits the run at a partial selection
  // exactly as the other properties do. This WIDENS formatRun/formatRange:
  // every payload that validated before still validates.
  characterStyleId: { anyOf: [styleId, { type: "null" }] },
  clear: boolean,
});

/** The paragraph properties a style DEFINITION can carry (styles.xml), as
 * opposed to the direct paragraph formatting setSpacing patches. */
const styleParaPatch = closedObject({
  alignment: { enum: ["left", "center", "right", "both", null] },
  spacingBeforePt: { anyOf: [number(0, 1584), { type: "null" }] },
  spacingAfterPt: { anyOf: [number(0, 1584), { type: "null" }] },
  lineMultiple: { anyOf: [number(0.1, 132), { type: "null" }] },
  indentLeftPt: { anyOf: [number(-1584, 1584), { type: "null" }] },
  indentFirstLinePt: { anyOf: [number(-1584, 1584), { type: "null" }] },
  keepNext: { type: ["boolean", "null"] },
  outlineLevel: { anyOf: [integer(0, 8), { type: "null" }] },
});

const styleRunPatch = closedObject({
  bold: boolean,
  italic: boolean,
  underline: boolean,
  strike: boolean,
  color: { anyOf: [rgb, { type: "null" }] },
  highlight: { type: ["string", "null"] },
  fontSizePt: number(1, 1638),
  fontFamily: string(64, 1),
  verticalAlign: { enum: ["superscript", "subscript", null] },
  characterStyleId: { anyOf: [styleId, { type: "null" }] },
});

const styleSpec = closedObject(
  {
    styleId,
    type: { enum: ["paragraph", "character"] },
    name: string(253, 1),
    basedOn: { anyOf: [styleId, { type: "null" }] },
    next: { anyOf: [styleId, { type: "null" }] },
    quickStyle: boolean,
    uiPriority: integer(0, 99),
    paragraph: styleParaPatch,
    run: styleRunPatch,
  },
  ["styleId", "type", "name"],
);

const stylePatch = closedObject({
  name: string(253, 1),
  basedOn: { anyOf: [styleId, { type: "null" }] },
  next: { anyOf: [styleId, { type: "null" }] },
  quickStyle: boolean,
  uiPriority: integer(0, 99),
  paragraph: styleParaPatch,
  run: styleRunPatch,
});

const numberingLevelPatch = closedObject({
  format: {
    enum: [
      "decimal", "decimalZero", "upperRoman", "lowerRoman", "upperLetter",
      "lowerLetter", "ordinal", "bullet", "none",
    ],
  },
  text: string(100, 1),
  start: integer(0, 32767),
  alignment: { enum: ["left", "center", "right"] },
  indentLeftPt: number(0, 1584),
  hangingPt: number(0, 1584),
});

const paragraphSpacingPatch = closedObject({
  lineMultiple: number(0.1, 20),
  exactLinePt: number(0.1, 2000),
  beforePt: { type: ["number", "null"] },
  afterPt: { type: ["number", "null"] },
});

const paragraphDivider = closedObject({
  style: { enum: ["single", "double", "dotted", "dashed", "thinThickSmallGap"] },
  color: rgb,
  widthPt: number(0.1, 100),
  spacePt: number(0, 1000),
}, ["style", "color", "widthPt", "spacePt"]);

const chartData = closedObject({
  type: { enum: ["column", "bar", "line", "pie"] },
  title: string(200),
  categories: { type: "array", minItems: 1, maxItems: 100, items: string(200) },
  series: {
    type: "array",
    minItems: 1,
    maxItems: 24,
    items: closedObject({
      name: string(200),
      values: { type: "array", maxItems: 100, items: number() },
    }, ["name", "values"]),
  },
}, ["type", "categories", "series"]);

const smartArtData = closedObject({
  layout: { enum: ["process", "cycle", "hierarchy", "list"] },
  items: { type: "array", minItems: 1, maxItems: 50, items: string(500) },
}, ["layout", "items"]);

const smartArtTextFormat = closedObject({
  fontFamily: string(100, 1),
  fontSizePt: number(1, 1638),
  color: { type: "string", pattern: "^[0-9A-Fa-f]{6}$" },
  bold: boolean,
  italic: boolean,
  alignment: { enum: ["left", "center", "right"] },
}, ["fontFamily", "fontSizePt", "color", "bold", "italic", "alignment"]);

const pageLayoutPatch = closedObject({
  margins: closedObject({
    top: number(0, 22),
    right: number(0, 22),
    bottom: number(0, 22),
    left: number(0, 22),
  }),
  mirrorMargins: boolean,
  size: closedObject({ width: number(1, 22), height: number(1, 22) }, ["width", "height"]),
  orientation: { enum: ["portrait", "landscape"] },
  columns: integer(1, 12),
  columnSeparator: boolean,
  pageBorders: {
    anyOf: [
      closedObject({
        sz: number(0.125, 96),
        color: rgb,
        offsetFrom: { enum: ["text", "page"] },
      }),
      { type: "null" },
    ],
  },
});

const lineNumberingPatch = closedObject({
  enabled: boolean,
  countBy: integer(1, 100),
  restart: { enum: ["continuous", "newPage", "newSection"] },
  start: integer(0, 100000),
}, ["enabled"]);

const tableOperation = {
  anyOf: [
    { enum: ["deleteRow", "deleteCol", "deleteTable", "rowAbove", "rowBelow", "colLeft", "colRight"] },
    closedObject({ kind: { const: "cellShading" }, fill: { anyOf: [rgb, { type: "null" }] } }, ["kind", "fill"]),
    closedObject({ kind: { const: "cellVAlign" }, v: { enum: ["top", "center", "bottom"] } }, ["kind", "v"]),
    closedObject({
      kind: { const: "textWrapping" },
      wrapping: { enum: ["none", "around"] },
      xPx: number(-5000, 20000),
      yPx: number(-5000, 20000),
    }, ["kind", "wrapping", "xPx", "yPx"]),
  ],
};

// Table formatting. "none" is a real border style (Word's No Border, written
// w:val="nil") and is distinct from a null `border`, which REMOVES the edge so
// it inherits again — the schema has to admit both.
const tableBorderSpec = {
  anyOf: [
    closedObject({
      style: {
        enum: [
          "single", "thick", "double", "dotted", "dashed", "dotDash",
          "dotDotDash", "thinThickSmallGap", "triple", "wave", "none",
        ],
      },
      sz: number(1, 96),
      color: { anyOf: [rgb, { const: "auto" }] },
      space: number(0, 31),
    }, ["style"]),
    { type: "null" },
  ],
};

const tableLookToggles = closedObject({
  firstRow: boolean,
  lastRow: boolean,
  firstColumn: boolean,
  lastColumn: boolean,
  bandedRows: boolean,
  bandedCols: boolean,
});

const cellMarginsPt = {
  anyOf: [
    closedObject({ top: number(0, 720), left: number(0, 720), bottom: number(0, 720), right: number(0, 720) }),
    { type: "null" },
  ],
};

const NESTED_SCHEMAS: Record<string, JsonSchema> = {
  "formatRun.patch": runFormatPatch,
  "formatRange.patch": runFormatPatch,
  "setSpacing.patch": paragraphSpacingPatch,
  "setDivider.divider": { anyOf: [paragraphDivider, { type: "null" }] },
  "setPageLayout.patch": pageLayoutPatch,
  "setLineNumbering.patch": lineNumberingPatch,
  "insertCoverPage.content": closedObject({ title: string(1000, 1), subtitle: string(1000), author: string(500) }, ["title"]),
  "insertChart.chart": chartData,
  "setChartData.chart": chartData,
  "insertSmartArt.smartArt": smartArtData,
  "setSmartArtData.smartArt": smartArtData,
  "setSmartArtTextFormat.format": smartArtTextFormat,
  "tableOp.op": tableOperation,
  "setTableBorders.border": tableBorderSpec,
  "setTableBorders.edges": {
    type: "array",
    minItems: 1,
    maxItems: 8,
    items: { enum: ["top", "bottom", "left", "right", "insideH", "insideV", "tl2br", "tr2bl"] },
  },
  "setTableLook.look": tableLookToggles,
  "setTableCellMargins.margins": cellMarginsPt,
  "setTableWidth.value": number(0.1, 1584),
  "setTableColumnWidth.colIdx": integer(0, 200),
  "setTableColumnWidth.widthPt": number(1, 1584),
  "setTableHeaderRows.count": integer(0, 5000),
  "createStyle.style": styleSpec,
  "modifyStyle.styleId": styleId,
  "modifyStyle.patch": stylePatch,
  "deleteStyle.styleId": styleId,
  "setNumberingLevel.ilvl": { anyOf: [integer(0, 8), { type: "null" }] },
  "setNumberingLevel.patch": numberingLevelPatch,
  "setNumberingRestart.start": { anyOf: [integer(0, 32767), { type: "null" }] },
};

const ENUMS: Record<string, readonly unknown[]> = {
  "formatParagraph.align": ["left", "center", "right", "justify"],
  "setListType.listKind": ["bullet", "number", null],
  "tableOp.op": ["deleteRow", "deleteCol", "deleteTable", "rowAbove", "rowBelow", "colLeft", "colRight"],
  "insertBreak.breakKind": ["page", "column"],
  "insertShape.preset": ["line", "verticalLine", "rectangle", "roundedRectangle", "ellipse", "diamond", "textBox"],
  "adjustIndent.direction": [1, -1],
  "insertPageField.fieldKind": ["page", "pageOfTotal"],
  "setDropCap.mode": ["drop", "margin", null],
  "insertSectionBreak.breakType": ["nextPage", "continuous"],
  "insertCrossRef.refKind": ["text", "page"],
  "setListLevel.delta": [1, -1],
  "insertWordArt.preset": ["plain", "archUp", "archDown", "wave", "chevron"],
  "insertDateTimeField.dtKind": ["date", "time"],
  "setDrawingLineStyle.dash": ["solid", "dashed", "dotted"],
  "setImageWrap.mode": ["inline", "square", "topAndBottom", "none", "behind"],
  "setDrawingOrder.order": ["front", "back"],
  "ensureHeaderFooter.hfKind": ["header", "footer"],
  "setTableBorders.scope": ["cell", "table"],
  "setTableCellMargins.scope": ["cell", "table"],
  "setTableWidth.unit": ["pt", "pct", "auto"],
  "setTableLayout.layout": ["fixed", "autofit"],
};

function schemaForField(kind: Intent["kind"], field: string): JsonSchema {
  const nested = NESTED_SCHEMAS[`${kind}.${field}`];
  if (nested) return nested;
  const values = ENUMS[`${kind}.${field}`];
  if (values) return { enum: values };
  if (field === "blockRef" || field === "cellRef" || field === "afterBlockRef") return { type: "string", pattern: "^block:[0-9]+$" };
  if (field === "runRef") return { type: "string", pattern: "^run:[0-9]+$" };
  if (field === "objectRef") return { type: "string", pattern: "^object:[0-9]+:[0-9]+$" };
  if (field === "assetRef") return { type: "string", pattern: "^asset:.+$" };
  if (field === "at") return wirePosition;
  if (field === "ranges") {
    return {
      type: "array",
      items: {
        type: "object",
        properties: {
          blockRef: { type: "string", pattern: "^block:[0-9]+$" },
          runRef: { type: "string", pattern: "^run:[0-9]+$" },
          start: { type: "integer", minimum: 0 },
          end: { type: "integer", minimum: 1 },
        },
        required: ["blockRef", "runRef", "start", "end"],
        additionalProperties: false,
      },
    };
  }
  if (field === "marks") {
    return {
      type: "array",
      items: {
        type: "object",
        properties: { blockRef: { type: "string", pattern: "^block:[0-9]+$" }, glyph: { enum: ["ins", "del"] } },
        required: ["blockRef", "glyph"],
        additionalProperties: false,
      },
    };
  }
  if (field === "suggest" || field === "preservePageStart") return { type: "boolean" };
  if (field === "rows" || field === "cols") return integer(1, 50);
  if (field === "boundary") return integer(1, 200);
  if (field === "rowIdx") return integer(0, 5000);
  if (field === "pageDelta") return integer(-500, 500);
  if (field === "index" || field === "nodeIndex") return integer(0, 100000);
  if (field === "start" || field === "end") return integer(0);
  if (field === "degrees") return number();
  if (field === "opacity") return number(0, 1);
  if (field === "widthPx" || field === "heightPx") return number(1, kind === "resizeTableRow" ? 20000 : 5000);
  if (field === "xPx" || field === "yPx") return number(-5000, kind === "moveTable" ? 20000 : 5000);
  if (field === "deltaPx") return number(-5000, 5000);
  if (field === "renderedWidths") return { type: "array", maxItems: 200, items: number(0, 20000) };
  // updateFields carries one recomputed cached result per field in document
  // order; the caps match the operation's own validate in the core registry.
  if (field === "results") return { type: "array", maxItems: 20000, items: string(4096) };
  if (field === "color") {
    const nullable = kind === "setDrawingFill" || kind === "setSmartArtFill" || kind === "setDrawingLineStyle";
    const value = { type: "string", pattern: "^[0-9A-Fa-f]{6}$" };
    return nullable ? { anyOf: [value, { type: "null" }] } : value;
  }
  if (field === "styleId") return { type: ["string", "null"] };
  if (field === "text") {
    const max = kind === "insertText" ? 100000
      : kind === "insertWordArt" || kind === "setDrawingWordArtText" || kind === "setSmartArtNodeText" ? 500
        : 20000;
    return string(max, kind === "insertShape" || kind === "setSmartArtNodeText" ? 0 : 1);
  }
  if (field === "blocksXml") return string(2_000_000, 1);
  if (field === "url") return string(4000, 1);
  if (field === "alt") return string(1000);
  if (field === "picture") return string(100, 1);
  if (field === "instruction") return string(200, 1);
  if (field === "cachedResult") return string(1000);
  if (field === "name" && kind === "insertBookmarkRange") return { type: "string", pattern: "^[A-Za-z_][A-Za-z0-9_]{0,39}$" };
  return string();
}

export function agentOperationSchema(kind: Intent["kind"]): JsonSchema {
  const capability = AGENT_EDIT_CAPABILITIES[kind];
  const fields = [...capability.required, ...(capability.optional ?? [])];
  return {
    type: "object",
    properties: Object.fromEntries([["kind", { const: kind }], ...fields.map((field) => [field, schemaForField(kind, field)])]),
    required: ["kind", ...capability.required],
    additionalProperties: false,
  };
}

export function validateJsonSchema(value: unknown, schema: JsonSchema, path: string): string | null {
  if (Array.isArray(schema.anyOf)) {
    return schema.anyOf.some((candidate) => !validateJsonSchema(value, candidate as JsonSchema, path))
      ? null
      : `${path} does not match an allowed shape`;
  }
  if ("const" in schema && value !== schema.const) return `${path} must equal ${String(schema.const)}`;
  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) return `${path} has an unsupported value`;
  const types = Array.isArray(schema.type) ? schema.type : schema.type ? [schema.type] : [];
  if (types.length > 0) {
    const actual = value === null ? "null" : Array.isArray(value) ? "array" : Number.isInteger(value) ? "integer" : typeof value;
    const matches = types.includes(actual) || (actual === "integer" && types.includes("number"));
    if (!matches) return `${path} must be ${types.join(" or ")}`;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return `${path} must be finite`;
    if (typeof schema.minimum === "number" && value < schema.minimum) return `${path} is below its minimum`;
    if (typeof schema.maximum === "number" && value > schema.maximum) return `${path} is above its maximum`;
  }
  if (typeof value === "string") {
    if (typeof schema.minLength === "number" && value.length < schema.minLength) return `${path} is too short`;
    if (typeof schema.maxLength === "number" && value.length > schema.maxLength) return `${path} is too long`;
    if (typeof schema.pattern === "string" && !new RegExp(schema.pattern).test(value)) return `${path} has an invalid reference`;
  }
  if (Array.isArray(value) && schema.items && typeof schema.items === "object") {
    if (typeof schema.minItems === "number" && value.length < schema.minItems) return `${path} has too few items`;
    if (typeof schema.maxItems === "number" && value.length > schema.maxItems) return `${path} has too many items`;
    for (let index = 0; index < value.length; index++) {
      const error = validateJsonSchema(value[index], schema.items as JsonSchema, `${path}[${index}]`);
      if (error) return error;
    }
  }
  if (value && typeof value === "object" && !Array.isArray(value) && schema.type === "object") {
    const record = value as Record<string, unknown>;
    const properties = (schema.properties ?? {}) as Record<string, JsonSchema>;
    for (const required of (schema.required ?? []) as string[]) if (!(required in record)) return `${path} requires ${required}`;
    if (schema.additionalProperties === false) {
      const extra = Object.keys(record).find((key) => !(key in properties));
      if (extra) return `${path} does not allow ${extra}`;
    }
    for (const [key, child] of Object.entries(properties)) {
      if (!(key in record)) continue;
      const error = validateJsonSchema(record[key], child, `${path}.${key}`);
      if (error) return error;
    }
  }
  return null;
}

export function validateAgentOperationShape(input: unknown): string | null {
  if (!input || typeof input !== "object" || Array.isArray(input)) return "Operation must be an object";
  const kind = (input as { kind?: unknown }).kind;
  if (typeof kind !== "string" || !(kind in AGENT_EDIT_CAPABILITIES)) return "Unknown edit operation";
  return validateJsonSchema(input, agentOperationSchema(kind as Intent["kind"]), "operation");
}

export function agentCapabilities(category?: AgentEditCapability["category"], requestedKind?: Intent["kind"]): Array<AgentEditCapability & { kind: Intent["kind"]; inputSchema: Record<string, unknown> }> {
  return INTENT_KINDS
    .filter((kind) => (!category || AGENT_EDIT_CAPABILITIES[kind].category === category) && (!requestedKind || kind === requestedKind))
    .map((kind) => {
      const capability = AGENT_EDIT_CAPABILITIES[kind];
      return {
        kind,
        ...capability,
        inputSchema: agentOperationSchema(kind),
      };
    });
}
