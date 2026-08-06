import { checkboxStateElement, toggleCheckbox } from "../checkbox.js";
import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement } from "../xml.js";
import { insertTableAfter } from "./blocks.js";
import { setListType } from "./lists.js";
import {
  NUMBER_FORMATS,
  continueNumberingAt,
  restartNumberingAt,
  setNumberingLevelAt,
  type LevelPatch,
} from "./numbering.js";
import {
  createStyle,
  deleteStyle,
  modifyStyle,
  type StyleParaPatch,
  type StyleRunPatch,
  type StyleSpec,
  type StylePatch,
} from "./styles.js";
import { suggestMeta } from "./suggest.js";
import { applyFieldResults } from "./update-fields.js";
import {
  CELL_SCOPE_EDGES,
  TABLE_BORDER_STYLES,
  TABLE_SCOPE_EDGES,
  resizeTableRow,
  setTableBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableHeaderRows,
  setTableLayoutMode,
  setTableLook,
  setTableStyle,
  setTableWidth,
  type CellMarginsPt,
  type TableBorderEdge,
  type TableBorderSpec,
  type TableLookToggles,
} from "./tables.js";

/**
 * The operation registry: ONE declaration per edit operation, read by every
 * surface that used to carry its own hand-written copy.
 *
 * Adding an operation used to cost four to six coordinated edits — a core
 * mutation, an editor dispatch, a wire intent variant in TWO packages, a
 * transform/validate case, an agent capability row, an agent id-budget case.
 * Each of those is guarded by a compile-time exhaustiveness gate, so nothing
 * could silently go missing; the cost was the coordination, not the risk. A
 * registered operation is declared once here, and the gates then consume the
 * registry instead of a second hand-maintained list. The gates are NOT
 * relaxed: an operation that is not registered still has to satisfy every one
 * of them by hand.
 *
 * WHAT MAY BE REGISTERED. The declaration shape below is deliberately narrow.
 * A registered operation must be:
 *
 *  1. ADDRESSED BY STABLE ID — a run, a paragraph, or a table cell paragraph.
 *     The address is the whole of its collaborative "honest no-op" predicate:
 *     when the id does not resolve in the room, the operation applies as a
 *     clean rejection everywhere rather than mutating anything locally.
 *  2. POSITION-STABLE — it moves no run's text, so a concurrent intent's
 *     offsets need no remapping and its transform is identity. The
 *     text-shaped intents (insertText, deleteText, splitParagraph,
 *     formatRange) each carry bespoke transform logic and stay hand-written.
 *  3. WITHOUT A WIRE INVERSE — collaborative undo skips it, which is the
 *     status quo for every intent except the four in @wordinweb/collab's
 *     invert.ts. The local path still takes a history checkpoint.
 *
 * Those three preconditions ARE the operation's undo and transform
 * classification. They are stated once, here, instead of repeated per
 * operation, because every operation the registry admits shares them; an
 * operation that breaks one of them does not belong in the registry yet.
 */

/** A stable id from the core side table (the wire's addressing unit). */
export type StableId = number;

/**
 * How a registered operation names its target on the wire.
 *
 * "document" scopes an operation to the whole document, which means it has no
 * stable id and therefore none of the honest-no-op protection an id gives: a
 * document address ALWAYS resolves. Such an operation has to supply its own
 * rejection predicate from the payload — see updateFields below, whose result
 * count must match the document's field count — and it must carry every value
 * a replica cannot re-derive identically.
 */
export type OperationAddress = "run" | "block" | "cell" | "document";

/** An operation addressed by a stable id, as opposed to document-scoped. */
export type AddressedOperation = Exclude<OperationAddress, "document">;

/** The wire field carrying the address, per addressed kind. */
export const ADDRESS_WIRE_FIELD = {
  run: "runId",
  block: "blockId",
  cell: "cellParagraphId",
} as const satisfies Record<AddressedOperation, string>;

/** The agent-facing reference field, per addressed kind. Agents address content
 * with opaque strings ("run:12"), never raw ids. */
export const ADDRESS_AGENT_FIELD = {
  run: "runRef",
  block: "blockRef",
  cell: "cellRef",
} as const satisfies Record<AddressedOperation, string>;

/**
 * Agent capability category. Mirrors @wordinweb/agent's
 * AgentEditCapability["category"]; the agent's typed capability map is what
 * checks the two stay equal.
 */
export type OperationCategory =
  | "text"
  | "paragraph"
  | "review"
  | "table"
  | "insert"
  | "drawing"
  | "math"
  | "document";

/** An agent-facing payload field beyond the address reference. Order is the
 * order the agent's JSON schema lists them in. */
export interface OperationField {
  name: string;
  /** Absent means required. */
  optional?: true;
}

/** What an address resolved to in the document. */
export interface OperationTarget {
  /** The addressed element: the w:r, the w:p, or the w:tbl owning the cell. */
  el: XmlElement;
  /** The first w:t under a run, or the first text descendant of a paragraph.
   * Null when the target carries no text (and for cell addressing). */
  t: XmlElement | null;
  /** The parsed run for a run address, null for the other two. Needed by
   * operations whose state is reachable only through the model — the checkbox
   * marker parsing hangs on a run's content — rather than from the XML. */
  run: Run | null;
  /**
   * For cell addressing: the addressed w:p itself. `el` deliberately widens a
   * cell address to the owning w:tbl, which is what a table-scoped operation
   * wants; a CELL-scoped one (per-edge borders, a margin override) needs the
   * one cell the caret is in, and the paragraph is how it gets there. Null
   * for run and block addressing, whose addressed element IS `el`.
   */
  cellParagraph: XmlElement | null;
}

export interface OperationContext<Payload> {
  doc: DocxDocument;
  target: OperationTarget;
  payload: Payload;
}

/**
 * The caller-supplied part of a wire payload: everything except the address
 * field the dispatcher fills in and the carried ids it allocates.
 */
export type OperationArgs<Payload> = Omit<
  Payload,
  "runId" | "blockId" | "cellParagraphId" | "nodeIds"
>;

export interface OperationDefinition<Kind extends string, Payload> {
  kind: Kind;
  address: OperationAddress;
  /** Agent capability row: how the operation is grouped and described. */
  category: OperationCategory;
  description: string;
  /** Agent-facing fields beyond the address reference, in schema order. */
  fields: readonly OperationField[];
  /**
   * How many fresh stable ids the mutation needs for the id-tracked nodes it
   * creates. Omit when it creates none. Every producer (the editor, the React
   * host, the agent compiler) reads this instead of repeating the arithmetic.
   */
  nodeIds?: (args: OperationArgs<Payload>) => number;
  /**
   * Reject a malformed payload before it is sequenced. Runs on both sides of
   * the wire, so it must be a pure function of the payload.
   */
  validate?: (payload: Payload) => string | null;
  /**
   * The XML mutation, headless and from a resolved target. False means a
   * clean no-op: the operation applied nothing and every replica agrees.
   */
  apply: (ctx: OperationContext<Payload>) => boolean;
}

/**
 * Declare one operation. Curried so the wire payload is written explicitly
 * and the kind is still inferred as a literal:
 *
 *     defineOperation<{ runId: StableId; rows: number }>()({ kind: "…", … })
 */
export function defineOperation<Payload>() {
  return <Kind extends string>(
    definition: OperationDefinition<Kind, Payload>,
  ): OperationDefinition<Kind, Payload> => definition;
}

// ---------------------------------------------------------------------------
// Registered operations
// ---------------------------------------------------------------------------

/** Turn a paragraph into a bullet/numbered list item, or clear its list
 * formatting (listKind null). Mutates w:pPr numbering in place. With `suggest`
 * the change is TRACKED (w:pPrChange) instead of applied outright; the author
 * and date travel in the payload so every replica writes the same XML. */
const setListTypeOperation = defineOperation<{
  blockId: StableId;
  listKind: "bullet" | "number" | null;
  suggest?: { author: string; date: string };
}>()({
  kind: "setListType",
  address: "block",
  category: "paragraph",
  description: "Set or clear paragraph list formatting.",
  fields: [{ name: "listKind" }, { name: "suggest", optional: true }],
  validate: ({ suggest }) => {
    if (suggest === undefined) return null;
    if (typeof suggest.author !== "string" || suggest.author.length > 100) return "setListType: bad author";
    if (typeof suggest.date !== "string" || suggest.date.length > 40) return "setListType: bad date";
    return null;
  },
  // setListType resolves the paragraph by walking UP from a target, so pass a
  // descendant w:t when the paragraph has one and the paragraph itself
  // otherwise.
  apply: ({ doc, target, payload }) =>
    setListType(doc, [target.t ?? target.el], payload.listKind, suggestMeta(doc, payload.suggest)),
});

/** Insert a rows×cols table after the paragraph containing the anchor run. */
const insertTableOperation = defineOperation<{
  runId: StableId;
  rows: number;
  cols: number;
  nodeIds: StableId[];
}>()({
  kind: "insertTable",
  address: "run",
  category: "insert",
  description: "Insert a table.",
  fields: [{ name: "rows" }, { name: "cols" }],
  // A cell is a paragraph inside a run inside a row; the spare 8 cover the
  // table element and the trailing paragraph.
  nodeIds: ({ rows, cols }) =>
    Number.isInteger(rows) && Number.isInteger(cols) && rows > 0 && cols > 0
      ? rows * cols * 2 + 8
      : 8,
  validate: ({ rows, cols }) => {
    const okDim = (v: unknown) => typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= 50;
    return okDim(rows) && okDim(cols) ? null : "insertTable: bad dimensions";
  },
  apply: ({ doc, target, payload }) =>
    target.t ? insertTableAfter(doc, target.t, payload.rows, payload.cols) : false,
});

/** Set the height of one table row, addressed by a paragraph in the table. */
const resizeTableRowOperation = defineOperation<{
  cellParagraphId: StableId;
  rowIdx: number;
  heightPx: number;
}>()({
  kind: "resizeTableRow",
  address: "cell",
  category: "table",
  description: "Resize a table row.",
  fields: [{ name: "rowIdx" }, { name: "heightPx" }],
  validate: ({ rowIdx, heightPx }) => {
    if (!Number.isInteger(rowIdx) || rowIdx < 0 || rowIdx > 5000) return "resizeTableRow: bad row";
    if (typeof heightPx !== "number" || !Number.isFinite(heightPx) || heightPx < 1 || heightPx > 20000) {
      return "resizeTableRow: bad height";
    }
    return null;
  },
  apply: ({ doc, target, payload }) => resizeTableRow(doc, target.el, payload.rowIdx, payload.heightPx),
});

/** Flip the checked state of the checkbox content control a run carries. */
const toggleCheckboxOperation = defineOperation<{ runId: StableId }>()({
  kind: "toggleCheckbox",
  address: "run",
  category: "text",
  description: "Toggle a checkbox content control.",
  fields: [],
  // A run that carries no checkbox is a clean no-op, like an unresolvable
  // address: the ballot glyph of a legacy form field is a synthetic field
  // result with no w:t, so a null `t` is a lookup mode rather than a failure.
  apply: ({ doc, target }) => {
    const cbEl = checkboxStateElement(target.run ?? undefined, target.t);
    if (!cbEl) return false;
    toggleCheckbox(doc, cbEl);
    return true;
  },
});

/**
 * Write recomputed cached results into the document's fields — Word's F9.
 *
 * The results are CARRIED rather than recomputed per replica, because the
 * values a field update produces are not replica-independent: PAGE, NUMPAGES
 * and PAGEREF all come out of a layout, and layout depends on the host's font
 * metrics. Two browsers can paginate the same document differently, so a
 * replica that recomputed would install different text. This is the provenance
 * pattern (edit/provenance.ts) applied to a value that varies by font stack
 * rather than by clock.
 *
 * The result count is the operation's rejection predicate, standing in for the
 * stable id a document-scoped operation does not have: a replica whose field
 * count has moved under a concurrent edit applies nothing, and every replica in
 * that position rejects identically.
 *
 * Result runs are NOT created here (createResultRuns: false). A field that has
 * never been evaluated has no run to write into, and adding one would create an
 * id-tracked node this operation has no carried id for. Those fields keep their
 * empty result in a room and are filled in by a local update.
 */
const updateFieldsOperation = defineOperation<{ results: string[] }>()({
  kind: "updateFields",
  address: "document",
  category: "document",
  description: "Write recomputed cached results into the document's fields, one per field in document order.",
  fields: [{ name: "results" }],
  validate: ({ results }) => {
    if (!Array.isArray(results)) return "updateFields: results not an array";
    if (results.length > 20000) return "updateFields: too many results";
    for (const r of results) {
      if (typeof r !== "string") return "updateFields: result not a string";
      if (r.length > 4096) return "updateFields: result too long";
    }
    return null;
  },
  apply: ({ doc, payload }) => applyFieldResults(doc, payload.results, { createResultRuns: false }),
});

// ---------------------------------------------------------------------------
// Style definitions
// ---------------------------------------------------------------------------

/**
 * The three operations below edit styles.xml rather than content, so they are
 * DOCUMENT-scoped and carry no stable id. Their rejection predicate is the
 * styleId itself: createStyle rejects an id styles.xml already declares,
 * modifyStyle and deleteStyle reject one it does not. styles.xml is sequenced
 * state like any other, so every replica in the same position decides the same
 * way — and unlike updateFields, none of these carry a value a replica could
 * derive differently, because every property comes out of the payload.
 *
 * They are position-stable in the sense the registry requires: no run's text
 * moves. deleteStyle does re-point w:pStyle/w:rStyle values, which changes
 * paragraph XML without changing any offset, so a concurrent text intent still
 * transforms as identity.
 */

const STYLE_ID = /^[A-Za-z0-9\-_]{1,253}$/;

function badStyleId(id: unknown, what: string): string | null {
  return typeof id === "string" && STYLE_ID.test(id) ? null : `${what}: bad styleId`;
}

/** A number within bounds, or absent/null when the field is clearable. */
function badNumber(
  value: unknown,
  what: string,
  min: number,
  max: number,
  nullable: boolean,
): string | null {
  if (value === undefined) return null;
  if (value === null) return nullable ? null : `${what} is not clearable`;
  if (typeof value !== "number" || !Number.isFinite(value) || value < min || value > max) {
    return `${what} out of range`;
  }
  return null;
}

function badParaPatch(patch: unknown, what: string): string | null {
  if (patch === undefined) return null;
  if (!patch || typeof patch !== "object" || Array.isArray(patch)) return `${what}: bad paragraph`;
  const p = patch as StyleParaPatch;
  const allowed = [
    "alignment", "spacingBeforePt", "spacingAfterPt", "lineMultiple",
    "indentLeftPt", "indentFirstLinePt", "keepNext", "outlineLevel",
  ];
  for (const key of Object.keys(p)) {
    if (!allowed.includes(key)) return `${what}: unknown paragraph property ${key}`;
  }
  if (p.alignment !== undefined && p.alignment !== null &&
      !["left", "center", "right", "both"].includes(p.alignment)) {
    return `${what}: bad alignment`;
  }
  if (p.keepNext !== undefined && p.keepNext !== null && typeof p.keepNext !== "boolean") {
    return `${what}: bad keepNext`;
  }
  return (
    badNumber(p.spacingBeforePt, `${what}: spacingBeforePt`, 0, 1584, true) ??
    badNumber(p.spacingAfterPt, `${what}: spacingAfterPt`, 0, 1584, true) ??
    badNumber(p.lineMultiple, `${what}: lineMultiple`, 0.1, 132, true) ??
    badNumber(p.indentLeftPt, `${what}: indentLeftPt`, -1584, 1584, true) ??
    badNumber(p.indentFirstLinePt, `${what}: indentFirstLinePt`, -1584, 1584, true) ??
    badNumber(p.outlineLevel, `${what}: outlineLevel`, 0, 8, true)
  );
}

function badRunPatch(patch: unknown, what: string): string | null {
  if (patch === undefined) return null;
  if (!patch || typeof patch !== "object" || Array.isArray(patch)) return `${what}: bad run`;
  const r = patch as StyleRunPatch;
  const allowed = [
    "bold", "italic", "underline", "strike", "color", "highlight",
    "fontSizePt", "fontFamily", "characterStyleId", "verticalAlign",
  ];
  for (const key of Object.keys(r)) {
    if (!allowed.includes(key)) return `${what}: unknown run property ${key}`;
  }
  for (const key of ["bold", "italic", "underline", "strike"] as const) {
    if (r[key] !== undefined && typeof r[key] !== "boolean") return `${what}: bad ${key}`;
  }
  if (r.color !== undefined && r.color !== null && !/^#?[0-9A-Fa-f]{6}$/.test(r.color)) {
    return `${what}: bad color`;
  }
  if (r.highlight !== undefined && r.highlight !== null &&
      !(typeof r.highlight === "string" && /^[a-zA-Z]{1,20}$/.test(r.highlight))) {
    return `${what}: bad highlight`;
  }
  if (r.fontFamily !== undefined &&
      !(typeof r.fontFamily === "string" && r.fontFamily.length > 0 && r.fontFamily.length <= 64)) {
    return `${what}: bad fontFamily`;
  }
  if (r.verticalAlign !== undefined && r.verticalAlign !== null &&
      !["superscript", "subscript"].includes(r.verticalAlign)) {
    return `${what}: bad verticalAlign`;
  }
  if (r.characterStyleId !== undefined && r.characterStyleId !== null) {
    const bad = badStyleId(r.characterStyleId, what);
    if (bad) return bad;
  }
  return badNumber(r.fontSizePt, `${what}: fontSizePt`, 1, 1638, false);
}

function badStyleName(name: unknown, what: string): string | null {
  return typeof name === "string" && name.length > 0 && name.length <= 253
    ? null
    : `${what}: bad name`;
}

/** Add a paragraph or character style definition to styles.xml. */
const createStyleOperation = defineOperation<{ style: StyleSpec }>()({
  kind: "createStyle",
  address: "document",
  category: "document",
  description: "Create a paragraph or character style definition.",
  fields: [{ name: "style" }],
  validate: ({ style }) => {
    if (!style || typeof style !== "object" || Array.isArray(style)) return "createStyle: bad style";
    if (style.type !== "paragraph" && style.type !== "character") return "createStyle: bad type";
    if (style.type === "character" && style.paragraph) {
      return "createStyle: a character style has no paragraph properties";
    }
    if (style.quickStyle !== undefined && typeof style.quickStyle !== "boolean") {
      return "createStyle: bad quickStyle";
    }
    for (const key of ["basedOn", "next"] as const) {
      const value = style[key];
      if (value === undefined || value === null) continue;
      const bad = badStyleId(value, `createStyle: ${key}`);
      if (bad) return bad;
    }
    return (
      badStyleId(style.styleId, "createStyle") ??
      badStyleName(style.name, "createStyle") ??
      badNumber(style.uiPriority, "createStyle: uiPriority", 0, 99, false) ??
      badParaPatch(style.paragraph, "createStyle") ??
      badRunPatch(style.run, "createStyle")
    );
  },
  apply: ({ doc, payload }) => createStyle(doc, payload.style),
});

/** Patch an existing style definition in place. */
const modifyStyleOperation = defineOperation<{ styleId: string; patch: StylePatch }>()({
  kind: "modifyStyle",
  address: "document",
  category: "document",
  description: "Change an existing style definition.",
  fields: [{ name: "styleId" }, { name: "patch" }],
  validate: ({ styleId, patch }) => {
    const badId = badStyleId(styleId, "modifyStyle");
    if (badId) return badId;
    if (!patch || typeof patch !== "object" || Array.isArray(patch)) return "modifyStyle: bad patch";
    if (Object.keys(patch).length === 0) return "modifyStyle: empty patch";
    if (patch.name !== undefined) {
      const bad = badStyleName(patch.name, "modifyStyle");
      if (bad) return bad;
    }
    if (patch.quickStyle !== undefined && typeof patch.quickStyle !== "boolean") {
      return "modifyStyle: bad quickStyle";
    }
    for (const key of ["basedOn", "next"] as const) {
      const value = patch[key];
      if (value === undefined || value === null) continue;
      const bad = badStyleId(value, `modifyStyle: ${key}`);
      if (bad) return bad;
    }
    return (
      badNumber(patch.uiPriority, "modifyStyle: uiPriority", 0, 99, false) ??
      badParaPatch(patch.paragraph, "modifyStyle") ??
      badRunPatch(patch.run, "modifyStyle")
    );
  },
  apply: ({ doc, payload }) => modifyStyle(doc, payload.styleId, payload.patch),
});

/** Remove a style definition, re-pointing its users at the surviving parent. */
const deleteStyleOperation = defineOperation<{ styleId: string }>()({
  kind: "deleteStyle",
  address: "document",
  category: "document",
  description: "Delete a style definition; content using it falls back to the style it was based on.",
  fields: [{ name: "styleId" }],
  validate: ({ styleId }) => badStyleId(styleId, "deleteStyle"),
  apply: ({ doc, payload }) => deleteStyle(doc, payload.styleId),
});

// ---------------------------------------------------------------------------
// Multilevel numbering
// ---------------------------------------------------------------------------

/**
 * Numbering operations address a PARAGRAPH in the list rather than the numId,
 * so they get the stable id an addressed operation needs; the numId comes off
 * the paragraph's own w:numPr, which every replica reads identically.
 */

/** Change how one level of a multilevel list is labelled and indented. */
const setNumberingLevelOperation = defineOperation<{
  blockId: StableId;
  ilvl: number | null;
  patch: LevelPatch;
}>()({
  kind: "setNumberingLevel",
  address: "block",
  category: "paragraph",
  description: "Change a list level's number format, label text, or indent.",
  fields: [{ name: "ilvl" }, { name: "patch" }],
  validate: ({ ilvl, patch }) => {
    if (ilvl !== null && (!Number.isInteger(ilvl) || ilvl < 0 || ilvl > 8)) {
      return "setNumberingLevel: bad ilvl";
    }
    if (!patch || typeof patch !== "object" || Array.isArray(patch)) {
      return "setNumberingLevel: bad patch";
    }
    const allowed = ["format", "text", "start", "alignment", "indentLeftPt", "hangingPt"];
    for (const key of Object.keys(patch)) {
      if (!allowed.includes(key)) return `setNumberingLevel: unknown property ${key}`;
    }
    if (Object.keys(patch).length === 0) return "setNumberingLevel: empty patch";
    if (patch.format !== undefined && !NUMBER_FORMATS.includes(patch.format)) {
      return "setNumberingLevel: bad format";
    }
    if (patch.text !== undefined &&
        !(typeof patch.text === "string" && patch.text.length > 0 && patch.text.length <= 100)) {
      return "setNumberingLevel: bad text";
    }
    if (patch.alignment !== undefined && !["left", "center", "right"].includes(patch.alignment)) {
      return "setNumberingLevel: bad alignment";
    }
    return (
      badNumber(patch.start, "setNumberingLevel: start", 0, 32767, false) ??
      badNumber(patch.indentLeftPt, "setNumberingLevel: indentLeftPt", 0, 1584, false) ??
      badNumber(patch.hangingPt, "setNumberingLevel: hangingPt", 0, 1584, false)
    );
  },
  apply: ({ doc, target, payload }) =>
    setNumberingLevelAt(doc, target.el, payload.ilvl, payload.patch),
});

/**
 * Restart the list at a paragraph (start a number), or make it continue the
 * preceding list of the same definition (start null).
 *
 * The restart path creates a fresh w:num, whose numId is max+1 over
 * numbering.xml — derived from sequenced state, so every replica allocates the
 * same one without an id having to travel.
 */
const setNumberingRestartOperation = defineOperation<{
  blockId: StableId;
  start: number | null;
}>()({
  kind: "setNumberingRestart",
  address: "block",
  category: "paragraph",
  description: "Restart list numbering at a paragraph, or continue the preceding list.",
  fields: [{ name: "start" }],
  validate: ({ start }) =>
    start === null ? null : badNumber(start, "setNumberingRestart: start", 0, 32767, false),
  apply: ({ doc, target, payload }) =>
    payload.start === null
      ? continueNumberingAt(doc, target.el)
      : restartNumberingAt(doc, target.el, payload.start),
});

// ---------------------------------------------------------------------------
// Table formatting
// ---------------------------------------------------------------------------

/**
 * The scoped operations below address a paragraph inside a cell. A TABLE-scoped
 * one uses `target.el` (the owning w:tbl); a CELL-scoped one uses
 * `target.cellParagraph`, which is the only handle on the single cell the user
 * is in. Passing the paragraph to a `tables.ts` helper is the same convention
 * the local editor path uses, where the caret's w:t is passed instead.
 */
function cellAnchor(target: OperationTarget): XmlElement | null {
  return target.cellParagraph;
}

/** Set or clear per-edge borders on one cell or on the whole table. */
const setTableBordersOperation = defineOperation<{
  cellParagraphId: StableId;
  scope: "cell" | "table";
  edges: TableBorderEdge[];
  border: TableBorderSpec | null;
}>()({
  kind: "setTableBorders",
  address: "cell",
  category: "table",
  description: "Set or clear table or cell borders, per edge.",
  fields: [{ name: "scope" }, { name: "edges" }, { name: "border" }],
  validate: ({ scope, edges, border }) => {
    if (scope !== "cell" && scope !== "table") return "setTableBorders: bad scope";
    if (!Array.isArray(edges) || edges.length === 0 || edges.length > 8) {
      return "setTableBorders: bad edges";
    }
    const allowed = scope === "cell" ? CELL_SCOPE_EDGES : TABLE_SCOPE_EDGES;
    if (edges.some((e) => !allowed.includes(e))) return `setTableBorders: edge not valid at ${scope} scope`;
    if (border === null) return null;
    if (!border || !TABLE_BORDER_STYLES.includes(border.style)) return "setTableBorders: bad style";
    if (border.sz !== undefined && (!Number.isFinite(border.sz) || border.sz < 1 || border.sz > 96)) {
      return "setTableBorders: bad sz";
    }
    if (border.space !== undefined && (!Number.isFinite(border.space) || border.space < 0 || border.space > 31)) {
      return "setTableBorders: bad space";
    }
    if (border.color !== undefined && !/^(#?[0-9A-Fa-f]{6}|auto)$/.test(border.color)) {
      return "setTableBorders: bad color";
    }
    return null;
  },
  apply: ({ doc, target, payload }) => {
    const anchor = cellAnchor(target);
    return anchor
      ? setTableBorders(doc, anchor, payload.scope, payload.edges, payload.border)
      : false;
  },
});

/** Apply a named table style, or with null remove the reference. */
const setTableStyleOperation = defineOperation<{
  cellParagraphId: StableId;
  styleId: string | null;
}>()({
  kind: "setTableStyle",
  address: "cell",
  category: "table",
  description: "Apply a named table style to a table.",
  fields: [{ name: "styleId" }],
  validate: ({ styleId }) => {
    if (styleId === null) return null;
    return typeof styleId === "string" && styleId.length > 0 && styleId.length <= 253
      ? null
      : "setTableStyle: bad styleId";
  },
  apply: ({ doc, target, payload }) => setTableStyle(doc, target.el, payload.styleId),
});

/** Set some of Word's six table-style option toggles (w:tblLook). */
const setTableLookOperation = defineOperation<{
  cellParagraphId: StableId;
  look: Partial<TableLookToggles>;
}>()({
  kind: "setTableLook",
  address: "cell",
  category: "table",
  description: "Toggle which table style options apply (first/last row and column, banding).",
  fields: [{ name: "look" }],
  validate: ({ look }) => {
    if (!look || typeof look !== "object" || Array.isArray(look)) return "setTableLook: bad look";
    const keys: (keyof TableLookToggles)[] = [
      "firstRow",
      "lastRow",
      "firstColumn",
      "lastColumn",
      "bandedRows",
      "bandedCols",
    ];
    const entries = Object.entries(look);
    if (entries.length === 0) return "setTableLook: empty look";
    for (const [key, value] of entries) {
      if (!keys.includes(key as keyof TableLookToggles)) return `setTableLook: unknown toggle ${key}`;
      if (typeof value !== "boolean") return `setTableLook: ${key} is not a boolean`;
    }
    return null;
  },
  apply: ({ doc, target, payload }) => setTableLook(doc, target.el, payload.look),
});

/** Set the table's preferred width (points, percent, or auto). */
const setTableWidthOperation = defineOperation<{
  cellParagraphId: StableId;
  unit: "pt" | "pct" | "auto";
  value?: number;
}>()({
  kind: "setTableWidth",
  address: "cell",
  category: "table",
  description: "Set a table's preferred width in points, as a percent, or to auto.",
  fields: [{ name: "unit" }, { name: "value", optional: true }],
  validate: ({ unit, value }) => {
    if (unit !== "pt" && unit !== "pct" && unit !== "auto") return "setTableWidth: bad unit";
    if (unit === "auto") return null;
    if (typeof value !== "number" || !Number.isFinite(value)) return "setTableWidth: value required";
    const max = unit === "pct" ? 100 : 22 * 72;
    return value > 0 && value <= max ? null : "setTableWidth: value out of range";
  },
  apply: ({ doc, target, payload }) =>
    setTableWidth(doc, target.el, payload.unit, payload.value ?? 0),
});

/** Set one grid column to an exact width in points. */
const setTableColumnWidthOperation = defineOperation<{
  cellParagraphId: StableId;
  colIdx: number;
  widthPt: number;
}>()({
  kind: "setTableColumnWidth",
  address: "cell",
  category: "table",
  description: "Set one table column to an exact width in points.",
  fields: [{ name: "colIdx" }, { name: "widthPt" }],
  validate: ({ colIdx, widthPt }) => {
    if (!Number.isInteger(colIdx) || colIdx < 0 || colIdx > 200) return "setTableColumnWidth: bad column";
    if (typeof widthPt !== "number" || !Number.isFinite(widthPt) || widthPt < 1 || widthPt > 22 * 72) {
      return "setTableColumnWidth: bad width";
    }
    return null;
  },
  apply: ({ doc, target, payload }) =>
    setTableColumnWidth(doc, target.el, payload.colIdx, payload.widthPt),
});

/** Switch a table between fixed column widths and autofit-to-contents. */
const setTableLayoutOperation = defineOperation<{
  cellParagraphId: StableId;
  layout: "fixed" | "autofit";
  renderedWidths?: number[];
}>()({
  kind: "setTableLayout",
  address: "cell",
  category: "table",
  description: "Switch a table between fixed column widths and autofit.",
  // renderedWidths is the caller's MEASURED columns; it rides as data so
  // every replica freezes the same widths when switching to fixed.
  fields: [{ name: "layout" }, { name: "renderedWidths", optional: true }],
  validate: ({ layout, renderedWidths }) => {
    if (layout !== "fixed" && layout !== "autofit") return "setTableLayout: bad layout";
    if (renderedWidths === undefined) return null;
    if (!Array.isArray(renderedWidths) || renderedWidths.length > 200) {
      return "setTableLayout: bad renderedWidths";
    }
    return renderedWidths.every((n) => typeof n === "number" && Number.isFinite(n) && n >= 0 && n <= 20000)
      ? null
      : "setTableLayout: bad renderedWidths";
  },
  apply: ({ doc, target, payload }) =>
    setTableLayoutMode(doc, target.el, payload.layout, payload.renderedWidths),
});

/** Set the table's default cell margins, or one cell's override. */
const setTableCellMarginsOperation = defineOperation<{
  cellParagraphId: StableId;
  scope: "cell" | "table";
  margins: CellMarginsPt | null;
}>()({
  kind: "setTableCellMargins",
  address: "cell",
  category: "table",
  description: "Set table default cell margins, or one cell's override, in points.",
  fields: [{ name: "scope" }, { name: "margins" }],
  validate: ({ scope, margins }) => {
    if (scope !== "cell" && scope !== "table") return "setTableCellMargins: bad scope";
    if (margins === null) return null;
    if (!margins || typeof margins !== "object" || Array.isArray(margins)) {
      return "setTableCellMargins: bad margins";
    }
    const sides = ["top", "left", "bottom", "right"];
    const entries = Object.entries(margins);
    if (entries.length === 0) return "setTableCellMargins: empty margins";
    for (const [side, value] of entries) {
      if (!sides.includes(side)) return `setTableCellMargins: unknown side ${side}`;
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 720) {
        return `setTableCellMargins: bad ${side}`;
      }
    }
    return null;
  },
  apply: ({ doc, target, payload }) => {
    const anchor = cellAnchor(target);
    return anchor ? setTableCellMargins(doc, anchor, payload.scope, payload.margins) : false;
  },
});

/** Mark the first N rows as the repeating header band. */
const setTableHeaderRowsOperation = defineOperation<{
  cellParagraphId: StableId;
  count: number;
}>()({
  kind: "setTableHeaderRows",
  address: "cell",
  category: "table",
  description: "Repeat the first N rows as a header band on every page.",
  fields: [{ name: "count" }],
  validate: ({ count }) =>
    Number.isInteger(count) && count >= 0 && count <= 5000 ? null : "setTableHeaderRows: bad count",
  apply: ({ doc, target, payload }) => setTableHeaderRows(doc, target.el, payload.count),
});

const OPERATIONS = [
  setListTypeOperation,
  insertTableOperation,
  resizeTableRowOperation,
  toggleCheckboxOperation,
  updateFieldsOperation,
  createStyleOperation,
  modifyStyleOperation,
  deleteStyleOperation,
  setNumberingLevelOperation,
  setNumberingRestartOperation,
  setTableBordersOperation,
  setTableStyleOperation,
  setTableLookOperation,
  setTableWidthOperation,
  setTableColumnWidthOperation,
  setTableLayoutOperation,
  setTableCellMarginsOperation,
  setTableHeaderRowsOperation,
] as const;

// ---------------------------------------------------------------------------
// Derived surfaces
// ---------------------------------------------------------------------------

type AnyOperationDefinition = (typeof OPERATIONS)[number];

/** Distributes: one body type per registered operation. */
type BodyOf<D> = D extends OperationDefinition<infer Kind, infer Payload>
  ? { kind: Kind } & Payload
  : never;

export type RegisteredOperationKind = AnyOperationDefinition["kind"];

/**
 * The wire body of every registered operation. @wordinweb/collab intersects
 * this with its IntentBase bookkeeping to form the Intent variants, so the
 * payload is declared once and the wire shape follows.
 */
export type RegisteredOperationBody = BodyOf<AnyOperationDefinition>;

/** The body of one registered operation. */
export type RegisteredOperationBodyFor<Kind extends RegisteredOperationKind> = Extract<
  RegisteredOperationBody,
  { kind: Kind }
>;

/** The caller-supplied arguments of one registered operation. */
export type RegisteredOperationArgs<Kind extends RegisteredOperationKind> = OperationArgs<
  Omit<RegisteredOperationBodyFor<Kind>, "kind">
>;

const BY_KIND = new Map<string, AnyOperationDefinition>(OPERATIONS.map((op) => [op.kind, op]));

export const REGISTERED_OPERATION_KINDS: readonly RegisteredOperationKind[] = Object.freeze(
  OPERATIONS.map((op) => op.kind),
);

/** The definition for a kind, or undefined when the kind is not registered
 * (every unmigrated operation, which stays hand-wired). */
export function registeredOperation(kind: string): AnyOperationDefinition | undefined {
  return BY_KIND.get(kind);
}

export function isRegisteredOperationKind(kind: string): kind is RegisteredOperationKind {
  return BY_KIND.has(kind);
}

/**
 * Build a registered operation's wire body from its address id and arguments.
 * This is the single producer used by the core editor, the React host, and
 * the agent compiler: the address field name and the carried-id budget come
 * from the declaration rather than from a literal at each call site.
 */
export function operationBody<Kind extends RegisteredOperationKind>(
  kind: Kind,
  addressId: StableId,
  args: RegisteredOperationArgs<Kind>,
  allocIds: (n: number) => StableId[] = () => [],
): RegisteredOperationBodyFor<Kind> {
  const definition = BY_KIND.get(kind);
  if (!definition) throw new Error(`${kind} is not a registered operation`);
  if (definition.address === "document") {
    throw new Error(`${kind} is document-scoped; use documentOperationBody`);
  }
  const body: Record<string, unknown> = {
    kind,
    [ADDRESS_WIRE_FIELD[definition.address]]: addressId,
    ...args,
  };
  if (definition.nodeIds) body.nodeIds = allocIds(definition.nodeIds(args as never));
  return body as RegisteredOperationBodyFor<Kind>;
}

/** Build a DOCUMENT-SCOPED operation's wire body. It names no node, so there
 * is no address field and no id to allocate — the payload is the whole of it. */
export function documentOperationBody<Kind extends RegisteredOperationKind>(
  kind: Kind,
  args: RegisteredOperationArgs<Kind>,
): RegisteredOperationBodyFor<Kind> {
  const definition = BY_KIND.get(kind);
  if (!definition) throw new Error(`${kind} is not a registered operation`);
  if (definition.address !== "document") {
    throw new Error(`${kind} is addressed by ${definition.address}; use operationBody`);
  }
  return { kind, ...args } as RegisteredOperationBodyFor<Kind>;
}

/** Validate a registered operation's payload. Null means well-formed. */
export function validateRegisteredOperation(body: RegisteredOperationBody): string | null {
  const definition = BY_KIND.get(body.kind);
  if (!definition) return `${body.kind}: not a registered operation`;
  return definition.validate ? definition.validate(body as never) : null;
}

/** Run a registered operation's mutation against a resolved target. */
export function applyRegisteredOperation(
  doc: DocxDocument,
  target: OperationTarget,
  body: RegisteredOperationBody,
): boolean {
  const definition = BY_KIND.get(body.kind);
  if (!definition) return false;
  return definition.apply({ doc, target, payload: body as never });
}

/** The agent capability row a registration implies. */
export interface RegisteredOperationCapability {
  category: OperationCategory;
  description: string;
  required: string[];
  optional?: string[];
}

export function registeredOperationCapabilities(): Record<
  RegisteredOperationKind,
  RegisteredOperationCapability
> {
  const rows = {} as Record<RegisteredOperationKind, RegisteredOperationCapability>;
  for (const op of OPERATIONS) {
    const required = [
      // A document-scoped operation names no content, so it opens with its own
      // fields rather than an address reference.
      ...(op.address === "document" ? [] : [ADDRESS_AGENT_FIELD[op.address]]),
      ...op.fields.filter((f) => !f.optional).map((f) => f.name),
    ];
    const optional = op.fields.filter((f) => f.optional).map((f) => f.name);
    rows[op.kind] = {
      category: op.category,
      description: op.description,
      required,
      ...(optional.length > 0 ? { optional } : {}),
    };
  }
  return rows;
}
