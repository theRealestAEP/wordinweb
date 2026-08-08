import { checkboxStateElement, toggleCheckbox } from "../checkbox.js";
import { CITATION_STYLES, citationText, documentBibliography, type CitationStyle } from "../citations.js";
import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement, localName } from "../xml.js";
import { convertTableToText, convertTextToTable, insertTableAfter } from "./blocks.js";
import {
  insertCitationField,
  insertMergeField,
  isValidCitationTag,
  isValidMergeFieldName,
} from "./fields.js";
import { setImageCrop, type ImageCrop } from "./images.js";
import { setListType } from "./lists.js";
import { insertEndnote } from "./notes.js";
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
  STYLE_TYPES,
  type StyleParaPatch,
  type StyleRunPatch,
  type StyleSpec,
  type StylePatch,
} from "./styles.js";
import { insertBibliography, refreshBibliographies } from "./bibliography.js";
import { formulaInstruction, insertTableFormula } from "./formula.js";
import { insertIndex, insertIndexEntry, isValidIndexEntry, refreshIndexes } from "./index-field.js";
import { setModel3DRotation, type Model3DRotation } from "./objects.js";
import {
  badCitationSource,
  badCitationSourcePatch,
  createCitationSource,
  deleteCitationSource,
  editCitationSource,
  setCitationStyle,
  type CitationSourcePatch,
  type CitationSourceSpec,
} from "./sources.js";
import {
  MAX_TAB_STOP_PT,
  PARAGRAPH_BORDER_EDGES,
  TAB_STOP_ALIGNMENTS,
  TAB_STOP_LEADERS,
  setParagraphBorders,
  setTabStops,
  type ParagraphBordersPatch,
  type TabStopSpec,
} from "./paragraph.js";
import {
  PAGE_NUMBER_FORMATS,
  setEvenOddHeaders,
  setPageNumberFormat,
  setTitlePage,
  type PageNumberFormat,
} from "./sections.js";
import {
  HEADER_FOOTER_PRESETS,
  HEADER_FOOTER_PRESET_NODE_BUDGET,
  PAGE_NUMBER_ALIGNMENTS,
  PAGE_NUMBER_POSITIONS,
  insertHeaderFooterPreset,
  insertPageNumberPosition,
  removePageNumberFields,
  type HeaderFooterPreset,
  type PageNumberGalleryAlign,
  type PageNumberGalleryPosition,
} from "./hf-gallery.js";
import { suggestMeta } from "./suggest.js";
import { TOC_LEADERS, insertToc, isValidCaptionLabel, type TocLeader, type TocLevels } from "./toc.js";
import { ensureRefBookmark, insertCaptionAt } from "./references.js";
import { applyFieldResults } from "./update-fields.js";
import { insertWatermark, removeWatermark } from "./watermark.js";
import {
  CELL_SCOPE_EDGES,
  TABLE_BORDER_STYLES,
  TABLE_SCOPE_EDGES,
  resizeTableRow,
  setTableBorders,
  sortTableRows,
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
 *
 * "object" names a DRAWING: the stable id of the run that CARRIES it, plus an
 * `objectIndex` into that run's content picking out which drawing it is (a run
 * can hold several). This is the addressing every hand-written drawing intent
 * already uses — setImageWrap, resizeDrawing — so a migrated one keeps its
 * wire shape. The index is NOT in the table below: it is a position inside the
 * run rather than an id of its own, which is also why it stays out of the
 * agent-facing field list, where a raw index is forbidden. BOTH halves reject:
 * an id nobody has and an index that no longer names a drawing are each the
 * same clean no-op, so the honest-no-op predicate is exactly the one every
 * addressed operation has. The index's WELL-FORMEDNESS is checked centrally
 * (see validateRegisteredOperation) rather than by each operation, for the
 * same reason the three preconditions above are stated once.
 */
export type OperationAddress = "run" | "block" | "cell" | "object" | "document";

/** An operation addressed by a stable id, as opposed to document-scoped. */
export type AddressedOperation = Exclude<OperationAddress, "document">;

/** The wire field carrying the address, per addressed kind. An object address
 * rides on the carrying run's id, narrowed by the payload's objectIndex. */
export const ADDRESS_WIRE_FIELD = {
  run: "runId",
  block: "blockId",
  cell: "cellParagraphId",
  object: "runId",
} as const satisfies Record<AddressedOperation, string>;

/** The agent-facing reference field, per addressed kind. Agents address content
 * with opaque strings ("run:12", "object:12:0"), never raw ids. */
export const ADDRESS_AGENT_FIELD = {
  run: "runRef",
  block: "blockRef",
  cell: "cellRef",
  object: "objectRef",
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
  /** For object addressing: the w:drawing (or VML shape) the address named.
   * Null for every other address. */
  drawing: XmlElement | null;
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
  /** The mutation removes id-tracked nodes, so the caller must retire their
   * stable ids afterwards. Deletion and pruning are paired everywhere else in
   * the apply path (mergeParagraph, removeDrawing); declaring it here keeps
   * the pairing in the operation's single declaration. */
  prunesIds?: true;
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

/**
 * The optional suggest payload a TRACKABLE operation carries, and the
 * validation they share.
 *
 * With it the operation records what it replaced — w:pPrChange for a list
 * marker, w:tblPrChange / w:trPrChange / w:tcPrChange for table formatting —
 * instead of mutating outright, so a reviewer can accept or reject it. The
 * author and date travel in the payload so every replica writes byte-identical
 * XML; the w:id is scan-derived from identical tree state and needs no
 * carrying. Omitting it is a direct, untracked edit.
 */
interface SuggestablePayload {
  suggest?: { author: string; date: string };
}

const SUGGEST_FIELD = { name: "suggest", optional: true } as const;

function badSuggest(kind: string, { suggest }: SuggestablePayload): string | null {
  if (suggest === undefined) return null;
  if (typeof suggest.author !== "string" || suggest.author.length > 100) return `${kind}: bad author`;
  if (typeof suggest.date !== "string" || suggest.date.length > 40) return `${kind}: bad date`;
  return null;
}

/** First w:t under an element, or null. Mirrors the resolution the block
 * address applies to the primary target: setListType resolves paragraphs by
 * walking UP from a target, so each extra paragraph is named by a descendant
 * w:t when it has one and by the paragraph element itself otherwise. */
function firstTextDescendantOf(el: XmlElement): XmlElement | null {
  if (localName(el.name) === "t") return el;
  for (const c of el.children) {
    const found = firstTextDescendantOf(c);
    if (found) return found;
  }
  return null;
}

/** Turn a paragraph into a bullet/numbered list item, or clear its list
 * formatting (listKind null). Mutates w:pPr numbering in place. With `suggest`
 * the change is TRACKED (w:pPrChange) instead of applied outright; the author
 * and date travel in the payload so every replica writes the same XML.
 *
 * `moreBlockIds` extends the operation to a MULTI-PARAGRAPH selection: every
 * listed paragraph is formatted by the SAME mutation call, so exactly one
 * numbering definition is minted and shared — Word's semantics (one list,
 * continuous 1,2,3) and the only convergent shape. One intent per paragraph
 * (the old emission) minted a FRESH definition per apply on the server while
 * the originating client's single local mutation shared one: numId 1,1,1
 * locally vs 1,2,3 canonically — byte divergence plus per-paragraph restarts.
 * An id that no longer resolves is skipped (clean per-target no-op — the id
 * table is identical on every replica, so every replica skips the same ones). */
const setListTypeOperation = defineOperation<{
  blockId: StableId;
  listKind: "bullet" | "number" | null;
  /** Additional target paragraphs beyond the addressed one, sharing its
   * minted numbering definition. Absent = the old single-paragraph form. */
  moreBlockIds?: StableId[];
} & SuggestablePayload>()({
  kind: "setListType",
  address: "block",
  category: "paragraph",
  description: "Set or clear paragraph list formatting.",
  fields: [{ name: "listKind" }, SUGGEST_FIELD],
  validate: (payload) => {
    const more = payload.moreBlockIds;
    if (more !== undefined) {
      if (!Array.isArray(more) || more.length > 10_000 ||
          more.some((id) => typeof id !== "number" || !Number.isInteger(id) || id < 0)) {
        return "setListType: bad moreBlockIds";
      }
    }
    return badSuggest("setListType", payload);
  },
  // setListType resolves the paragraph by walking UP from a target, so pass a
  // descendant w:t when the paragraph has one and the paragraph itself
  // otherwise.
  apply: ({ doc, target, payload }) => {
    const targets = [target.t ?? target.el];
    for (const id of payload.moreBlockIds ?? []) {
      const el = doc.stableIds?.elOf(id);
      if (el) targets.push(firstTextDescendantOf(el) ?? el);
    }
    return setListType(doc, targets, payload.listKind, suggestMeta(doc, payload.suggest));
  },
});

/** Insert a rows×cols table after the paragraph containing the anchor run.
 * `cells` and `headerRow` author the table's content in the same insert. */
const insertTableOperation = defineOperation<{
  runId: StableId;
  rows: number;
  cols: number;
  cells?: string[][];
  headerRow?: boolean;
  nodeIds: StableId[];
}>()({
  kind: "insertTable",
  address: "run",
  category: "insert",
  description: "Insert a table. Pass cells (row-major cell texts) and headerRow to author the whole table — content and bold header row included — in this one operation.",
  fields: [{ name: "rows" }, { name: "cols" }, { name: "cells", optional: true }, { name: "headerRow", optional: true }],
  // A cell is a paragraph inside a run inside a row; the spare 8 cover the
  // table element and the trailing paragraph. Cell text lives inside those
  // runs, so authored content needs no extra ids.
  nodeIds: ({ rows, cols }) =>
    Number.isInteger(rows) && Number.isInteger(cols) && rows > 0 && cols > 0
      ? rows * cols * 2 + 8
      : 8,
  validate: ({ rows, cols, cells, headerRow }) => {
    const okDim = (v: unknown) => typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= 50;
    if (!okDim(rows) || !okDim(cols)) return "insertTable: bad dimensions";
    if (headerRow !== undefined && typeof headerRow !== "boolean") return "insertTable: bad headerRow";
    if (cells !== undefined) {
      if (!Array.isArray(cells) || cells.length > rows) return "insertTable: bad cells";
      for (const row of cells) {
        if (!Array.isArray(row) || row.length > cols) return "insertTable: bad cells";
        for (const cell of row) {
          if (typeof cell !== "string" || cell.length > 2000) return "insertTable: bad cells";
        }
      }
    }
    return null;
  },
  apply: ({ doc, target, payload }) =>
    target.t
      ? insertTableAfter(doc, target.t, payload.rows, payload.cols, {
          cells: payload.cells,
          headerRow: payload.headerRow,
        })
      : false,
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

/**
 * Insert a table of contents after the paragraph holding the anchor run.
 *
 * WHY THIS ONE CARRIES A COUNT. Every other insert's size comes from its own
 * arguments — insertTable is rows×cols — but a TOC's size comes from the
 * DOCUMENT: one entry paragraph per qualifying heading. The carried id
 * allocation is sized by `nodeIds`, which sees the payload and not the
 * document, so the originator asks `tocEntryCount` and sends the answer.
 *
 * `entryCount` is therefore a BUDGET, never an instruction: the mutation still
 * derives its entries from the replica's own headings. Too large wastes ids
 * harmlessly (the assignment stops at the shorter of the two lists); too small
 * would leave the tail of the new paragraphs on locally-generated ids, which
 * is why the budget rounds up. The two cannot disagree in practice because the
 * operation applies at one sequenced position, where every replica's headings
 * are the same.
 *
 * PAGE NUMBERS STAY PLACEHOLDERS in a room. They come from a layout, and a
 * layout depends on the host's font metrics — exactly the value updateFields
 * exists to CARRY rather than recompute. A user who wants them filled runs the
 * update pass, whose results replicate as data.
 */
const insertTocOperation = defineOperation<{
  runId: StableId;
  levels?: TocLevels;
  leader?: TocLeader;
  /** With a label the field is a TABLE OF FIGURES (`TOC \c "Label"`): the
   * entries come from the label's caption paragraphs instead of headings. */
  captionLabel?: string;
  entryCount: number;
  nodeIds: StableId[];
}>()({
  kind: "insertToc",
  address: "run",
  category: "insert",
  description: "Insert a table of contents built from the document's headings, or — with captionLabel — a table of figures built from that label's captions.",
  fields: [{ name: "entryCount" }, { name: "levels", optional: true }, { name: "leader", optional: true }, { name: "captionLabel", optional: true }],
  // Per entry: the w:p plus the seven runs of its hyperlink. The spare covers
  // the three field runs spliced into the first entry and the closing
  // paragraph with its own run.
  nodeIds: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount > 0 ? entryCount * 8 + 8 : 8,
  validate: ({ entryCount, levels, leader, captionLabel }) => {
    if (!Number.isInteger(entryCount) || entryCount < 1 || entryCount > 10000) {
      return "insertToc: bad entryCount";
    }
    if (levels !== undefined) {
      if (!Array.isArray(levels) || levels.length !== 2) return "insertToc: bad levels";
      const [min, max] = levels;
      if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max > 9 || min > max) {
        return "insertToc: bad levels";
      }
    }
    if (leader !== undefined && !TOC_LEADERS.includes(leader)) return "insertToc: bad leader";
    if (captionLabel !== undefined && !isValidCaptionLabel(captionLabel)) return "insertToc: bad captionLabel";
    return null;
  },
  apply: ({ doc, target, payload }) =>
    target.t
      ? insertToc(doc, target.t, {
          ...(payload.levels ? { levels: payload.levels } : {}),
          ...(payload.leader ? { leader: payload.leader } : {}),
          ...(payload.captionLabel ? { captionLabel: payload.captionLabel } : {}),
        })
      : false,
});

/** Insert a Word caption — "<label> <SEQ n>" plus optional text, Caption
 * style — below or above the addressed paragraph's block (the containing
 * table when the paragraph sits in a cell). The SEQ number seeds from
 * document order (deterministic); updateFields renumbers the set. */
const insertCaptionOperation = defineOperation<{
  blockId: StableId;
  label: string;
  text?: string;
  position?: "below" | "above";
  nodeIds: StableId[];
}>()({
  kind: "insertCaption",
  address: "block",
  category: "insert",
  description: 'Insert a caption ("Figure 1", "Table 2", …) below or above the addressed paragraph or its table.',
  fields: [{ name: "label" }, { name: "text", optional: true }, { name: "position", optional: true }],
  // The caption paragraph plus its label run, SEQ result run, and text run.
  nodeIds: () => 8,
  validate: ({ label, text, position }) => {
    if (typeof label !== "string" || !isValidCaptionLabel(label)) return "insertCaption: bad label";
    if (text !== undefined && (typeof text !== "string" || text.length > 2000)) return "insertCaption: bad text";
    if (position !== undefined && position !== "below" && position !== "above") return "insertCaption: bad position";
    return null;
  },
  apply: ({ doc, target, payload }) =>
    insertCaptionAt(doc, target.el, payload.label, payload.text ?? "", payload.position ?? "below"),
});

/** Wrap the addressed paragraph in a hidden `_Ref` bookmark so REF/PAGEREF
 * fields can point at it — the plumbing half of a cross-reference to a
 * heading, caption, or numbered item. The name is the originator's (carried,
 * sequential `_RefN`); a paragraph that already has one is a clean no-op. */
const ensureRefBookmarkOperation = defineOperation<{
  blockId: StableId;
  name: string;
}>()({
  kind: "ensureRefBookmark",
  address: "block",
  category: "insert",
  description: "Wrap the addressed paragraph in a hidden _Ref bookmark for cross-referencing.",
  fields: [{ name: "name" }],
  validate: ({ name }) =>
    typeof name === "string" && /^_Ref\d{1,12}$/.test(name) ? null : "ensureRefBookmark: bad name",
  apply: ({ doc, target, payload }) => ensureRefBookmark(doc, target.el, payload.name),
});

/**
 * Mark an index entry: an invisible XE field (§17.16.5.31) at the end of the
 * addressed run, entry text (with optional `Main:Sub` colon) in the payload.
 */
const insertIndexEntryOperation = defineOperation<{
  runId: StableId;
  entry: string;
  nodeIds: StableId[];
}>()({
  kind: "insertIndexEntry",
  address: "run",
  category: "insert",
  description: 'Mark an index entry: an invisible XE field for the INDEX build. Use a colon for a subentry ("Widgets:assembly").',
  fields: [{ name: "entry" }],
  // The three field runs, plus the runs a mid-text split creates.
  nodeIds: () => 8,
  validate: ({ entry }) => (isValidIndexEntry(entry) ? null : "insertIndexEntry: bad entry"),
  apply: ({ doc, target, payload }) =>
    target.t ? insertIndexEntry(doc, target.t, target.t.text.length, payload.entry) : false,
});

/**
 * Insert an INDEX built from the document's XE marks — the insertToc budget
 * pattern: `entryCount` (one per entry paragraph, asked of indexEntryCount)
 * sizes the carried ids. Entries derive from sequenced state on every
 * replica (deterministic `_Idx` bookmark names, locale-free sort); page
 * numbers land as PAGEREF placeholders and are filled by updateFields, whose
 * results ride as data.
 */
const insertIndexOperation = defineOperation<{
  runId: StableId;
  entryCount: number;
  nodeIds: StableId[];
}>()({
  kind: "insertIndex",
  address: "run",
  category: "insert",
  description: "Insert an alphabetized index built from the document's XE entry marks.",
  fields: [{ name: "entryCount" }],
  // Per entry paragraph: the w:p, its entry text run, and up to ~3 page
  // references at 7 runs each (", " + the five field runs, rounded up).
  // Excess references fall back to parse-order ids, which every replica
  // derives identically (the convertTextToTable rule).
  nodeIds: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount > 0 ? entryCount * 24 + 8 : 8,
  validate: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount >= 1 && entryCount <= 10000
      ? null
      : "insertIndex: bad entryCount",
  apply: ({ doc, target }) => (target.t ? insertIndex(doc, target.t) : false),
});

/**
 * Rebuild every index from the current XE marks — refreshBibliography's
 * shape: document-scoped, structural, rejection predicate the change itself
 * (the comparison blanks harvested page numbers, so an index whose entry
 * structure is unchanged applies nothing and keeps its numbers).
 */
const refreshIndexOperation = defineOperation<{
  entryCount: number;
  nodeIds: StableId[];
}>()({
  kind: "refreshIndex",
  address: "document",
  category: "document",
  description: "Rebuild every index from the document's XE entry marks.",
  fields: [{ name: "entryCount" }],
  nodeIds: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount > 0 ? entryCount * 24 + 8 : 8,
  prunesIds: true,
  validate: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount >= 1 && entryCount <= 10000
      ? null
      : "refreshIndex: bad entryCount",
  apply: ({ doc }) => refreshIndexes(doc),
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

/** One w:tblBorders edge, wherever a border spec is accepted. */
function badBorderSpec(border: unknown, what: string): string | null {
  const spec = border as TableBorderSpec | undefined;
  if (!spec || typeof spec !== "object" || !TABLE_BORDER_STYLES.includes(spec.style)) {
    return `${what}: bad style`;
  }
  if (spec.sz !== undefined && (!Number.isFinite(spec.sz) || spec.sz < 1 || spec.sz > 96)) {
    return `${what}: bad sz`;
  }
  if (spec.space !== undefined && (!Number.isFinite(spec.space) || spec.space < 0 || spec.space > 31)) {
    return `${what}: bad space`;
  }
  if (spec.color !== undefined && !/^(#?[0-9A-Fa-f]{6}|auto)$/.test(spec.color)) {
    return `${what}: bad color`;
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

/** Add a style definition of any of the four kinds to styles.xml. */
const createStyleOperation = defineOperation<{ style: StyleSpec }>()({
  kind: "createStyle",
  address: "document",
  category: "document",
  description: "Create a paragraph, character, table, or numbering style definition.",
  fields: [{ name: "style" }],
  validate: ({ style }) => {
    if (!style || typeof style !== "object" || Array.isArray(style)) return "createStyle: bad style";
    if (!STYLE_TYPES.includes(style.type)) return "createStyle: bad type";
    if (style.type === "character" && style.paragraph) {
      return "createStyle: a character style has no paragraph properties";
    }
    if (style.quickStyle !== undefined && typeof style.quickStyle !== "boolean") {
      return "createStyle: bad quickStyle";
    }
    // A numbering style is a NAME for a list definition and carries nothing
    // else, so its reference is required and its formatting is refused —
    // both, rather than silently dropping what the caller sent.
    if (style.type === "numbering") {
      if (style.paragraph || style.run) return "createStyle: a numbering style carries no formatting";
      const numId = style.numbering?.numId;
      if (!style.numbering || typeof numId !== "number" || !Number.isInteger(numId) || numId < 1 || numId > 32767) {
        return "createStyle: a numbering style needs a numbering definition";
      }
    } else if (style.numbering) {
      return "createStyle: only a numbering style names a numbering definition";
    }
    if (style.table) {
      if (style.type !== "table") return "createStyle: only a table style has table properties";
      const borders = style.table.borders;
      if (borders !== undefined) {
        if (!borders || typeof borders !== "object" || Array.isArray(borders)) return "createStyle: bad borders";
        for (const [edge, spec] of Object.entries(borders)) {
          if (!TABLE_SCOPE_EDGES.includes(edge as TableBorderEdge)) return `createStyle: bad border edge ${edge}`;
          const bad = badBorderSpec(spec, "createStyle");
          if (bad) return bad;
        }
      }
    }
    if (style.linked !== undefined) {
      if (typeof style.linked !== "boolean") return "createStyle: bad linked";
      // The companion is a CHARACTER style carrying a paragraph style's run
      // properties. Nothing else has a companion to link to.
      if (style.linked && style.type !== "paragraph") {
        return "createStyle: only a paragraph style has a linked companion";
      }
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
const HEX_FILL = /^#?[0-9A-Fa-f]{6}$/;

/** Patch a paragraph's DIRECT borders (per edge, w:pBdr §17.3.1.24) and/or
 * shading fill (w:shd §17.3.1.31). Position-stable (pPr only) — transform
 * identity; with `suggest` the paragraph records a w:pPrChange. */
const setParagraphBordersOperation = defineOperation<{
  blockId: StableId;
  patch: ParagraphBordersPatch;
} & SuggestablePayload>()({
  kind: "setParagraphBorders",
  address: "block",
  category: "paragraph",
  description: "Set or clear the paragraph's border edges and shading fill.",
  fields: [{ name: "patch" }, SUGGEST_FIELD],
  validate: (payload) => {
    const patch = payload.patch;
    if (!patch || typeof patch !== "object") return "setParagraphBorders: bad patch";
    const edges = Object.entries(patch.borders ?? {});
    if (edges.length === 0 && patch.shading === undefined) return "setParagraphBorders: empty patch";
    for (const [edge, spec] of edges) {
      if (!(PARAGRAPH_BORDER_EDGES as readonly string[]).includes(edge)) return "setParagraphBorders: bad edge";
      if (spec === null) continue;
      if (!spec || typeof spec !== "object") return "setParagraphBorders: bad border";
      if (!(TABLE_BORDER_STYLES as readonly string[]).includes(spec.style)) return "setParagraphBorders: bad style";
      if (spec.sz !== undefined && (typeof spec.sz !== "number" || spec.sz < 1 || spec.sz > 96)) return "setParagraphBorders: bad sz";
      if (spec.space !== undefined && (typeof spec.space !== "number" || spec.space < 0 || spec.space > 31)) return "setParagraphBorders: bad space";
      if (spec.color !== undefined && !(spec.color === "auto" || HEX_FILL.test(spec.color))) return "setParagraphBorders: bad color";
    }
    if (patch.shading !== undefined && patch.shading !== null && !HEX_FILL.test(patch.shading)) {
      return "setParagraphBorders: bad shading";
    }
    return badSuggest("setParagraphBorders", payload);
  },
  apply: ({ doc, target, payload }) =>
    setParagraphBorders(doc, [target.t ?? target.el], payload.patch, suggestMeta(doc, payload.suggest)),
});

/** Replace a paragraph's DIRECT tab stops (w:tabs, §17.3.1.38). An empty
 * list removes the element so style stops and the default grid apply again.
 * Position-stable (pPr only), so the transform is identity; with `suggest`
 * the paragraph records its prior properties in a w:pPrChange. */
const setTabStopsOperation = defineOperation<{
  blockId: StableId;
  stops: TabStopSpec[];
} & SuggestablePayload>()({
  kind: "setTabStops",
  address: "block",
  category: "paragraph",
  description: "Replace the paragraph's tab stops (position, alignment, leader). An empty list removes them.",
  fields: [{ name: "stops" }, SUGGEST_FIELD],
  validate: (payload) => {
    const stops = payload.stops;
    if (!Array.isArray(stops) || stops.length > 64) return "setTabStops: bad stops";
    for (const stop of stops) {
      if (!stop || typeof stop !== "object") return "setTabStops: bad stop";
      if (typeof stop.posPt !== "number" || !Number.isFinite(stop.posPt) ||
          stop.posPt < -MAX_TAB_STOP_PT || stop.posPt > MAX_TAB_STOP_PT) {
        return "setTabStops: bad pos";
      }
      if (!(TAB_STOP_ALIGNMENTS as readonly string[]).includes(stop.align)) return "setTabStops: bad align";
      if (!(TAB_STOP_LEADERS as readonly string[]).includes(stop.leader)) return "setTabStops: bad leader";
    }
    return badSuggest("setTabStops", payload);
  },
  apply: ({ doc, target, payload }) =>
    setTabStops(doc, [target.t ?? target.el], payload.stops, suggestMeta(doc, payload.suggest)),
});

const setTableBordersOperation = defineOperation<{
  cellParagraphId: StableId;
  scope: "cell" | "table";
  edges: TableBorderEdge[];
  border: TableBorderSpec | null;
} & SuggestablePayload>()({
  kind: "setTableBorders",
  address: "cell",
  category: "table",
  description: "Set or clear table or cell borders, per edge.",
  fields: [{ name: "scope" }, { name: "edges" }, { name: "border" }, SUGGEST_FIELD],
  validate: (payload) => {
    const { scope, edges, border } = payload;
    const bad = badSuggest("setTableBorders", payload);
    if (bad) return bad;
    if (scope !== "cell" && scope !== "table") return "setTableBorders: bad scope";
    if (!Array.isArray(edges) || edges.length === 0 || edges.length > 8) {
      return "setTableBorders: bad edges";
    }
    const allowed = scope === "cell" ? CELL_SCOPE_EDGES : TABLE_SCOPE_EDGES;
    if (edges.some((e) => !allowed.includes(e))) return `setTableBorders: edge not valid at ${scope} scope`;
    if (border === null) return null;
    return badBorderSpec(border, "setTableBorders");
  },
  apply: ({ doc, target, payload }) => {
    const anchor = cellAnchor(target);
    return anchor
      ? setTableBorders(doc, anchor, payload.scope, payload.edges, payload.border, suggestMeta(doc, payload.suggest))
      : false;
  },
});

/** Apply a named table style, or with null remove the reference. */
const setTableStyleOperation = defineOperation<{
  cellParagraphId: StableId;
  styleId: string | null;
} & SuggestablePayload>()({
  kind: "setTableStyle",
  address: "cell",
  category: "table",
  description: "Apply a named table style to a table.",
  fields: [{ name: "styleId" }, SUGGEST_FIELD],
  validate: (payload) => {
    const bad = badSuggest("setTableStyle", payload);
    if (bad) return bad;
    const { styleId } = payload;
    if (styleId === null) return null;
    return typeof styleId === "string" && styleId.length > 0 && styleId.length <= 253
      ? null
      : "setTableStyle: bad styleId";
  },
  apply: ({ doc, target, payload }) =>
    setTableStyle(doc, target.el, payload.styleId, suggestMeta(doc, payload.suggest)),
});

/** Set some of Word's six table-style option toggles (w:tblLook). */
const setTableLookOperation = defineOperation<{
  cellParagraphId: StableId;
  look: Partial<TableLookToggles>;
} & SuggestablePayload>()({
  kind: "setTableLook",
  address: "cell",
  category: "table",
  description: "Toggle which table style options apply (first/last row and column, banding).",
  fields: [{ name: "look" }, SUGGEST_FIELD],
  validate: (payload) => {
    const bad = badSuggest("setTableLook", payload);
    if (bad) return bad;
    const { look } = payload;
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
  apply: ({ doc, target, payload }) =>
    setTableLook(doc, target.el, payload.look, suggestMeta(doc, payload.suggest)),
});

/** Set the table's preferred width (points, percent, or auto). */
const setTableWidthOperation = defineOperation<{
  cellParagraphId: StableId;
  unit: "pt" | "pct" | "auto";
  value?: number;
} & SuggestablePayload>()({
  kind: "setTableWidth",
  address: "cell",
  category: "table",
  description: "Set a table's preferred width in points, as a percent, or to auto.",
  fields: [{ name: "unit" }, { name: "value", optional: true }, SUGGEST_FIELD],
  validate: (payload) => {
    const bad = badSuggest("setTableWidth", payload);
    if (bad) return bad;
    const { unit, value } = payload;
    if (unit !== "pt" && unit !== "pct" && unit !== "auto") return "setTableWidth: bad unit";
    if (unit === "auto") return null;
    if (typeof value !== "number" || !Number.isFinite(value)) return "setTableWidth: value required";
    const max = unit === "pct" ? 100 : 22 * 72;
    return value > 0 && value <= max ? null : "setTableWidth: value out of range";
  },
  apply: ({ doc, target, payload }) =>
    setTableWidth(doc, target.el, payload.unit, payload.value ?? 0, suggestMeta(doc, payload.suggest)),
});

/** Set one grid column to an exact width in points. */
const setTableColumnWidthOperation = defineOperation<{
  cellParagraphId: StableId;
  colIdx: number;
  widthPt: number;
} & SuggestablePayload>()({
  kind: "setTableColumnWidth",
  address: "cell",
  category: "table",
  description: "Set one table column to an exact width in points.",
  fields: [{ name: "colIdx" }, { name: "widthPt" }, SUGGEST_FIELD],
  validate: (payload) => {
    const bad = badSuggest("setTableColumnWidth", payload);
    if (bad) return bad;
    const { colIdx, widthPt } = payload;
    if (!Number.isInteger(colIdx) || colIdx < 0 || colIdx > 200) return "setTableColumnWidth: bad column";
    if (typeof widthPt !== "number" || !Number.isFinite(widthPt) || widthPt < 1 || widthPt > 22 * 72) {
      return "setTableColumnWidth: bad width";
    }
    return null;
  },
  apply: ({ doc, target, payload }) =>
    setTableColumnWidth(doc, target.el, payload.colIdx, payload.widthPt, suggestMeta(doc, payload.suggest)),
});

/** Switch a table between fixed column widths and autofit-to-contents. */
const setTableLayoutOperation = defineOperation<{
  cellParagraphId: StableId;
  layout: "fixed" | "autofit";
  renderedWidths?: number[];
} & SuggestablePayload>()({
  kind: "setTableLayout",
  address: "cell",
  category: "table",
  description: "Switch a table between fixed column widths and autofit.",
  // renderedWidths is the caller's MEASURED columns; it rides as data so
  // every replica freezes the same widths when switching to fixed.
  fields: [{ name: "layout" }, { name: "renderedWidths", optional: true }, SUGGEST_FIELD],
  validate: (payload) => {
    const bad = badSuggest("setTableLayout", payload);
    if (bad) return bad;
    const { layout, renderedWidths } = payload;
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
    setTableLayoutMode(doc, target.el, payload.layout, payload.renderedWidths, suggestMeta(doc, payload.suggest)),
});

/** Set the table's default cell margins, or one cell's override. */
const setTableCellMarginsOperation = defineOperation<{
  cellParagraphId: StableId;
  scope: "cell" | "table";
  margins: CellMarginsPt | null;
} & SuggestablePayload>()({
  kind: "setTableCellMargins",
  address: "cell",
  category: "table",
  description: "Set table default cell margins, or one cell's override, in points.",
  fields: [{ name: "scope" }, { name: "margins" }, SUGGEST_FIELD],
  validate: (payload) => {
    const bad = badSuggest("setTableCellMargins", payload);
    if (bad) return bad;
    const { scope, margins } = payload;
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
    return anchor
      ? setTableCellMargins(doc, anchor, payload.scope, payload.margins, suggestMeta(doc, payload.suggest))
      : false;
  },
});

/**
 * Insert a table formula field (`=SUM(ABOVE)` …, §17.16.5.22) at the end of
 * the addressed cell paragraph. The cached result is EVALUATED on each
 * replica from the containing table's cell texts — a pure function of
 * sequenced state, evaluated locale-free (edit/formula.ts), so nothing
 * nondeterministic needs carrying; updateFields recomputes it thereafter.
 */
const insertTableFormulaOperation = defineOperation<{
  cellParagraphId: StableId;
  formula: string;
  numFmt?: string;
  nodeIds: StableId[];
}>()({
  kind: "insertTableFormula",
  address: "cell",
  category: "table",
  description: 'Insert a table formula field ("=SUM(ABOVE)", "=A1+B2", …) in the addressed cell, with an optional \\# number format like "#,##0.00".',
  fields: [{ name: "formula" }, { name: "numFmt", optional: true }],
  // The fldSimple's result run; the spare covers nothing else today.
  nodeIds: () => 4,
  validate: ({ formula, numFmt }) =>
    formulaInstruction(formula, numFmt) !== null ? null : "insertTableFormula: bad formula",
  apply: ({ doc, target, payload }) => {
    const anchor = cellAnchor(target);
    return anchor ? insertTableFormula(doc, anchor, payload.formula, payload.numFmt) : false;
  },
});

/** Mark the first N rows as the repeating header band. */
const setTableHeaderRowsOperation = defineOperation<{
  cellParagraphId: StableId;
  count: number;
} & SuggestablePayload>()({
  kind: "setTableHeaderRows",
  address: "cell",
  category: "table",
  description: "Repeat the first N rows as a header band on every page.",
  fields: [{ name: "count" }, SUGGEST_FIELD],
  validate: (payload) =>
    badSuggest("setTableHeaderRows", payload) ??
    (Number.isInteger(payload.count) && payload.count >= 0 && payload.count <= 5000
      ? null
      : "setTableHeaderRows: bad count"),
  apply: ({ doc, target, payload }) =>
    setTableHeaderRows(doc, target.el, payload.count, suggestMeta(doc, payload.suggest)),
});

// ---------------------------------------------------------------------------
// Notes and drawings
// ---------------------------------------------------------------------------

/**
 * Insert an endnote at the end of the addressed run, with the given body text.
 *
 * The reference run is APPENDED after the addressed run, so no existing run's
 * text moves and the transform stays identity — the registry's position-stable
 * precondition. (insertNote can also split a run mid-text; that path is the
 * local editor's, where the caret offset is known and no wire transform is
 * involved.)
 *
 * The id comes from max+1 over endnotes.xml, which is sequenced state, so
 * every replica allocates the same one and no id has to travel. The MARK a
 * reader sees — its number format and start value, which a section may set
 * with w:endnotePr — is derived at layout from document order, so insertion
 * writes no numbering of its own.
 */
const insertEndnoteOperation = defineOperation<{
  runId: StableId;
  text: string;
  nodeIds: StableId[];
}>()({
  kind: "insertEndnote",
  address: "run",
  category: "insert",
  description: "Insert an endnote.",
  fields: [{ name: "text" }],
  // The reference run and its rPr, plus the note body's paragraph, mark run
  // and text run; the spare covers a run split when the caret is mid-text.
  nodeIds: () => 8,
  validate: ({ text }) => {
    if (typeof text !== "string" || text.trim().length === 0) return "insertEndnote: empty";
    if (text.length > 20_000) return "insertEndnote: too long";
    return null;
  },
  apply: ({ doc, target, payload }) =>
    target.t ? insertEndnote(doc, target.t, target.t.text.length, payload.text) !== null : false,
});

/**
 * Crop an image: write the a:srcRect fractions trimmed off each edge.
 *
 * OBJECT-addressed, like setModel3DRotation below. It qualifies on the
 * registry's three preconditions like any other: the carrying run's stable id
 * does the rejecting, no run's text moves, and there is no wire inverse.
 */
const setCropOperation = defineOperation<{
  runId: StableId;
  objectIndex?: number;
  crop: ImageCrop;
}>()({
  kind: "setCrop",
  address: "object",
  category: "drawing",
  description: "Crop an image to a fraction of the source bitmap on each edge.",
  fields: [{ name: "crop" }],
  validate: ({ crop }) => {
    if (!crop || typeof crop !== "object" || Array.isArray(crop)) return "setCrop: bad crop";
    const edges = ["l", "t", "r", "b"] as const;
    for (const key of Object.keys(crop)) {
      if (!edges.includes(key as (typeof edges)[number])) return `setCrop: unknown edge ${key}`;
    }
    for (const edge of edges) {
      const value = crop[edge];
      if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || value >= 1) {
        return `setCrop: bad ${edge}`;
      }
    }
    // A crop that trims everything off an axis leaves no picture to draw.
    if (crop.l + crop.r >= 1 || crop.t + crop.b >= 1) return "setCrop: nothing left to show";
    return null;
  },
  apply: ({ doc, target, payload }) =>
    target.drawing ? setImageCrop(doc, target.drawing, payload.crop) : false,
});

/**
 * Save a 3D model's orientation. OBJECT-addressed, like setCrop above.
 *
 * A drag used to mutate the model locally with no wire form and no gate, so it
 * forked any room it happened in: the `model3D` toolbar flag only hides
 * INSERTION, and a model already present in an opened file stays draggable.
 * The rotation is a property of one drawing, which is exactly what the object
 * address names, so nothing here needs the document escape hatch.
 *
 * The angles are CARRIED rather than recomputed. Unlike updateFields that is
 * not because a replica would derive them differently — it is because they
 * come from a pointer drag only the originator saw.
 */
const setModel3DRotationOperation = defineOperation<{
  runId: StableId;
  objectIndex?: number;
  rotation: Model3DRotation;
}>()({
  kind: "setModel3DRotation",
  address: "object",
  category: "drawing",
  description: "Set a 3D model's X/Y/Z orientation, in degrees.",
  fields: [{ name: "rotation" }],
  validate: ({ rotation }) => {
    if (!rotation || typeof rotation !== "object") return "setModel3DRotation: bad rotation";
    for (const axis of ["x", "y", "z"] as const) {
      // Any finite angle is legal; the mutation normalizes into 0..360.
      if (typeof rotation[axis] !== "number" || !Number.isFinite(rotation[axis])) {
        return `setModel3DRotation: bad ${axis}`;
      }
    }
    return null;
  },
  // A drawing that is not a 3D model has no am3d:model3d to write into, which
  // setModel3DRotation reports as a clean no-op.
  apply: ({ doc, target, payload }) =>
    target.drawing ? setModel3DRotation(doc, target.drawing, payload.rotation) : false,
});

/**
 * Insert a MERGEFIELD at the end of the anchor run. With no mail-merge data
 * source attached — this editor never attaches one — the field displays the
 * «Name» placeholder, which is what Word inserts too; the placeholder is
 * written as the cached result so the file shows it everywhere.
 */
const insertMergeFieldOperation = defineOperation<{
  runId: StableId;
  name: string;
  nodeIds: StableId[];
}>()({
  kind: "insertMergeField",
  address: "run",
  category: "insert",
  description: "Insert a mail-merge placeholder field («Name»).",
  fields: [{ name: "name" }],
  // The fldSimple's result run, plus the runs a mid-text split creates; the
  // budget matches the hand-written insertField intent's.
  nodeIds: () => 8,
  validate: ({ name }) => (isValidMergeFieldName(name) ? null : "insertMergeField: bad name"),
  apply: ({ doc, target, payload }) =>
    target.t ? insertMergeField(doc, target.t, target.t.text.length, payload.name) : false,
});

/**
 * Insert a CITATION to a source the document's bibliography already holds.
 *
 * The tag must name a source in the package's sources part — a tag that does
 * not is the operation's honest no-op, and every replica agrees because the
 * sources part is package state, identical at the intent's sequenced
 * position. The cached display text is DERIVED on each replica from that same
 * part (src/citations.ts), so nothing nondeterministic needs carrying.
 */
const insertCitationOperation = defineOperation<{
  runId: StableId;
  tag: string;
  nodeIds: StableId[];
}>()({
  kind: "insertCitation",
  address: "run",
  category: "insert",
  description: "Insert a citation to a bibliography source the document already has, by source tag.",
  fields: [{ name: "tag" }],
  nodeIds: () => 8,
  validate: ({ tag }) => (isValidCitationTag(tag) ? null : "insertCitation: bad tag"),
  apply: ({ doc, target, payload }) => {
    if (!target.t) return false;
    const bibliography = documentBibliography(doc);
    if (!bibliography || !bibliography.sources.has(payload.tag)) return false;
    const display = citationText(`CITATION ${payload.tag} \\l 1033`, bibliography) ?? "";
    return insertCitationField(doc, target.t, target.t.text.length, payload.tag, display);
  },
});

/**
 * Insert a bibliography after the paragraph holding the anchor run: a
 * BIBLIOGRAPHY field whose entries are GENERATED from the package's sources
 * part (edit/bibliography.ts).
 *
 * `entryCount` is the insertToc budget pattern — the field's size comes from
 * the DOCUMENT (one paragraph per source), so the originator asks
 * bibliographyEntryCount and carries the answer to size the id allocation.
 * Unlike a TOC there is no placeholder phase: the entries are a pure function
 * of the sources part, which is sequenced state, so every replica derives
 * identical bytes with nothing else carried.
 */
const insertBibliographyOperation = defineOperation<{
  runId: StableId;
  entryCount: number;
  nodeIds: StableId[];
}>()({
  kind: "insertBibliography",
  address: "run",
  category: "insert",
  description: "Insert a bibliography built from the document's citation sources.",
  fields: [{ name: "entryCount" }],
  // Per entry: the w:p and its text run. The spare covers the three field
  // runs spliced into the first entry and the closing paragraph and run.
  nodeIds: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount > 0 ? entryCount * 2 + 8 : 8,
  validate: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount >= 1 && entryCount <= 10000
      ? null
      : "insertBibliography: bad entryCount",
  apply: ({ doc, target }) => (target.t ? insertBibliography(doc, target.t) : false),
});

/**
 * Regenerate every bibliography's entries from the current sources part —
 * the structural half of keeping one current, the rebuildToc split: entry
 * paragraphs are replaced outright, which the string-carrying updateFields
 * pass can never do. DOCUMENT-scoped with the refreshed content derived per
 * replica (deterministic, unlike a TOC rebuild whose page numbers come from
 * a layout — which is why TOC rebuilds stay local-only while this rides the
 * wire). Its rejection predicate is the change itself: a replica whose
 * bibliographies already match the sources part applies nothing, and every
 * replica compares identical sequenced state. `entryCount` is the carried-id
 * budget: total fresh entry paragraphs across every bibliography field.
 */
const refreshBibliographyOperation = defineOperation<{
  entryCount: number;
  nodeIds: StableId[];
}>()({
  kind: "refreshBibliography",
  address: "document",
  category: "document",
  description: "Regenerate every bibliography from the document's citation sources.",
  fields: [{ name: "entryCount" }],
  nodeIds: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount > 0 ? entryCount * 2 + 8 : 8,
  prunesIds: true,
  validate: ({ entryCount }) =>
    Number.isInteger(entryCount) && entryCount >= 1 && entryCount <= 10000
      ? null
      : "refreshBibliography: bad entryCount",
  apply: ({ doc }) => refreshBibliographies(doc),
});

// ---------------------------------------------------------------------------
// Bibliography sources
// ---------------------------------------------------------------------------

/**
 * The four operations below edit the bibliography SOURCES part (the customXml
 * b:Sources data of ECMA-376 §22.6) rather than content, so they are
 * DOCUMENT-scoped like the style-definition operations, and for the same
 * reasons: the part is sequenced state, every property comes out of the
 * payload, and no run's text moves. The rejection predicate is the source
 * tag — createCitationSource rejects a tag the part already declares, edit
 * and delete reject one it does not — plus deleteCitationSource's honesty
 * clause: a source a CITATION field still references is refused, not
 * orphaned. Creating the part itself (first source in a fresh document) is
 * deterministic; see DocxDocument.sourcesTree.
 */

const createCitationSourceOperation = defineOperation<{ source: CitationSourceSpec }>()({
  kind: "createCitationSource",
  address: "document",
  category: "document",
  description: "Add a bibliography source (book, article, website, or report) to the document's citation sources, creating the sources part when the document has none.",
  fields: [{ name: "source" }],
  validate: ({ source }) => badCitationSource(source, "createCitationSource"),
  apply: ({ doc, payload }) => createCitationSource(doc, payload.source),
});

const editCitationSourceOperation = defineOperation<{
  tag: string;
  patch: CitationSourcePatch;
}>()({
  kind: "editCitationSource",
  address: "document",
  category: "document",
  description: "Change a bibliography source's fields, by source tag. Run updateFields afterwards to refresh citation text.",
  fields: [{ name: "tag" }, { name: "patch" }],
  validate: ({ tag, patch }) =>
    isValidCitationTag(tag)
      ? badCitationSourcePatch(patch, "editCitationSource")
      : "editCitationSource: bad tag",
  apply: ({ doc, payload }) => editCitationSource(doc, payload.tag, payload.patch),
});

const deleteCitationSourceOperation = defineOperation<{ tag: string }>()({
  kind: "deleteCitationSource",
  address: "document",
  category: "document",
  description: "Delete a bibliography source, by tag. Refused while any CITATION field still references it.",
  fields: [{ name: "tag" }],
  validate: ({ tag }) => (isValidCitationTag(tag) ? null : "deleteCitationSource: bad tag"),
  apply: ({ doc, payload }) => deleteCitationSource(doc, payload.tag),
});

const setCitationStyleOperation = defineOperation<{ style: CitationStyle }>()({
  kind: "setCitationStyle",
  address: "document",
  category: "document",
  description: "Select the citation style (APA or MLA) for citations and the bibliography. Run updateFields afterwards to refresh citation text.",
  fields: [{ name: "style" }],
  validate: ({ style }) =>
    CITATION_STYLES.includes(style) ? null : "setCitationStyle: bad style",
  apply: ({ doc, payload }) => setCitationStyle(doc, payload.style),
});

// ---------------------------------------------------------------------------
// Watermarks
// ---------------------------------------------------------------------------

/**
 * The two watermark operations are DOCUMENT-scoped because a watermark is a
 * property of the document, not of a place in it: Word writes one into every
 * header part so it shows on every page, and there is no single node either
 * operation could be addressed at.
 *
 * `headerCount` is the pattern insertToc established, doing both jobs at once.
 * As a BUDGET it sizes the carried ids, because `nodeIds` sees the payload and
 * not the document, and the number of nodes created depends on how many header
 * parts the document has. As a REJECTION PREDICATE it stands in for the stable
 * id a document-scoped operation does not have: a replica whose header set has
 * moved under a concurrent edit applies nothing, and every replica in that
 * position rejects identically.
 *
 * Both are position-stable in the registry's sense. The watermark run is
 * prepended to a header paragraph that already exists, so no run's text moves
 * and a concurrent text intent still transforms as identity.
 *
 * Neither creates the header part it needs. That is `ensureHeaderFooter`'s
 * job, which is a structural intent of its own; a caller in a room submits it
 * first and this one second.
 */

/** RRGGBB, the colour form every other operation on the wire already uses.
 * VML would also take a keyword ("silver"), but one convention is worth more
 * than the second spelling. */
const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;

function badHeaderCount(count: unknown, what: string): string | null {
  return Number.isInteger(count) && (count as number) >= 1 && (count as number) <= 50
    ? null
    : `${what}: bad headerCount`;
}

const insertWatermarkOperation = defineOperation<{
  text: string;
  headerCount: number;
  diagonal?: boolean;
  color?: string;
  opacity?: number;
  nodeIds: StableId[];
}>()({
  kind: "insertWatermark",
  address: "document",
  category: "insert",
  description: "Stamp a text watermark across every page, in the document's header parts.",
  fields: [
    { name: "text" },
    { name: "headerCount" },
    { name: "diagonal", optional: true },
    { name: "color", optional: true },
    { name: "opacity", optional: true },
  ],
  // Per header part: the watermark run, plus the paragraph to host it when the
  // part has none of its own.
  nodeIds: ({ headerCount }) => (badHeaderCount(headerCount, "x") ? 2 : headerCount * 2),
  // Inserting REPLACES any watermark already there, so it deletes runs too.
  prunesIds: true,
  validate: ({ text, headerCount, diagonal, color, opacity }) => {
    if (typeof text !== "string" || text.length === 0 || text.length > 255) {
      return "insertWatermark: bad text";
    }
    const bad = badHeaderCount(headerCount, "insertWatermark");
    if (bad) return bad;
    if (diagonal !== undefined && typeof diagonal !== "boolean") return "insertWatermark: bad diagonal";
    if (color !== undefined && (typeof color !== "string" || !HEX_COLOR.test(color))) {
      return "insertWatermark: bad color";
    }
    if (opacity !== undefined && (typeof opacity !== "number" || !(opacity >= 0 && opacity <= 1))) {
      return "insertWatermark: bad opacity";
    }
    return null;
  },
  apply: ({ doc, payload }) =>
    doc.headerRoots().length === payload.headerCount &&
    insertWatermark(doc, {
      text: payload.text,
      ...(payload.diagonal !== undefined ? { diagonal: payload.diagonal } : {}),
      ...(payload.color !== undefined ? { color: payload.color } : {}),
      ...(payload.opacity !== undefined ? { opacity: payload.opacity } : {}),
    }),
});

/**
 * Sort a table's body rows by one grid column. STRUCTURAL (rows reorder, no
 * tracked form), so suggestion-mode surfaces refuse it like the other
 * structural table ops; rows keep their element identity, so every stable id
 * survives and the transform stays identity. The mutation's comparisons are
 * locale-free by design — see sortTableRows — so every replica orders the
 * rows identically.
 */
const sortTableRowsOperation = defineOperation<{
  cellParagraphId: StableId;
  colIdx: number;
  order: "asc" | "desc";
  /** "kind" would collide with the intent discriminant, hence "compare". */
  compare: "text" | "number";
  hasHeader?: boolean;
}>()({
  kind: "sortTableRows",
  address: "cell",
  category: "table",
  description: "Sort a table's body rows by one column, as text or as numbers.",
  fields: [
    { name: "colIdx" },
    { name: "order" },
    { name: "compare" },
    { name: "hasHeader", optional: true },
  ],
  validate: (payload) => {
    if (!Number.isInteger(payload.colIdx) || payload.colIdx < 0 || payload.colIdx > 200) {
      return "sortTableRows: bad column";
    }
    if (payload.order !== "asc" && payload.order !== "desc") return "sortTableRows: bad order";
    if (payload.compare !== "text" && payload.compare !== "number") return "sortTableRows: bad compare";
    if (payload.hasHeader !== undefined && typeof payload.hasHeader !== "boolean") {
      return "sortTableRows: bad hasHeader";
    }
    return null;
  },
  apply: ({ doc, target, payload }) =>
    sortTableRows(doc, target.el, payload.colIdx, payload.order, payload.compare, payload.hasHeader ?? false),
});

/**
 * Convert paragraphs into a table (Word's Convert Text to Table): each
 * paragraph becomes a row, split on the separator into cells. STRUCTURAL —
 * the paragraphs are replaced by a fresh table, so the removed ids are pruned
 * and the new table's nodes take carried ids. `moreBlockIds` extends the
 * conversion to the rest of a multi-paragraph selection, exactly one intent
 * for the whole gesture (the setListType shape). `cellCount` is the
 * originator's carried-id BUDGET (rows × columns as it computed them);
 * replicas whose apply mints more nodes fall back to parse-order ids for the
 * excess, which every replica derives identically.
 */
const convertTextToTableOperation = defineOperation<{
  blockId: StableId;
  moreBlockIds?: StableId[];
  separator: "tab" | "comma";
  cellCount: number;
  nodeIds: StableId[];
}>()({
  kind: "convertTextToTable",
  address: "block",
  category: "table",
  description: "Convert paragraphs into a table, one row per paragraph, cells split on a separator.",
  fields: [{ name: "separator" }, { name: "cellCount" }],
  // tbl + one p and r per cell + the trailing empty paragraph's p and r.
  nodeIds: (args) => 3 + 2 * Math.max(1, Math.min(args.cellCount, 10_000)),
  prunesIds: true,
  validate: (payload) => {
    if (payload.separator !== "tab" && payload.separator !== "comma") return "convertTextToTable: bad separator";
    if (!Number.isInteger(payload.cellCount) || payload.cellCount < 1 || payload.cellCount > 10_000) {
      return "convertTextToTable: bad cellCount";
    }
    const more = payload.moreBlockIds;
    if (more !== undefined) {
      if (!Array.isArray(more) || more.length > 10_000 ||
          more.some((id) => typeof id !== "number" || !Number.isInteger(id) || id < 0)) {
        return "convertTextToTable: bad moreBlockIds";
      }
    }
    return null;
  },
  apply: ({ doc, target, payload }) => {
    const targets = [target.el];
    for (const id of payload.moreBlockIds ?? []) {
      const el = doc.stableIds?.elOf(id);
      if (el) targets.push(el);
    }
    return convertTextToTable(doc, targets, payload.separator);
  },
});

/**
 * Convert a table into paragraphs (Word's Convert Table to Text), one per
 * row, cell texts joined by tabs or commas. STRUCTURAL — the table is
 * replaced, so its ids are pruned and the new paragraphs take carried ids.
 * `rowCount` is the originator's carried-id budget, like cellCount above.
 */
const convertTableToTextOperation = defineOperation<{
  cellParagraphId: StableId;
  separator: "tab" | "comma";
  rowCount: number;
  nodeIds: StableId[];
}>()({
  kind: "convertTableToText",
  address: "cell",
  category: "table",
  description: "Convert a table into paragraphs, one per row, cells joined by a separator.",
  fields: [{ name: "separator" }, { name: "rowCount" }],
  // One p and one r per row.
  nodeIds: (args) => 2 * Math.max(1, Math.min(args.rowCount, 10_000)),
  prunesIds: true,
  validate: (payload) => {
    if (payload.separator !== "tab" && payload.separator !== "comma") return "convertTableToText: bad separator";
    if (!Number.isInteger(payload.rowCount) || payload.rowCount < 1 || payload.rowCount > 10_000) {
      return "convertTableToText: bad rowCount";
    }
    return null;
  },
  apply: ({ doc, target, payload }) => convertTableToText(doc, target.el, payload.separator),
});

/**
 * Toggle Word's "Different First Page" (w:titlePg — ECMA-376 §17.10.6) on
 * every section. Enabling creates the missing first-page header and footer
 * parts (empty, referenced w:type="first" — §17.10.5) the way Word does when
 * the box is first checked; the created part's paragraph and run take carried
 * ids so later edits in the band address them identically everywhere.
 * Disabling removes only the toggle — parts and their content stay, so
 * re-enabling restores them. The rejection predicate is the document's own
 * state: a replica already in the requested state applies nothing, and every
 * replica at the sequenced position shares that state.
 */
const setTitlePageOperation = defineOperation<{
  enabled: boolean;
  nodeIds: StableId[];
}>()({
  kind: "setTitlePage",
  address: "document",
  category: "document",
  description: "Toggle different-first-page headers and footers (w:titlePg), creating the empty first-page parts on enable.",
  fields: [{ name: "enabled" }],
  // At most one new header part and one new footer part, each holding one
  // paragraph with one run.
  nodeIds: () => 4,
  validate: ({ enabled }) => (typeof enabled === "boolean" ? null : "setTitlePage: bad enabled"),
  apply: ({ doc, payload }) => setTitlePage(doc, payload.enabled),
});

/**
 * Toggle Word's "Different Odd & Even Pages" — the document-global
 * w:evenAndOddHeaders switch in settings.xml (ECMA-376 §17.10.1). Enabling
 * creates the missing even-page header and footer parts (empty, referenced
 * w:type="even") because an even page then uses ONLY the even variant.
 * Same shape as setTitlePage in every other respect.
 */
const setEvenOddHeadersOperation = defineOperation<{
  enabled: boolean;
  nodeIds: StableId[];
}>()({
  kind: "setEvenOddHeaders",
  address: "document",
  category: "document",
  description: "Toggle different odd & even page headers and footers (w:evenAndOddHeaders), creating the empty even-page parts on enable.",
  fields: [{ name: "enabled" }],
  nodeIds: () => 4,
  validate: ({ enabled }) => (typeof enabled === "boolean" ? null : "setEvenOddHeaders: bad enabled"),
  apply: ({ doc, payload }) => setEvenOddHeaders(doc, payload.enabled),
});

/**
 * Patch the hyphenation settings — w:autoHyphenation, w:hyphenationZone (in
 * points on the wire, the api's unit convention), w:doNotHyphenateCaps.
 * DOCUMENT-scoped like the other settings.xml toggles; rejection predicate is
 * the change itself (doc.setHyphenation returns false when every present key
 * already holds its value, identically on every replica). NOTE: this engine's
 * layout does not hyphenate automatically — the settings are round-trip state
 * that governs Word's rendering of the file.
 */
const setHyphenationOperation = defineOperation<{
  auto?: boolean;
  zonePt?: number | null;
  noCaps?: boolean;
}>()({
  kind: "setHyphenation",
  address: "document",
  category: "document",
  description: "Set automatic-hyphenation settings: on/off, the hyphenation zone in points, and whether words in capitals stay whole.",
  fields: [
    { name: "auto", optional: true },
    { name: "zonePt", optional: true },
    { name: "noCaps", optional: true },
  ],
  validate: ({ auto, zonePt, noCaps }) => {
    if (auto === undefined && zonePt === undefined && noCaps === undefined) return "setHyphenation: empty patch";
    if (auto !== undefined && typeof auto !== "boolean") return "setHyphenation: bad auto";
    if (zonePt !== undefined && zonePt !== null &&
        (typeof zonePt !== "number" || !Number.isFinite(zonePt) || zonePt <= 0 || zonePt > 1584)) {
      return "setHyphenation: bad zonePt";
    }
    if (noCaps !== undefined && typeof noCaps !== "boolean") return "setHyphenation: bad noCaps";
    return null;
  },
  apply: ({ doc, payload }) =>
    doc.setHyphenation({
      ...(payload.auto !== undefined ? { auto: payload.auto } : {}),
      ...(payload.zonePt !== undefined
        ? { zoneTwips: payload.zonePt === null ? null : Math.round(payload.zonePt * 20) }
        : {}),
      ...(payload.noCaps !== undefined ? { noCaps: payload.noCaps } : {}),
    }),
});

/**
 * Set the page-number format and/or start-at value (w:pgNumType —
 * ECMA-376 §17.6.12; formats are the editable subset of ST_NumberFormat
 * §17.18.59 the layout paints). Document-level like setPageLayout: every
 * section updates, which is the demo's section-scope limitation in a room.
 * A null fmt or start removes the attribute; an empty w:pgNumType is removed
 * entirely. Rejects (honest no-op) when nothing would change.
 */
const setPageNumberFormatOperation = defineOperation<{
  fmt?: PageNumberFormat | null;
  start?: number | null;
}>()({
  kind: "setPageNumberFormat",
  address: "document",
  category: "document",
  description: "Set the page-number format (decimal, roman, letter) and/or the start-at value for the document's sections.",
  fields: [{ name: "fmt", optional: true }, { name: "start", optional: true }],
  validate: ({ fmt, start }) => {
    if (fmt === undefined && start === undefined) return "setPageNumberFormat: empty patch";
    if (fmt !== undefined && fmt !== null && !PAGE_NUMBER_FORMATS.includes(fmt)) {
      return "setPageNumberFormat: bad fmt";
    }
    if (start !== undefined && start !== null &&
        (!Number.isInteger(start) || start < 0 || start > 32767)) {
      return "setPageNumberFormat: bad start";
    }
    return null;
  },
  apply: ({ doc, payload }) =>
    setPageNumberFormat(doc, {
      ...(payload.fmt !== undefined ? { fmt: payload.fmt } : {}),
      ...(payload.start !== undefined ? { start: payload.start } : {}),
    }),
});

/**
 * Word's Insert > Page Number > Top/Bottom of Page position gallery: a
 * single live PAGE field in the header ("top") or footer ("bottom"),
 * aligned left/center/right. Composed from ensureHfPart (the same part
 * creation `ensureHeaderFooter` uses) plus the PAGE field vocabulary
 * fields.ts already allows — no new wire vocabulary. Like insertWatermark,
 * a pick REPLACES the part's content, matching Word's own gallery, so it
 * always applies rather than checking for an unchanged-content no-op.
 */
const insertPageNumberPositionOperation = defineOperation<{
  position: PageNumberGalleryPosition;
  align: PageNumberGalleryAlign;
  nodeIds: StableId[];
}>()({
  kind: "insertPageNumberPosition",
  address: "document",
  category: "insert",
  description: "Insert a page number into the header (top) or footer (bottom), aligned left/center/right — the page-number position gallery.",
  fields: [{ name: "position" }, { name: "align" }],
  // The gallery paragraph plus the fldSimple's one result run.
  nodeIds: () => 2,
  prunesIds: true,
  validate: ({ position, align }) => {
    if (!(PAGE_NUMBER_POSITIONS as readonly string[]).includes(position)) return "insertPageNumberPosition: bad position";
    if (!(PAGE_NUMBER_ALIGNMENTS as readonly string[]).includes(align)) return "insertPageNumberPosition: bad align";
    return null;
  },
  apply: ({ doc, payload }) => insertPageNumberPosition(doc, payload.position, payload.align),
});

/** Word's "Remove Page Numbers": strip PAGE/NUMPAGES fields from every
 * header and footer part. Rejection predicate is what it found: a document
 * with no page-number fields applies nothing. */
const removePageNumbersOperation = defineOperation<Record<never, never>>()({
  kind: "removePageNumbers",
  address: "document",
  category: "insert",
  description: "Remove page-number fields from every header and footer part.",
  fields: [],
  prunesIds: true,
  apply: ({ doc }) => removePageNumberFields(doc),
});

/**
 * The Header & Footer preset gallery: replace a header or footer's content
 * with one of four preset layouts (blank, centered title, title + date,
 * three-column). Composed from literal placeholder text, the existing DATE
 * and PAGE field vocabulary, and setTabStops (the wave-2 tab-stop op) for
 * the three-column layout's tab positions — see hf-gallery.ts. Same
 * always-applies shape as insertPageNumberPosition.
 */
const insertHeaderFooterPresetOperation = defineOperation<{
  hfKind: "header" | "footer";
  preset: HeaderFooterPreset;
  nodeIds: StableId[];
}>()({
  kind: "insertHeaderFooterPreset",
  address: "document",
  category: "insert",
  description: "Replace a header or footer's content with a preset layout: blank, centered title, title + date, or three-column.",
  fields: [{ name: "hfKind" }, { name: "preset" }],
  nodeIds: ({ preset }) => HEADER_FOOTER_PRESET_NODE_BUDGET[preset] ?? 2,
  prunesIds: true,
  validate: ({ hfKind, preset }) => {
    if (hfKind !== "header" && hfKind !== "footer") return "insertHeaderFooterPreset: bad hfKind";
    if (!(HEADER_FOOTER_PRESETS as readonly string[]).includes(preset)) return "insertHeaderFooterPreset: bad preset";
    return null;
  },
  apply: ({ doc, payload }) => insertHeaderFooterPreset(doc, payload.hfKind, payload.preset),
});

/** A payload with no fields of its own. Not `Record<string, never>`: that
 * carries an index signature, which would forbid the clientId/clientSeq/base
 * bookkeeping @wordinweb/collab intersects onto every wire body. */
type NoPayload = Record<never, never>;

/** Take the watermark back off every page. Its rejection predicate is the
 * watermark's own existence: a replica with none applies nothing. */
const removeWatermarkOperation = defineOperation<NoPayload>()({
  kind: "removeWatermark",
  address: "document",
  category: "insert",
  description: "Remove the text watermark from every page.",
  fields: [],
  prunesIds: true,
  apply: ({ doc }) => removeWatermark(doc),
});

const OPERATIONS = [
  setListTypeOperation,
  insertTableOperation,
  resizeTableRowOperation,
  toggleCheckboxOperation,
  updateFieldsOperation,
  insertTocOperation,
  insertCaptionOperation,
  ensureRefBookmarkOperation,
  insertIndexEntryOperation,
  insertIndexOperation,
  refreshIndexOperation,
  createStyleOperation,
  modifyStyleOperation,
  deleteStyleOperation,
  setNumberingLevelOperation,
  setNumberingRestartOperation,
  setParagraphBordersOperation,
  setTabStopsOperation,
  setTableBordersOperation,
  setTableStyleOperation,
  setTableLookOperation,
  setTableWidthOperation,
  setTableColumnWidthOperation,
  setTableLayoutOperation,
  setTableCellMarginsOperation,
  setTableHeaderRowsOperation,
  insertTableFormulaOperation,
  sortTableRowsOperation,
  convertTextToTableOperation,
  convertTableToTextOperation,
  insertEndnoteOperation,
  setCropOperation,
  setModel3DRotationOperation,
  insertMergeFieldOperation,
  insertCitationOperation,
  insertBibliographyOperation,
  refreshBibliographyOperation,
  createCitationSourceOperation,
  editCitationSourceOperation,
  deleteCitationSourceOperation,
  setCitationStyleOperation,
  insertWatermarkOperation,
  removeWatermarkOperation,
  setTitlePageOperation,
  setEvenOddHeadersOperation,
  setPageNumberFormatOperation,
  setHyphenationOperation,
  insertPageNumberPositionOperation,
  removePageNumbersOperation,
  insertHeaderFooterPresetOperation,
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
 * is no address field — but it may still CREATE nodes (insertWatermark adds a
 * run per header part), and those need carried ids like any other insert. */
export function documentOperationBody<Kind extends RegisteredOperationKind>(
  kind: Kind,
  args: RegisteredOperationArgs<Kind>,
  allocIds: (n: number) => StableId[] = () => [],
): RegisteredOperationBodyFor<Kind> {
  const definition = BY_KIND.get(kind);
  if (!definition) throw new Error(`${kind} is not a registered operation`);
  if (definition.address !== "document") {
    throw new Error(`${kind} is addressed by ${definition.address}; use operationBody`);
  }
  const body: Record<string, unknown> = { kind, ...args };
  if (definition.nodeIds) body.nodeIds = allocIds(definition.nodeIds(args as never));
  return body as RegisteredOperationBodyFor<Kind>;
}

/** Validate a registered operation's payload. Null means well-formed. */
export function validateRegisteredOperation(body: RegisteredOperationBody): string | null {
  const definition = BY_KIND.get(body.kind);
  if (!definition) return `${body.kind}: not a registered operation`;
  // The object address's second half is checked HERE, not per operation: it
  // belongs to the address rather than to any one payload, so validating it
  // once means a new object-addressed operation cannot forget to. Absent is
  // legal — it means the run's first drawing.
  if (definition.address === "object") {
    const { objectIndex } = body as { objectIndex?: unknown };
    if (objectIndex !== undefined) {
      if (!Number.isInteger(objectIndex) || (objectIndex as number) < 0 || (objectIndex as number) > 1000) {
        return `${body.kind}: bad objectIndex`;
      }
    }
  }
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
