import { checkboxStateElement, toggleCheckbox } from "../checkbox.js";
import { DocxDocument } from "../docx.js";
import { Run } from "../model.js";
import { XmlElement } from "../xml.js";
import { insertTableAfter } from "./blocks.js";
import { setListType } from "./lists.js";
import { resizeTableRow } from "./tables.js";
import { applyFieldResults } from "./update-fields.js";

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
 * formatting (listKind null). Mutates w:pPr numbering in place. */
const setListTypeOperation = defineOperation<{
  blockId: StableId;
  listKind: "bullet" | "number" | null;
}>()({
  kind: "setListType",
  address: "block",
  category: "paragraph",
  description: "Set or clear paragraph list formatting.",
  fields: [{ name: "listKind" }],
  // setListType resolves the paragraph by walking UP from a target, so pass a
  // descendant w:t when the paragraph has one and the paragraph itself
  // otherwise.
  apply: ({ doc, target, payload }) => setListType(doc, [target.t ?? target.el], payload.listKind),
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

const OPERATIONS = [
  setListTypeOperation,
  insertTableOperation,
  resizeTableRowOperation,
  toggleCheckboxOperation,
  updateFieldsOperation,
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
