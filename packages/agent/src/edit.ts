import { localName, registeredOperation, runWireLength, type DocxDocument, type StableIds } from "@wordinweb/core";
import { sha256Hex, type Intent, type IntentBody } from "@wordinweb/collab/client";
import { validateIntent } from "@wordinweb/collab/server";
import { AGENT_EDIT_CAPABILITIES, validateAgentOperationShape } from "./capabilities.js";
import { parseObjectRef, resolveParagraphRef, resolveRunRef } from "./refs.js";
import type { AgentAsset, AgentOperation, AgentProvenance } from "./types.js";

interface CompileContext {
  doc: DocxDocument;
  ids: StableIds;
  allocateIds(count: number): number[];
  provenance: Required<AgentProvenance>;
  asset(ref: string): AgentAsset;
  prepareMedia(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number; iv?: string }>;
}

/** Operations whose agent-facing `suggest: true` compiles to carried
 * author/date. Each one has a tracked OOXML form: w:ins for an insertion or a
 * split mark, a `*PrChange` for a formatting change, a struck paragraph mark
 * for a merge. */
const SUGGESTABLE_KINDS = new Set<Intent["kind"]>([
  "insertText",
  "insertSeparator",
  "deleteSeparator",
  "splitParagraph",
  "mergeParagraph",
  "formatRun",
  "formatRange",
  "formatParagraph",
  "setListType",
  "adjustIndent",
  "setSpacing",
  // Table formatting: w:tblPrChange, w:trPrChange, w:tcPrChange.
  "tableOp",
  "setTableBorders",
  "setTableStyle",
  "setTableLook",
  "setTableWidth",
  "setTableColumnWidth",
  "setTableLayout",
  "setTableCellMargins",
  "setTableHeaderRows",
]);

/**
 * The tableOp operations that carry a tracked form: the three that change
 * PROPERTIES. The rest — every row and column insert and delete, the
 * whole-table delete, a merge and a split — are STRUCTURAL. Word records those
 * on the row (w:trPr/w:ins, w:trPr/w:del, with the cell content marked too)
 * and this engine writes none of it, so suggestion mode refuses them rather
 * than restructure a table behind the reviewer's back.
 */
const SUGGESTABLE_TABLE_OPS = new Set(["cellShading", "cellVAlign", "textWrapping"]);

/** Whether a tableOp's `op` has a tracked form, so it may run while
 * suggesting. The structural ops are plain strings and never qualify. */
export function isSuggestableTableOp(op: unknown): boolean {
  if (!op || typeof op !== "object" || Array.isArray(op)) return false;
  return SUGGESTABLE_TABLE_OPS.has(String((op as { kind?: unknown }).kind));
}

const INTERNAL_FIELDS = ["clientId", "clientSeq", "base", "nodeIds", "newBlockId", "newRunId", "beforeId", "middleId", "afterId", "blockId", "runId", "objectIndex", "cellParagraphId", "afterBlockId"];

function cloneOperation(operation: AgentOperation): Record<string, unknown> {
  if (!operation || typeof operation !== "object" || Array.isArray(operation)) throw new Error("Operation must be an object");
  return JSON.parse(JSON.stringify(operation)) as Record<string, unknown>;
}

function requireFields(operation: Record<string, unknown>, required: string[]): void {
  for (const field of required) if (!(field in operation)) throw new Error(`${String(operation.kind)} requires ${field}`);
}

function position(value: unknown, ids: StableIds): { blockId: number; runId: number; offset: number } {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Position must be an object");
  const input = value as Record<string, unknown>;
  const offset = input.offset;
  if (!Number.isInteger(offset) || (offset as number) < 0) throw new Error("Position offset must be a non-negative integer");
  return { blockId: resolveParagraphRef(ids, input.blockRef), runId: resolveRunRef(ids, input.runRef), offset: offset as number };
}

function paraId(context: CompileContext): string {
  return context.provenance.nextId().replace(/[^0-9a-f]/gi, "").slice(-8).padStart(8, "0").toUpperCase();
}

function idsForOperation(kind: Intent["kind"], operation: Record<string, unknown>): number {
  // A registered operation declares its own carried-id budget, so the editor,
  // the React host, and this compiler all size the allocation the same way.
  const registered = registeredOperation(kind);
  if (registered) return registered.nodeIds ? registered.nodeIds(operation as never) : 0;
  switch (kind) {
    case "pasteBlocks": {
      const xml = String(operation.blocksXml ?? "");
      return (xml.match(/<(?:\w+:)?(?:p|r)(?:\s|>)/g)?.length ?? 0) + 8;
    }
    case "tableOp": return 16;
    case "insertImage":
    case "insertBreak":
    case "insertPageField":
    case "insertFootnote":
    case "setDropCap":
    case "insertBlankPage":
    case "insertSectionBreak":
    case "insertCrossRef":
    case "insertDateTimeField":
    case "insertField":
    case "moveMath":
    case "ensureHeaderFooter": return 8;
    case "setLink": return 4;
    case "insertShape":
    case "insertWordArt": return 12;
    case "insertMath":
    case "insertCoverPage":
    case "insertChart":
    case "insertSmartArt": return 24;
    default: return 0;
  }
}

function imageExtension(mediaType: string): string {
  const extensions: Record<string, string> = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/webp": "webp",
  };
  const extension = extensions[mediaType];
  if (!extension) throw new Error("The image asset type is not supported for insertion");
  return extension;
}

/**
 * Expand an insertText whose text contains newlines into the intent sequence
 * that creates real paragraphs: one splitParagraph per "\n" (each with carried
 * ids, all at the original caret, so no split addresses a run an insertion
 * created), then one insertText per non-empty line into a run known ahead of
 * time. Front-to-back splits at one point stack the new paragraphs directly
 * after the addressed one, with the original tail in the paragraph the FIRST
 * split created — so line k lands in the paragraph from split n-k.
 */
function expandInsertText(operation: Record<string, unknown>, context: CompileContext): Array<Record<string, unknown>> {
  const text = String(operation.text).replace(/\r\n?/g, "\n");
  if (/[\v\f]/.test(text)) {
    throw new Error('insertText text may not contain vertical-tab or form-feed characters. Use "\\n" to start a new paragraph.');
  }
  if (!text.includes("\n")) return [{ ...operation, text }];
  const at = operation.at as { blockId: number; runId: number; offset: number };
  const suggest = operation.suggest;
  const lines = text.split("\n");
  const intents: Array<Record<string, unknown>> = [];
  const created: Array<{ blockId: number; runId: number }> = [];
  for (let index = 0; index < lines.length - 1; index++) {
    const [newBlockId, newRunId] = context.allocateIds(2);
    intents.push({ kind: "splitParagraph", at: { ...at }, newBlockId, newRunId, ...(suggest ? { suggest } : {}) });
    created.push({ blockId: newBlockId, runId: newRunId });
  }
  if (lines[0]) intents.push({ kind: "insertText", at: { ...at }, text: lines[0], ...(suggest ? { suggest } : {}) });
  for (let index = 1; index < lines.length; index++) {
    if (!lines[index]) continue;
    const target = created[lines.length - 1 - index];
    intents.push({ kind: "insertText", at: { blockId: target.blockId, runId: target.runId, offset: 0 }, text: lines[index], ...(suggest ? { suggest } : {}) });
  }
  return intents;
}

export async function compileAgentOperation(input: AgentOperation, context: CompileContext): Promise<IntentBody[]> {
  const operation = cloneOperation(input);
  const kind = operation.kind;
  if (typeof kind !== "string" || !(kind in AGENT_EDIT_CAPABILITIES)) throw new Error("Unknown edit operation");
  const typedKind = kind as Intent["kind"];
  for (const field of INTERNAL_FIELDS) if (field in operation) throw new Error(`${field} is internal; use opaque references`);
  const shapeError = validateAgentOperationShape(input);
  if (shapeError) throw new Error(shapeError);
  requireFields(operation, AGENT_EDIT_CAPABILITIES[typedKind].required);

  // The agent schema admits "#RRGGBB" for cell shading; the wire wants bare
  // hex. Normalize here so the one agent-only spelling cannot be rejected
  // after it passed the schema.
  if (typedKind === "tableOp") {
    const op = operation.op as { kind?: unknown; fill?: unknown } | string | undefined;
    if (op && typeof op === "object" && op.kind === "cellShading" && typeof op.fill === "string") {
      op.fill = op.fill.replace(/^#/, "");
    }
  }

  if ("blockRef" in operation) {
    operation.blockId = resolveParagraphRef(context.ids, operation.blockRef);
    delete operation.blockRef;
  }
  if ("runRef" in operation) {
    operation.runId = resolveRunRef(context.ids, operation.runRef);
    delete operation.runRef;
  }
  if ("objectRef" in operation) {
    const target = parseObjectRef(String(operation.objectRef));
    if (target.runId === undefined || target.objectIndex === undefined) throw new Error("The object reference is read-only");
    operation.runId = resolveRunRef(context.ids, `run:${target.runId}`);
    operation.objectIndex = target.objectIndex;
    delete operation.objectRef;
  }
  if ("cellRef" in operation) {
    operation.cellParagraphId = resolveParagraphRef(context.ids, operation.cellRef);
    delete operation.cellRef;
  }
  if ("afterBlockRef" in operation) {
    operation.afterBlockId = resolveParagraphRef(context.ids, operation.afterBlockRef);
    delete operation.afterBlockRef;
  }
  if ("at" in operation) operation.at = position(operation.at, context.ids);
  if (Array.isArray(operation.ranges)) {
    operation.ranges = operation.ranges.map((range) => {
      if (!range || typeof range !== "object") throw new Error("Revision range must be an object");
      const value = range as Record<string, unknown>;
      return { blockId: resolveParagraphRef(context.ids, value.blockRef), runId: resolveRunRef(context.ids, value.runRef), start: value.start, end: value.end };
    });
  }
  if (Array.isArray(operation.marks)) {
    operation.marks = operation.marks.map((mark) => {
      if (!mark || typeof mark !== "object") throw new Error("Revision mark must be an object");
      const value = mark as Record<string, unknown>;
      return { blockId: resolveParagraphRef(context.ids, value.blockRef), glyph: value.glyph };
    });
  }

  if (typedKind === "splitParagraph") {
    const [newBlockId, newRunId] = context.allocateIds(2);
    operation.newBlockId = newBlockId;
    operation.newRunId = newRunId;
  } else if (typedKind === "formatRange") {
    const runId = operation.runId as number;
    const run = context.ids.elOf(runId);
    if (!run || localName(run.name) !== "r") throw new Error("Unknown run reference");
    const start = operation.start;
    const end = operation.end;
    if (!Number.isInteger(start) || !Number.isInteger(end)) throw new Error("Format offsets must be integers");
    const length = runWireLength(run);
    if ((start as number) < 0 || (end as number) > length || (end as number) <= (start as number)) throw new Error("Format range is outside the run");
    const allocated = context.allocateIds(3);
    if ((start as number) > 0) operation.beforeId = allocated.shift();
    operation.middleId = allocated.shift();
    if ((end as number) < length) operation.afterId = allocated.shift();
  }

  const nodeCount = idsForOperation(typedKind, operation);
  if (nodeCount) operation.nodeIds = context.allocateIds(nodeCount);

  if (typedKind === "commentRun") {
    operation.author ??= context.provenance.author;
    operation.date ??= context.provenance.now();
    operation.paraId ??= paraId(context);
  } else if (typedKind === "replyComment") {
    operation.author ??= context.provenance.author;
    operation.date ??= context.provenance.now();
    operation.paraIds ??= [paraId(context), paraId(context)];
  } else if (typedKind === "suggestRevision") {
    operation.suggest ??= { author: context.provenance.author, date: context.provenance.now() };
  } else if (SUGGESTABLE_KINDS.has(typedKind)) {
    // Agents ask for a tracked change with a boolean; the author and date the
    // replicas need are stamped here, once, like every other provenance value.
    if (operation.suggest === true) operation.suggest = { author: context.provenance.author, date: context.provenance.now() };
    else delete operation.suggest;
  }

  if (typedKind === "insertImage") {
    const ref = operation.assetRef;
    if (typeof ref !== "string") throw new Error("insertImage requires assetRef");
    const asset = context.asset(ref);
    const prepared = await context.prepareMedia(asset.bytes);
    delete operation.assetRef;
    operation.blobSha = prepared.blobSha;
    operation.bytesLen = prepared.bytesLen;
    operation.ext = imageExtension(asset.mediaType);
    if (prepared.iv) operation.iv = prepared.iv;
  }

  const bodies = (typedKind === "insertText" ? expandInsertText(operation, context) : [operation]) as unknown as IntentBody[];
  for (const body of bodies) {
    const full = { ...body, clientId: "agent-validation", clientSeq: 0, base: 0 } as Intent;
    const error = validateIntent(full);
    if (error) throw new Error(error);
  }
  return bodies;
}

export async function localMedia(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number }> {
  return { blobSha: await sha256Hex(bytes), bytesLen: bytes.length };
}
