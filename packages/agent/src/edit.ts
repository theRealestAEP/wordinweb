import { localName, runWireLength, type DocxDocument, type StableIds } from "@wordinweb/core";
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
    case "insertTable": {
      const rows = Number(operation.rows);
      const cols = Number(operation.cols);
      return Number.isInteger(rows) && Number.isInteger(cols) && rows > 0 && cols > 0 ? rows * cols * 2 + 8 : 8;
    }
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

export async function compileAgentOperation(input: AgentOperation, context: CompileContext): Promise<IntentBody> {
  const operation = cloneOperation(input);
  const kind = operation.kind;
  if (typeof kind !== "string" || !(kind in AGENT_EDIT_CAPABILITIES)) throw new Error("Unknown edit operation");
  const typedKind = kind as Intent["kind"];
  for (const field of INTERNAL_FIELDS) if (field in operation) throw new Error(`${field} is internal; use opaque references`);
  const shapeError = validateAgentOperationShape(input);
  if (shapeError) throw new Error(shapeError);
  requireFields(operation, AGENT_EDIT_CAPABILITIES[typedKind].required);

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
  } else if (typedKind === "insertText" || typedKind === "splitParagraph") {
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

  const body = operation as unknown as IntentBody;
  const full = { ...body, clientId: "agent-validation", clientSeq: 0, base: 0 } as Intent;
  const error = validateIntent(full);
  if (error) throw new Error(error);
  return body;
}

export async function localMedia(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number }> {
  return { blobSha: await sha256Hex(bytes), bytesLen: bytes.length };
}
