import { DocxDocument, localName, runWireLength, serializeXml, type Block, type Paragraph, type Run, type StableIds, type XmlElement } from "@wordinweb/core";
import { INTENT_KINDS, applyIntentScoped, resyncScope, type Intent, type IntentBody, type PresencePosition } from "@wordinweb/collab/client";
import { agentCapabilities, agentOperationSchema, hoistRepeatedSubschemas, type AgentEditCapability } from "./capabilities.js";
import { AGENT_COMPOSE_SCHEMA, composeDocxBytes } from "./compose.js";
import { compileAgentOperation, localMedia } from "./edit.js";
import { SemanticInspector } from "./inspect.js";
import { hunksFromUnifiedDiff, patchOperations } from "./patch.js";
import { projectStory, projectedLines } from "./project.js";
import { blockRef, objectRef, runRef } from "./refs.js";
import { fitInspect, spatialInspect } from "./spatial.js";
import type {
  AgentAsset,
  AgentCollaborativeTarget,
  AgentComposeBlock,
  AgentComposeRequest,
  AgentComposeResult,
  AgentEditRequest,
  AgentEditResult,
  AgentInspectRequest,
  AgentInspectResult,
  AgentOperation,
  AgentPatchRequest,
  AgentPatchResult,
  AgentProjectRequest,
  AgentProjectResult,
  AgentProvenance,
  AgentTool,
} from "./types.js";

function blankDocxBytes(): Uint8Array {
  return composeDocxBytes({ revision: "0", body: [{ type: "paragraph", text: "" }] });
}

function defaultProvenance(input?: AgentProvenance): Required<AgentProvenance> {
  let id = 0;
  return {
    author: input?.author ?? "AI Agent",
    now: input?.now ?? (() => new Date().toISOString()),
    nextId: input?.nextId ?? (() => `${Date.now().toString(16)}${(++id).toString(16)}`),
  };
}

function cloneWithIds(doc: DocxDocument, ids: StableIds): { doc: DocxDocument; ids: StableIds } {
  const sidecar = ids.exportSidecar(doc.editableRoots());
  const clone = DocxDocument.load(doc.save());
  const cloneIds = clone.enableStableIds();
  cloneIds.importSidecar(clone.editableRoots(), sidecar);
  return { doc: clone, ids: cloneIds };
}

const TARGET_SCOPED_REF = /^(?:block:\d+|run:\d+|object:\d+:\d+)$/;
const GLOBAL_REVISION_KINDS = new Set(["mergeParagraph", "insertBookmark", "insertBookmarkRange", "insertCrossRef"]);

function inspectionReferences(value: unknown, refs = new Set<string>()): Set<string> {
  if (typeof value === "string") {
    if (TARGET_SCOPED_REF.test(value)) refs.add(value);
    return refs;
  }
  if (Array.isArray(value)) {
    for (const item of value) inspectionReferences(item, refs);
    return refs;
  }
  if (value && typeof value === "object") {
    for (const item of Object.values(value as Record<string, unknown>)) inspectionReferences(item, refs);
  }
  return refs;
}

function operationReferences(operation: AgentOperation): Set<string> | null {
  if (GLOBAL_REVISION_KINDS.has(operation.kind)) return null;
  const refs = new Set<string>();
  const add = (value: unknown): void => {
    if (typeof value === "string" && TARGET_SCOPED_REF.test(value)) refs.add(value);
  };
  add(operation.blockRef);
  add(operation.runRef);
  add(operation.objectRef);
  add(operation.cellRef);
  add(operation.afterBlockRef);
  if (operation.at && typeof operation.at === "object" && !Array.isArray(operation.at)) {
    const at = operation.at as Record<string, unknown>;
    add(at.blockRef);
    add(at.runRef);
  }
  for (const field of ["ranges", "marks"] as const) {
    if (!Array.isArray(operation[field])) continue;
    for (const item of operation[field]) {
      if (!item || typeof item !== "object" || Array.isArray(item)) continue;
      const target = item as Record<string, unknown>;
      add(target.blockRef);
      add(target.runRef);
    }
  }
  return refs.size > 0 ? refs : null;
}

function editReferences(operations: AgentOperation[]): Set<string> | null {
  const refs = new Set<string>();
  for (const operation of operations) {
    const operationTargets = operationReferences(operation);
    if (!operationTargets) return null;
    for (const ref of operationTargets) refs.add(ref);
  }
  return refs;
}

function referenceElement(ids: StableIds, ref: string): XmlElement | null {
  const id = /^(?:block|run):(\d+)$/.exec(ref)?.[1] ?? /^object:(\d+):\d+$/.exec(ref)?.[1];
  return id === undefined ? null : ids.elOf(Number(id)) ?? null;
}

function referenceScope(doc: DocxDocument, element: XmlElement): XmlElement {
  let current: XmlElement | undefined = element;
  let paragraph: XmlElement | undefined;
  while (current) {
    const name = localName(current.name);
    if (name === "tbl") return current;
    if (name === "p") paragraph = current;
    current = doc.findParentOf(current);
  }
  return paragraph ?? element;
}

function referenceFingerprint(
  doc: DocxDocument,
  inspector: SemanticInspector,
  ref: string,
  xmlByScope: WeakMap<XmlElement, string>,
): string | null {
  const element = referenceElement(doc.enableStableIds(), ref);
  if (!element) return null;
  const scope = referenceScope(doc, element);
  let xml = xmlByScope.get(scope);
  if (xml === undefined) {
    xml = serializeXml(scope);
    xmlByScope.set(scope, xml);
  }
  if (!ref.startsWith("object:")) return xml;
  try {
    return `${xml}\u0000${JSON.stringify(inspector.object(ref).detail)}`;
  } catch {
    return null;
  }
}

function installPreparedImage(doc: DocxDocument, operation: IntentBody, asset?: AgentAsset): void {
  if (operation.kind !== "insertImage" || !asset) return;
  for (const [part, metadata] of doc.pendingMedia) {
    if (metadata.sha === operation.blobSha) {
      if (!doc.installMedia(part, asset.bytes)) throw new Error("Image bytes did not match the registered asset");
      return;
    }
  }
  throw new Error("The image registration was not created");
}

function runsInParagraph(paragraph: XmlElement, ids: StableIds): Array<{ id: number; length: number }> {
  const runs: Array<{ id: number; length: number }> = [];
  const visit = (element: XmlElement): void => {
    if (element !== paragraph && localName(element.name) === "p") return;
    if (localName(element.name) === "r") {
      const id = ids.idOf(element);
      if (id !== undefined) runs.push({ id, length: runWireLength(element) });
      return;
    }
    for (const child of element.children) visit(child);
  };
  visit(paragraph);
  return runs;
}

function paragraphForRun(doc: DocxDocument, runId: number): { blockId: number; runs: Array<{ id: number; length: number }> } | null {
  const ids = doc.enableStableIds();
  const target = ids.elOf(runId);
  if (!target) return null;
  let result: { blockId: number; runs: Array<{ id: number; length: number }> } | null = null;
  const visit = (element: XmlElement, paragraph?: XmlElement): void => {
    if (result) return;
    const current = localName(element.name) === "p" ? element : paragraph;
    if (element === target && current) {
      const blockId = ids.idOf(current);
      if (blockId !== undefined) result = { blockId, runs: runsInParagraph(current, ids) };
      return;
    }
    for (const child of element.children) visit(child, current);
  };
  for (const root of doc.editableRoots()) visit(root);
  return result;
}

function presenceInParagraph(doc: DocxDocument, blockId: number, absoluteOffset: number): PresencePosition | null {
  const ids = doc.enableStableIds();
  const paragraph = ids.elOf(blockId);
  if (!paragraph || localName(paragraph.name) !== "p") return null;
  const runs = runsInParagraph(paragraph, ids);
  if (runs.length === 0) return null;
  let remaining = Math.max(0, absoluteOffset);
  for (const run of runs) {
    if (remaining <= run.length) return { anchor: { blockId, runId: run.id, offset: remaining } };
    remaining -= run.length;
  }
  const last = runs[runs.length - 1];
  return { anchor: { blockId, runId: last.id, offset: last.length } };
}

function absoluteRunOffset(doc: DocxDocument, blockId: number, runId: number, offset: number): number | null {
  const ids = doc.enableStableIds();
  const paragraph = ids.elOf(blockId);
  if (!paragraph || localName(paragraph.name) !== "p") return null;
  let absolute = 0;
  for (const run of runsInParagraph(paragraph, ids)) {
    if (run.id === runId) return absolute + Math.min(Math.max(0, offset), run.length);
    absolute += run.length;
  }
  return null;
}

interface PresenceHint {
  blockId: number;
  absoluteOffset: number;
}

function presenceHint(doc: DocxDocument, operation: IntentBody): PresenceHint | null {
  const value = operation as IntentBody & Record<string, unknown>;
  const range = Array.isArray(value.ranges) ? value.ranges[value.ranges.length - 1] as Record<string, unknown> | undefined : undefined;
  const mark = Array.isArray(value.marks) ? value.marks[value.marks.length - 1] as Record<string, unknown> | undefined : undefined;
  const at = value.at && typeof value.at === "object" ? value.at as Record<string, unknown> : undefined;
  const runId = Number(range?.runId ?? at?.runId ?? value.runId);
  let blockId = Number(range?.blockId ?? at?.blockId ?? mark?.blockId ?? value.blockId ?? value.cellParagraphId ?? value.afterBlockId);
  if (!Number.isInteger(blockId) && Number.isInteger(runId)) blockId = paragraphForRun(doc, runId)?.blockId ?? NaN;
  if (!Number.isInteger(blockId)) return null;
  if (!Number.isInteger(runId)) {
    const paragraph = doc.enableStableIds().elOf(blockId);
    if (!paragraph) return null;
    const absoluteOffset = runsInParagraph(paragraph, doc.enableStableIds()).reduce((total, run) => total + run.length, 0);
    return { blockId, absoluteOffset };
  }
  let offset = Number(range?.end ?? at?.offset ?? value.end);
  if (value.kind === "deleteText") offset = Number(value.start);
  if (!Number.isFinite(offset)) {
    const run = doc.enableStableIds().elOf(runId);
    offset = run ? runWireLength(run) : 0;
  }
  const absolute = absoluteRunOffset(doc, blockId, runId, offset);
  if (absolute === null) return null;
  return {
    blockId,
    absoluteOffset: absolute + (value.kind === "insertText" ? String(value.text ?? "").length : 0),
  };
}

function presenceAfterOperation(doc: DocxDocument, operation: IntentBody, hint: PresenceHint | null): PresencePosition | null {
  if (operation.kind === "splitParagraph") {
    return { anchor: { blockId: operation.newBlockId, runId: operation.newRunId, offset: 0 } };
  }
  if (hint) return presenceInParagraph(doc, hint.blockId, hint.absoluteOffset);
  const value = operation as IntentBody & Record<string, unknown>;
  const nodeIds = Array.isArray(value.nodeIds) ? value.nodeIds : [];
  for (let index = nodeIds.length - 1; index >= 0; index--) {
    const runId = Number(nodeIds[index]);
    if (!Number.isInteger(runId)) continue;
    const paragraph = paragraphForRun(doc, runId);
    if (!paragraph) continue;
    const run = paragraph.runs.find((candidate) => candidate.id === runId);
    if (run) return { anchor: { blockId: paragraph.blockId, runId, offset: run.length } };
  }
  return null;
}

function projectionKey(story: string, mode: string, cursor: string | null, revision: string): string {
  return [revision, story, mode, cursor ?? ""].join("\u0000");
}

function asObject(input: unknown): Record<string, unknown> {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new Error("Tool input must be an object");
  return input as Record<string, unknown>;
}

// The Anthropic Messages API rejects anyOf/oneOf/allOf at the top level of a
// tool's input_schema, so the kind-discriminated union is flattened into one
// object schema. Which fields go with which kind is stated on the fields and
// enforced by the handler, which reports a violation as a normal tool error.
const inspectToolSchema: Record<string, unknown> = {
  type: "object",
  properties: {
    kind: {
      enum: ["overview", "context", "read", "search", "object", "spatial", "fit"],
      description:
        "overview: whole-document shape. context: bulk text for stories. read: one story range. search: find text (requires query). object: one object (requires ref). spatial: page geometry. fit: per-drawing text fit plus page fill — check it after inserting or resizing a drawing, to see whether the drawing's text still fits its box.",
    },
    stories: { type: "array", minItems: 1, maxItems: 100, uniqueItems: true, items: { type: "string", minLength: 1 }, description: "context only" },
    maxBlocks: { type: "integer", minimum: 1, maximum: 200, description: "context or read" },
    maxCharacters: { type: "integer", minimum: 1, maximum: 100000, description: "context or read" },
    include: { type: "array", uniqueItems: true, items: { enum: ["bookmarks", "objects"] }, description: "context only" },
    includeEmpty: { type: "boolean", description: "context only" },
    story: { type: "string", minLength: 1, description: "read only" },
    cursor: {
      type: "object",
      properties: { value: { type: "string", minLength: 1 } },
      required: ["value"],
      additionalProperties: false,
      description: "read only",
    },
    query: { type: "string", minLength: 1, maxLength: 1000, description: "search only, required" },
    maxResults: { type: "integer", minimum: 1, maximum: 500, description: "search only" },
    ref: { type: "string", pattern: "^(object|view):", description: "object only, required" },
    pages: {
      type: "object",
      properties: {
        start: { type: "integer", minimum: 1 },
        count: { type: "integer", minimum: 1, maximum: 100 },
      },
      required: ["start", "count"],
      additionalProperties: false,
      description: "spatial or fit",
    },
    includeOverlaps: { type: "boolean", description: "spatial only" },
  },
  required: ["kind"],
  additionalProperties: false,
};

// The operations union is by far the largest thing the panel sends, and it
// ships on every request of every round. hoistRepeatedSubschemas collapses the
// shapes it spells out more than once into one $defs entry each; the schema a
// reader resolves is unchanged.
const editToolSchema: Record<string, unknown> = hoistRepeatedSubschemas({
  type: "object",
  properties: {
    revision: { type: "string", minLength: 1 },
    operations: {
      type: "array",
      minItems: 1,
      maxItems: 100,
      items: { anyOf: INTENT_KINDS.map(agentOperationSchema) },
    },
  },
  required: ["revision", "operations"],
  additionalProperties: false,
});

const projectionWindowSchema: Record<string, Record<string, unknown>> = {
  story: { type: "string", minLength: 1 },
  mode: { enum: ["text", "md", "outline"] },
  cursor: {
    type: "object",
    properties: { value: { type: "string", minLength: 1 } },
    required: ["value"],
    additionalProperties: false,
  },
  maxBlocks: { type: "integer", minimum: 1, maximum: 2000 },
  maxCharacters: { type: "integer", minimum: 1, maximum: 200000 },
};

const projectToolSchema: Record<string, unknown> = {
  type: "object",
  properties: { ...projectionWindowSchema },
  additionalProperties: false,
};

const patchToolSchema: Record<string, unknown> = {
  type: "object",
  properties: {
    revision: { type: "string", minLength: 1 },
    ...projectionWindowSchema,
    edits: {
      type: "array",
      minItems: 1,
      maxItems: 100,
      items: {
        type: "object",
        properties: {
          startLine: { type: "integer", minimum: 1 },
          endLine: { type: "integer", minimum: 1 },
          newText: { type: "string", maxLength: 100000 },
        },
        required: ["startLine", "endLine", "newText"],
        additionalProperties: false,
      },
    },
    diff: { type: "string", minLength: 1, maxLength: 1000000 },
    suggest: { type: "boolean" },
  },
  required: ["revision"],
  additionalProperties: false,
};

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunk) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunk));
  }
  return btoa(binary);
}

type ComposeStory = "body" | "header" | "footer" | "firstHeader" | "firstFooter";

function composeBlocks(document: DocxDocument, story: ComposeStory): Block[] {
  if (story === "body") return document.sections[0]?.blocks ?? [];
  const section = document.sections[0];
  const header = story === "header" || story === "firstHeader";
  const ref = header
    ? story === "firstHeader" ? section?.props.headerRefs.first : section?.props.headerRefs.default
    : story === "firstFooter" ? section?.props.footerRefs.first : section?.props.footerRefs.default;
  return (header ? document.headers : document.footers).get(ref ?? "")?.blocks ?? [];
}

function directRuns(paragraph: Paragraph): Run[] {
  return paragraph.children.flatMap((child) => child.type === "run" ? [child] : child.runs);
}

function composeTarget(document: DocxDocument, story: ComposeStory, index: number): { blockRef: string; runRef: string } {
  const block = composeBlocks(document, story)[index];
  if (!block || block.type !== "paragraph" || !block.src) throw new Error(`compose ${story} block ${index} has no paragraph target`);
  const run = directRuns(block)[0];
  if (!run?.src) throw new Error(`compose ${story} block ${index} has no run target`);
  const ids = document.enableStableIds();
  const blockId = ids.idOf(block.src);
  const runId = ids.idOf(run.src);
  if (blockId === undefined || runId === undefined) throw new Error(`compose ${story} target has no stable identity`);
  return { blockRef: blockRef(blockId), runRef: runRef(runId) };
}

function composeInsert(block: AgentComposeBlock, target: { blockRef: string; runRef: string }): AgentOperation | null {
  switch (block.type) {
    case "paragraph": return null;
    case "equation": return { kind: "insertMath", runRef: target.runRef, mathText: block.mathText };
    case "chart": return { kind: "insertChart", runRef: target.runRef, chart: block.chart };
    case "smartArt": return { kind: "insertSmartArt", runRef: target.runRef, smartArt: block.smartArt };
    case "image": return { kind: "insertImage", runRef: target.runRef, assetRef: block.assetRef, widthPx: block.widthPx, heightPx: block.heightPx };
    case "shape": return { kind: "insertShape", runRef: target.runRef, preset: block.preset, text: block.text ?? "" };
    case "wordArt": return { kind: "insertWordArt", runRef: target.runRef, text: block.text, preset: block.preset };
    case "pageNumber": return { kind: "insertPageField", runRef: target.runRef, fieldKind: block.fieldKind };
    case "heading":
    case "table":
    case "pageBreak": return null;
  }
}

function composeObjectRef(document: DocxDocument, story: ComposeStory, index: number): string {
  const block = composeBlocks(document, story)[index];
  if (!block || block.type !== "paragraph") throw new Error(`compose ${story} object target is missing`);
  const ids = document.enableStableIds();
  for (const run of directRuns(block)) {
    if (!run.src) continue;
    const contentIndex = run.content.findIndex((content) => content.kind === "image" || content.kind === "anchor" || content.kind === "drawing");
    const runId = ids.idOf(run.src);
    if (contentIndex >= 0 && runId !== undefined) return objectRef(runId, contentIndex);
  }
  throw new Error(`compose ${story} object was not created`);
}

function composeShapeTextTarget(document: DocxDocument, story: ComposeStory, index: number): { blockRef: string; runRef: string } | null {
  const block = composeBlocks(document, story)[index];
  if (!block || block.type !== "paragraph") return null;
  const ids = document.enableStableIds();
  for (const run of directRuns(block)) {
    for (const content of run.content) {
      const blocks = content.kind === "anchor" && content.shape.type === "textbox"
        ? content.shape.blocks
        : content.kind === "drawing" ? content.textbox?.blocks : undefined;
      const paragraph = blocks?.find((candidate): candidate is Paragraph => candidate.type === "paragraph");
      const textRun = paragraph && directRuns(paragraph)[0];
      const blockId = paragraph?.src ? ids.idOf(paragraph.src) : undefined;
      const runId = textRun?.src ? ids.idOf(textRun.src) : undefined;
      if (blockId !== undefined && runId !== undefined) return { blockRef: blockRef(blockId), runRef: runRef(runId) };
    }
  }
  return null;
}

function composeAdjustments(block: AgentComposeBlock, ref: string, textTarget?: { blockRef: string; runRef: string } | null): AgentOperation[] {
  if (block.type === "image") {
    const operations: AgentOperation[] = [];
    if (block.alt !== undefined) operations.push({ kind: "setImageAltText", objectRef: ref, alt: block.alt });
    const wrap = block.wrap ?? (block.position ? "none" : undefined);
    if (wrap) operations.push({ kind: "setImageWrap", objectRef: ref, mode: wrap });
    if (block.position) operations.push({ kind: "setFloatingPagePosition", objectRef: ref, ...block.position });
    return operations;
  }
  if (block.type === "chart" || block.type === "smartArt") {
    return block.widthPx !== undefined && block.heightPx !== undefined
      ? [{ kind: "resizeDrawing", objectRef: ref, widthPx: block.widthPx, heightPx: block.heightPx }]
      : [];
  }
  if (block.type !== "shape" && block.type !== "wordArt") return [];
  const operations: AgentOperation[] = [];
  if (block.widthPx !== undefined && block.heightPx !== undefined) operations.push({ kind: "resizeDrawing", objectRef: ref, widthPx: block.widthPx, heightPx: block.heightPx });
  if (block.position) operations.push({ kind: "setFloatingPagePosition", objectRef: ref, ...block.position });
  if (block.type === "shape" && block.fill !== undefined) operations.push({ kind: "setDrawingFill", objectRef: ref, color: block.fill?.replace(/^#/, "") ?? null });
  if (block.type === "wordArt" && (block.fill !== undefined || block.opacity !== undefined)) {
    operations.push({ kind: "setDrawingWordArtStyle", objectRef: ref, color: (block.fill ?? "2E74B5").replace(/^#/, ""), opacity: block.opacity ?? 1 });
  }
  const wrap = block.wrap ?? (block.type === "shape" ? (block.position ? "none" : "topAndBottom") : undefined);
  if (wrap) operations.push({ kind: "setImageWrap", objectRef: ref, mode: wrap });
  if (block.order) operations.push({ kind: "setDrawingOrder", objectRef: ref, order: block.order });
  if (block.type === "shape" && block.line !== undefined) {
    operations.push(block.line === null
      ? { kind: "setDrawingLineStyle", objectRef: ref, color: null }
      : { kind: "setDrawingLineStyle", objectRef: ref, ...block.line, color: block.line.color.replace(/^#/, "") });
  }
  if (block.type === "shape" && block.textStyle && textTarget) {
    const { align, ...textPatch } = block.textStyle;
    if (Object.keys(textPatch).length) {
      operations.push({
        kind: "formatRun",
        ...textTarget,
        patch: { ...textPatch, ...(textPatch.color ? { color: textPatch.color.replace(/^#/, "") } : {}) },
      });
    }
    if (align) operations.push({ kind: "formatParagraph", blockRef: textTarget.blockRef, align });
  }
  if (block.type === "wordArt" && block.rotation !== undefined) operations.push({ kind: "setDrawingRotation", objectRef: ref, degrees: block.rotation });
  return operations;
}

function componentBlock(block: AgentComposeBlock): boolean {
  return ["equation", "chart", "smartArt", "image", "shape", "wordArt", "pageNumber"].includes(block.type);
}

/** An operation compiled but the engine could not place it. Name the usual
 * cause and the working alternative, so the model is never stranded. The
 * sentence about what survived belongs to the rollback wrapper below, which
 * every write path shares. */
function applyFailure(kind: string): Error {
  const paragraphHint = kind === "splitParagraph"
    ? ' To create paragraphs, one insertText whose text contains "\\n" between paragraphs also works, in one operation.'
    : "";
  return new Error(
    `${kind} could not apply to the referenced document state. `
    + `A reference or offset is likely stale: an earlier operation in this transaction can restructure the runs it touched.${paragraphHint}`,
  );
}

// Every write path here applies to a trial clone and adopts it only once the
// whole request succeeds, so a throw leaves the document byte-for-byte as it
// was. An error has to say so: a caller told to "send the rest" resumes from a
// state that never existed and inserts the earlier operations a second time.
const EDIT_ROLLBACK =
  "NOTHING was applied. The transaction is all or nothing, so the document is unchanged "
  + "and no earlier operation in this transaction landed either. "
  + "Re-inspect the document, then send every operation again.";

const PATCH_ROLLBACK =
  "NOTHING was applied. The patch is all or nothing, so the document is unchanged "
  + "and no earlier hunk in this patch landed either. "
  + "Re-project the story, then send every edit again.";

const COMPOSE_ROLLBACK =
  "NOTHING was created. Compose is all or nothing, so the document is unchanged. "
  + "Send the whole compose request again.";

function rolledBack(error: unknown, resolution: string): Error {
  const message = error instanceof Error ? error.message : String(error);
  return new Error(`${message} ${resolution}`);
}

export class AgentDocument {
  private localDoc: DocxDocument | null;
  private localRevision = 0;
  private readonly target: AgentCollaborativeTarget | null;
  private readonly provenance: Required<AgentProvenance>;
  private readonly addedAssets = new Map<string, AgentAsset>();
  private assetCounter = 0;
  private composeAvailable: boolean;
  private readonly listeners = new Set<(document: DocxDocument, revision: string) => void>();
  private readonly inspectedTargets = new Map<string, Map<string, string>>();
  private readonly projections = new Map<string, AgentProjectResult>();

  private constructor(doc: DocxDocument | null, target: AgentCollaborativeTarget | null, provenance?: AgentProvenance, composeAvailable = false) {
    this.localDoc = doc;
    this.target = target;
    this.provenance = defaultProvenance(provenance);
    this.composeAvailable = composeAvailable;
    doc?.enableStableIds();
  }

  static load(bytes: ArrayBuffer | Uint8Array, options?: { provenance?: AgentProvenance }): AgentDocument {
    return new AgentDocument(DocxDocument.load(bytes), null, options?.provenance);
  }

  static create(options?: { provenance?: AgentProvenance }): AgentDocument {
    return new AgentDocument(DocxDocument.load(blankDocxBytes()), null, options?.provenance, true);
  }

  static connect(target: AgentCollaborativeTarget, options?: { provenance?: AgentProvenance }): AgentDocument {
    return new AgentDocument(null, target, options?.provenance);
  }

  get revision(): string {
    return this.target ? String(this.target.getRevision()) : String(this.localRevision);
  }

  get document(): DocxDocument {
    const doc = this.target ? this.target.getDocument() : this.localDoc;
    if (!doc) throw new Error("The document is not ready");
    doc.enableStableIds();
    return doc;
  }

  subscribe(listener: (document: DocxDocument, revision: string) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  addAsset(bytes: Uint8Array, mediaType: string): string {
    if (!(bytes instanceof Uint8Array) || bytes.length === 0) throw new Error("Asset bytes are required");
    if (typeof mediaType !== "string" || mediaType.length === 0) throw new Error("Asset media type is required");
    const ref = `asset:added:${++this.assetCounter}`;
    this.addedAssets.set(ref, { ref, bytes: new Uint8Array(bytes), mediaType });
    return ref;
  }

  private inspector(doc = this.document, revision = this.revision): SemanticInspector {
    return new SemanticInspector(doc, doc.enableStableIds(), revision);
  }

  /**
   * Read the document the way its READER reads it, whatever is on screen.
   *
   * The model these reads walk is derived per revision view, and markup view
   * keeps deleted runs in it so a reviewer can see the strike-through. Right
   * for a screen, wrong for an agent: the AI panel turns suggesting on for its
   * turn, which forces markup, so a SECOND suggestion projected the pending
   * insertion and the pending deletion interleaved — "The quick brown fox
   * leajumps over the lazy dog" for a document that reads "leaps". The model
   * then either rewrote a string no human had ever seen, or, because deleted
   * runs are not editable text, got "cannot be rewritten across a non-text
   * atom" and could not stack a second change at all.
   *
   * So every model-derived read runs in final view: insertions are text,
   * deletions are gone — the document as it will read once the suggestions are
   * accepted. Only the READS move. The operations they produce address runs by
   * stable id, at offsets within a run's own text; both are properties of the
   * XML and mean the same thing in either view, so applying them needs no
   * switch and the view on screen never flickers.
   */
  private inFinalView<T>(read: () => T): T {
    const doc = this.document;
    const shown = doc.revisionView;
    if (shown === "final") return read();
    doc.setRevisionView("final");
    try {
      return read();
    } finally {
      doc.setRevisionView(shown);
    }
  }

  asset(ref: string): AgentAsset {
    return this.addedAssets.get(ref) ?? this.inspector().asset(ref);
  }

  inspect(request: AgentInspectRequest): AgentInspectResult {
    if (!request || typeof request !== "object") throw new Error("Inspection request is required");
    return this.inFinalView(() => this.inspectFinal(request));
  }

  private inspectFinal(request: AgentInspectRequest): AgentInspectResult {
    const inspector = this.inspector();
    let result: AgentInspectResult;
    switch (request.kind) {
      case "overview": result = inspector.overview(); break;
      case "context": result = inspector.context(request.stories, request.maxBlocks, request.maxCharacters, request.include, request.includeEmpty); break;
      case "read": result = inspector.read(request.story, request.cursor?.value, request.maxBlocks, request.maxCharacters); break;
      case "search": result = inspector.search(request.query, request.maxResults); break;
      case "object": result = inspector.object(request.ref); break;
      case "spatial": result = spatialInspect(this.document, this.document.enableStableIds(), this.revision, inspector, request.pages, request.includeOverlaps); break;
      case "fit": result = fitInspect(this.document, this.document.enableStableIds(), this.revision, inspector, request.pages); break;
      default: {
        const exhaustive: never = request;
        return exhaustive;
      }
    }
    if (this.target) this.recordInspection(result, result.revision, inspector);
    return result;
  }

  private recordInspection(value: unknown, revision: string, inspector: SemanticInspector): void {
    const fingerprints = this.inspectedTargets.get(revision) ?? new Map<string, string>();
    const xmlByScope = new WeakMap<XmlElement, string>();
    for (const ref of inspectionReferences(value)) {
      const fingerprint = referenceFingerprint(this.document, inspector, ref, xmlByScope);
      if (fingerprint !== null) fingerprints.set(ref, fingerprint);
    }
    if (!this.inspectedTargets.has(revision)) {
      this.inspectedTargets.set(revision, fingerprints);
      while (this.inspectedTargets.size > 8) this.inspectedTargets.delete(this.inspectedTargets.keys().next().value!);
    }
  }

  /** Fingerprints are recorded during a final-view read, so they have to be
   * re-taken in one: compared across views, every block holding a tracked
   * change reads as concurrently edited and the patch dies as "stale". */
  private fingerprintsMatch(revision: string, refs: Set<string>): boolean {
    const observed = this.inspectedTargets.get(revision);
    if (!observed) return false;
    return this.inFinalView(() => {
      const inspector = this.inspector();
      const xmlByScope = new WeakMap<XmlElement, string>();
      for (const ref of refs) {
        const before = observed.get(ref);
        const current = referenceFingerprint(this.document, inspector, ref, xmlByScope);
        if (before === undefined || current === null || before !== current) return false;
      }
      return true;
    });
  }

  private revisionMatchesEdit(request: AgentEditRequest): boolean {
    if (request.revision === this.revision) return true;
    if (!this.target) return false;
    const refs = editReferences(request.operations);
    return refs ? this.fingerprintsMatch(request.revision, refs) : false;
  }

  /** Render the document as deterministic text plus the anchor map that owns
   * the rendered-to-wire translation. Read-only: nothing here mutates. */
  project(request: AgentProjectRequest = {}): AgentProjectResult {
    if (!request || typeof request !== "object") throw new Error("Projection request is required");
    return this.inFinalView(() => this.projectFinal(request));
  }

  private projectFinal(request: AgentProjectRequest): AgentProjectResult {
    const doc = this.document;
    const result = projectStory(
      doc,
      doc.enableStableIds(),
      this.revision,
      request.story ?? "body",
      request.mode ?? "md",
      request.cursor?.value,
      request.maxBlocks,
      request.maxCharacters,
    );
    const key = projectionKey(result.story, result.mode, result.window.cursor, result.revision);
    this.projections.set(key, result);
    while (this.projections.size > 8) this.projections.delete(this.projections.keys().next().value!);
    // Anchors carry every block and run the window addresses, so the patch
    // guard can check exactly the blocks a hunk touches and let concurrent
    // edits elsewhere through — on local sessions as well as collaborative.
    this.recordInspection(result.anchors, result.revision, this.inspector());
    return result;
  }

  async patch(request: AgentPatchRequest): Promise<AgentPatchResult> {
    try {
      return await this.patchWindow(request);
    } catch (error) {
      throw rolledBack(error, PATCH_ROLLBACK);
    }
  }

  private async patchWindow(request: AgentPatchRequest): Promise<AgentPatchResult> {
    if (!request || typeof request !== "object" || typeof request.revision !== "string") {
      throw new Error("A patch needs the revision its projection came from");
    }
    const hasEdits = Array.isArray(request.edits) && request.edits.length > 0;
    const hasDiff = typeof request.diff === "string";
    if (hasEdits === hasDiff) throw new Error("A patch needs either edits or diff");
    const story = request.story ?? "body";
    const mode = request.mode ?? "md";
    const cursor = request.cursor?.value ?? null;

    // The hunks are line numbers into the window the agent read, so patch the
    // window it read: the cached projection, or a fresh one at the current
    // revision. A window from a revision this document never projected cannot
    // be placed at all.
    const projection = this.projections.get(projectionKey(story, mode, cursor, request.revision))
      ?? (request.revision === this.revision
        ? this.project({ story, mode, cursor: request.cursor, maxBlocks: request.maxBlocks, maxCharacters: request.maxCharacters })
        : undefined);
    if (!projection) throw new Error("The document revision is stale");

    const lines = projectedLines(projection);
    const hunks = hasEdits ? request.edits! : hunksFromUnifiedDiff(request.diff!, lines.map((line) => line.text));
    const operations = patchOperations(lines, hunks, mode, request.suggest === true);

    if (request.revision !== this.revision) {
      const touched = new Set<string>();
      for (const hunk of hunks) {
        for (const line of lines.slice(hunk.startLine - 1, hunk.endLine)) if (line.anchor.blockRef) touched.add(line.anchor.blockRef);
      }
      if (!this.fingerprintsMatch(request.revision, touched)) throw new Error("The document revision is stale");
    }

    // Refresh the window the agent actually held, not the one this request's
    // budgets would produce, so the returned projection re-anchors the same
    // region under a cursor the new revision accepts.
    const refresh = (): AgentProjectResult => this.project({
      story,
      mode,
      ...(projection.window.startBlock > 0 ? { cursor: { value: `docmd:${this.revision}:${projection.window.startBlock}` } } : {}),
      maxBlocks: projection.window.maxBlocks,
      maxCharacters: projection.window.maxCharacters,
    });
    if (operations.length === 0) {
      return { revision: this.revision, status: this.target ? "submitted" : "applied", operations: [], projection: refresh() };
    }
    const result = await this.editTransaction({ revision: this.revision, operations });
    return { ...result, projection: refresh() };
  }

  capabilities(category?: AgentEditCapability["category"], kind?: Intent["kind"]): ReturnType<typeof agentCapabilities> {
    return agentCapabilities(category, kind);
  }

  async compose(request: AgentComposeRequest): Promise<AgentComposeResult> {
    try {
      return await this.composeDocument(request);
    } catch (error) {
      throw rolledBack(error, COMPOSE_ROLLBACK);
    }
  }

  private async composeDocument(request: AgentComposeRequest): Promise<AgentComposeResult> {
    if (!this.composeAvailable || this.target) throw new Error("Composition is available on a new AgentDocument.create() document");
    if (request.revision !== this.revision) throw new Error("The document revision is stale");
    const trial = new AgentDocument(DocxDocument.load(composeDocxBytes(request)), null, this.provenance);
    for (const [ref, asset] of this.addedAssets) trial.addedAssets.set(ref, { ...asset, bytes: new Uint8Array(asset.bytes) });
    trial.assetCounter = this.assetCounter;

    const operations: AgentOperation[] = request.page ? [{ kind: "setPageLayout", patch: request.page }] : [];
    const stories: Array<[ComposeStory, AgentComposeBlock[]]> = [
      ["body", request.body],
      ...(request.header ? [["header", request.header] as [ComposeStory, AgentComposeBlock[]]] : []),
      ...(request.footer ? [["footer", request.footer] as [ComposeStory, AgentComposeBlock[]]] : []),
      ...(request.firstHeader ? [["firstHeader", request.firstHeader] as [ComposeStory, AgentComposeBlock[]]] : []),
      ...(request.firstFooter ? [["firstFooter", request.firstFooter] as [ComposeStory, AgentComposeBlock[]]] : []),
    ];
    for (const [story, blocks] of stories) {
      blocks.forEach((block, index) => {
        if (block.type === "paragraph" || block.type === "heading" || block.type === "table" || block.type === "pageBreak") return;
        const operation = composeInsert(block, composeTarget(trial.document, story, index));
        if (operation) operations.push(operation);
      });
    }
    for (let offset = 0; offset < operations.length; offset += 100) {
      await trial.editTransaction({ revision: trial.revision, operations: operations.slice(offset, offset + 100) });
    }

    const adjustments: AgentOperation[] = [];
    for (const [story, blocks] of stories) {
      blocks.forEach((block, index) => {
        if (block.type !== "image" && block.type !== "shape" && block.type !== "wordArt" && block.type !== "chart" && block.type !== "smartArt") return;
        adjustments.push(...composeAdjustments(
          block,
          composeObjectRef(trial.document, story, index),
          block.type === "shape" && block.textStyle ? composeShapeTextTarget(trial.document, story, index) : null,
        ));
      });
    }
    for (let offset = 0; offset < adjustments.length; offset += 100) {
      await trial.editTransaction({ revision: trial.revision, operations: adjustments.slice(offset, offset + 100) });
    }

    this.localDoc = trial.document;
    this.localRevision++;
    this.composeAvailable = false;
    for (const listener of this.listeners) listener(this.localDoc, this.revision);
    const blocks = stories.reduce((sum, [, content]) => sum + content.length, 0);
    const components = stories.reduce((sum, [, content]) => sum + content.filter(componentBlock).length, 0);
    const inspector = this.inspector();
    const catalog = inspector.objectCatalog();
    return {
      revision: this.revision,
      status: "applied",
      blocks,
      components,
      overview: inspector.overview(),
      createdObjects: catalog.items,
      createdObjectsTruncated: catalog.truncated,
    };
  }

  async edit(request: AgentEditRequest): Promise<AgentEditResult> {
    try {
      return await this.editTransaction(request);
    } catch (error) {
      throw rolledBack(error, EDIT_ROLLBACK);
    }
  }

  /** The transaction itself. patch() and compose() call this rather than
   * edit(), so each states the rollback in its own caller's terms instead of
   * telling a patch author to resend "operations" it never sent. */
  private async editTransaction(request: AgentEditRequest): Promise<AgentEditResult> {
    if (!request || typeof request !== "object" || !Array.isArray(request.operations) || request.operations.length === 0) throw new Error("At least one edit operation is required");
    if (request.operations.length > 100) throw new Error("A transaction can contain at most 100 operations");
    if (!this.revisionMatchesEdit(request)) throw new Error("The document revision is stale");
    return this.target ? this.editCollaborative(request) : this.editLocal(request);
  }

  private async editLocal(request: AgentEditRequest): Promise<AgentEditResult> {
    const current = this.document;
    const currentIds = current.enableStableIds();
    const trial = cloneWithIds(current, currentIds);
    let nextId = trial.ids.nextId();
    // Skip ids the trial already holds: applying a suggested insertion mid
    // transaction assigns table-counter ids to the runs it creates, and a
    // carried id that collides with one of those is a hard reject.
    const allocateIds = (count: number): number[] => {
      const out: number[] = [];
      while (out.length < count) {
        if (!trial.ids.elOf(nextId)) out.push(nextId);
        nextId++;
      }
      return out;
    };
    const kinds: string[] = [];
    for (const input of request.operations) {
      const sourceAsset = typeof input.assetRef === "string" ? this.asset(input.assetRef) : undefined;
      const operations = await compileAgentOperation(input, {
        doc: trial.doc,
        ids: trial.ids,
        allocateIds,
        provenance: this.provenance,
        asset: (ref) => this.asset(ref),
        prepareMedia: localMedia,
      });
      for (const operation of operations) {
        const full = { ...operation, clientId: "local-agent", clientSeq: 0, base: 0 } as Intent;
        const result = applyIntentScoped(trial.doc, trial.ids, full);
        if (!result.applied) throw applyFailure(operation.kind);
        resyncScope(trial.doc, trial.ids, result);
        installPreparedImage(trial.doc, operation, sourceAsset);
        kinds.push(operation.kind);
      }
    }
    this.localDoc = trial.doc;
    this.localRevision++;
    this.composeAvailable = false;
    for (const listener of this.listeners) listener(trial.doc, this.revision);
    return { revision: this.revision, status: "applied", operations: kinds, connection: "local" };
  }

  private async editCollaborative(request: AgentEditRequest): Promise<AgentEditResult> {
    const target = this.target;
    if (!target) throw new Error("Collaborative target is unavailable");
    const live = this.document;
    const ids = live.enableStableIds();
    const trial = cloneWithIds(live, ids);
    const compiled: Array<{ operation: IntentBody; asset?: AgentAsset }> = [];
    let presence: PresencePosition | null = null;
    for (const input of request.operations) {
      const sourceAsset = typeof input.assetRef === "string" ? this.asset(input.assetRef) : undefined;
      const operations = await compileAgentOperation(input, {
        doc: trial.doc,
        ids: trial.ids,
        // Drop allocated ids the trial already holds (allocators never reuse
        // a number, so a skipped id just leaks): applying a suggested
        // insertion mid transaction assigns table-counter ids to the runs it
        // creates, and a carried id colliding with one is a hard reject.
        allocateIds: (count) => {
          const out: number[] = [];
          while (out.length < count) {
            for (const id of target.allocateIds(count - out.length)) {
              if (!trial.ids.elOf(id)) out.push(id);
            }
          }
          return out;
        },
        provenance: this.provenance,
        asset: (ref) => this.asset(ref),
        prepareMedia: async (bytes) => {
          if (!target.uploadMedia) throw new Error("This collaborative target cannot upload image assets");
          const prepared = await target.uploadMedia(bytes);
          if (!prepared) throw new Error("The collaborative media relay refused the image asset");
          return prepared;
        },
      });
      for (const operation of operations) {
        const hint = presenceHint(trial.doc, operation);
        const result = applyIntentScoped(trial.doc, trial.ids, { ...operation, clientId: "agent-trial", clientSeq: 0, base: 0 } as Intent);
        if (!result.applied) throw applyFailure(operation.kind);
        resyncScope(trial.doc, trial.ids, result);
        installPreparedImage(trial.doc, operation, sourceAsset);
        presence = presenceAfterOperation(trial.doc, operation, hint) ?? presence;
        compiled.push({ operation, asset: sourceAsset });
      }
    }
    if (!this.revisionMatchesEdit(request)) throw new Error("The document revision is stale");
    for (const { operation } of compiled) await target.submit(operation);
    if (presence) target.setPresence?.(presence);
    return {
      revision: this.revision,
      status: "submitted",
      operations: compiled.map(({ operation }) => operation.kind),
      connection: target.getConnectionState?.() ?? "live",
    };
  }

  save(): Uint8Array {
    return this.document.save();
  }

  tools(): AgentTool[] {
    // Compose exists only on a fresh AgentDocument.create() document.
    // Advertising it on a connected or loaded document hands the model a
    // guaranteed-error call, so the tool is simply absent there.
    const composeTool: AgentTool = {
      name: "word_document_compose",
      description: "Create a complete new document from schema-enforced blocks, tables, media, native objects, headers, and footers.",
      inputSchema: AGENT_COMPOSE_SCHEMA,
      execute: async (input) => this.compose(asObject(input) as unknown as AgentComposeRequest),
    };
    return [
      {
        name: "word_document_capabilities",
        description: "List schema-enforced document edit operations. Filter by category to keep context small.",
        inputSchema: {
          type: "object",
          properties: {
            category: { enum: ["text", "paragraph", "review", "table", "insert", "drawing", "math", "document"] },
            kind: { enum: INTENT_KINDS },
          },
          additionalProperties: false,
        },
        execute: async (input) => {
          const value = asObject(input);
          return this.capabilities(value.category as AgentEditCapability["category"] | undefined, value.kind as Intent["kind"] | undefined);
        },
      },
      ...(this.composeAvailable && !this.target ? [composeTool] : []),
      {
        name: "word_document_inspect",
        description: "Inspect compact bulk text context, overview, a detailed story range, search results, one object, or page geometry.",
        inputSchema: inspectToolSchema,
        execute: async (input) => this.inspect(asObject(input) as unknown as AgentInspectRequest),
      },
      {
        name: "word_document_edit",
        description: "Apply a validated document transaction against an observed revision.",
        inputSchema: editToolSchema,
        execute: async (input) => this.edit(asObject(input) as unknown as AgentEditRequest),
      },
      {
        name: "word_document_project",
        description: "Project a story as deterministic text or structural markdown with a line anchor map, windowed by cursor.",
        inputSchema: projectToolSchema,
        execute: async (input) => this.project(asObject(input) as unknown as AgentProjectRequest),
      },
      {
        name: "word_document_patch",
        description: "Apply line-range edits or a unified diff written against a projection window.",
        inputSchema: patchToolSchema,
        execute: async (input) => this.patch(asObject(input) as unknown as AgentPatchRequest),
      },
      {
        name: "word_document_asset",
        description: "Read image or embedded asset bytes referenced by document inspection.",
        inputSchema: { type: "object", properties: { ref: { type: "string" } }, required: ["ref"], additionalProperties: false },
        execute: async (input) => {
          const asset = this.asset(String(asObject(input).ref));
          return { ref: asset.ref, mediaType: asset.mediaType, base64: bytesToBase64(asset.bytes) };
        },
      },
      {
        name: "word_document_save",
        description: "Serialize the current document as DOCX bytes.",
        inputSchema: { type: "object", additionalProperties: false },
        execute: async () => ({ mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", base64: bytesToBase64(this.save()) }),
      },
    ];
  }
}

export interface CollabSessionLike {
  doc: DocxDocument | null;
  version: number;
  submitOp(operation: IntentBody): void;
  setPresence?(position: PresencePosition | null): void;
  allocIds(count: number): number[];
  uploadMedia?(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number; iv?: string } | null>;
  connection: "live" | "reconnecting" | "lost";
}

/** Adapt the existing browser CollabSession without adding a React dependency.
 * The accessor must return the latest hook value after each render. */
export function collaborativeAgentTarget(getSession: () => CollabSessionLike): AgentCollaborativeTarget {
  return {
    getDocument: () => getSession().doc,
    getRevision: () => getSession().version,
    allocateIds: (count) => getSession().allocIds(count),
    submit: (operation) => getSession().submitOp(operation),
    setPresence: (position) => getSession().setPresence?.(position),
    uploadMedia: (bytes) => getSession().uploadMedia?.(bytes) ?? Promise.resolve(null),
    getConnectionState: () => {
      const state = getSession().connection;
      return state === "live" ? "live" : state === "reconnecting" ? "reconnecting" : "offline";
    },
  };
}
