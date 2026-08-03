import { DocxDocument, type Block, type Paragraph, type Run, type StableIds } from "@wordinweb/core";
import { INTENT_KINDS, applyIntentScoped, resyncScope, type Intent, type IntentBody } from "@wordinweb/collab/client";
import { agentCapabilities, agentOperationSchema, type AgentEditCapability } from "./capabilities.js";
import { AGENT_COMPOSE_SCHEMA, composeDocxBytes } from "./compose.js";
import { compileAgentOperation, localMedia } from "./edit.js";
import { SemanticInspector } from "./inspect.js";
import { blockRef, objectRef, runRef } from "./refs.js";
import { spatialInspect } from "./spatial.js";
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

function asObject(input: unknown): Record<string, unknown> {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new Error("Tool input must be an object");
  return input as Record<string, unknown>;
}

const inspectToolSchema: Record<string, unknown> = {
  anyOf: [
    {
      type: "object",
      properties: { kind: { const: "overview" } },
      required: ["kind"],
      additionalProperties: false,
    },
    {
      type: "object",
      properties: {
        kind: { const: "read" },
        story: { type: "string", minLength: 1 },
        cursor: {
          type: "object",
          properties: { value: { type: "string", minLength: 1 } },
          required: ["value"],
          additionalProperties: false,
        },
        maxBlocks: { type: "integer", minimum: 1, maximum: 200 },
        maxCharacters: { type: "integer", minimum: 1, maximum: 100000 },
      },
      required: ["kind"],
      additionalProperties: false,
    },
    {
      type: "object",
      properties: {
        kind: { const: "search" },
        query: { type: "string", minLength: 1, maxLength: 1000 },
        maxResults: { type: "integer", minimum: 1, maximum: 500 },
      },
      required: ["kind", "query"],
      additionalProperties: false,
    },
    {
      type: "object",
      properties: { kind: { const: "object" }, ref: { type: "string", pattern: "^(object|view):" } },
      required: ["kind", "ref"],
      additionalProperties: false,
    },
    {
      type: "object",
      properties: {
        kind: { const: "spatial" },
        pages: {
          type: "object",
          properties: {
            start: { type: "integer", minimum: 1 },
            count: { type: "integer", minimum: 1, maximum: 100 },
          },
          required: ["start", "count"],
          additionalProperties: false,
        },
        includeOverlaps: { type: "boolean" },
      },
      required: ["kind"],
      additionalProperties: false,
    },
  ],
};

const editToolSchema: Record<string, unknown> = {
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

export class AgentDocument {
  private localDoc: DocxDocument | null;
  private localRevision = 0;
  private readonly target: AgentCollaborativeTarget | null;
  private readonly provenance: Required<AgentProvenance>;
  private readonly addedAssets = new Map<string, AgentAsset>();
  private assetCounter = 0;
  private composeAvailable: boolean;
  private readonly listeners = new Set<(document: DocxDocument, revision: string) => void>();

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

  asset(ref: string): AgentAsset {
    return this.addedAssets.get(ref) ?? this.inspector().asset(ref);
  }

  inspect(request: AgentInspectRequest): AgentInspectResult {
    if (!request || typeof request !== "object") throw new Error("Inspection request is required");
    const inspector = this.inspector();
    switch (request.kind) {
      case "overview": return inspector.overview();
      case "read": return inspector.read(request.story, request.cursor?.value, request.maxBlocks, request.maxCharacters);
      case "search": return inspector.search(request.query, request.maxResults);
      case "object": return inspector.object(request.ref);
      case "spatial": return spatialInspect(this.document, this.document.enableStableIds(), this.revision, inspector, request.pages, request.includeOverlaps);
      default: {
        const exhaustive: never = request;
        return exhaustive;
      }
    }
  }

  capabilities(category?: AgentEditCapability["category"], kind?: Intent["kind"]): ReturnType<typeof agentCapabilities> {
    return agentCapabilities(category, kind);
  }

  async compose(request: AgentComposeRequest): Promise<AgentComposeResult> {
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
      await trial.edit({ revision: trial.revision, operations: operations.slice(offset, offset + 100) });
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
      await trial.edit({ revision: trial.revision, operations: adjustments.slice(offset, offset + 100) });
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
    if (!request || typeof request !== "object" || !Array.isArray(request.operations) || request.operations.length === 0) throw new Error("At least one edit operation is required");
    if (request.operations.length > 100) throw new Error("A transaction can contain at most 100 operations");
    if (request.revision !== this.revision) throw new Error("The document revision is stale");
    return this.target ? this.editCollaborative(request) : this.editLocal(request);
  }

  private async editLocal(request: AgentEditRequest): Promise<AgentEditResult> {
    const current = this.document;
    const currentIds = current.enableStableIds();
    const trial = cloneWithIds(current, currentIds);
    let nextId = trial.ids.nextId();
    const allocateIds = (count: number) => Array.from({ length: count }, () => nextId++);
    const kinds: string[] = [];
    for (const input of request.operations) {
      const sourceAsset = typeof input.assetRef === "string" ? this.asset(input.assetRef) : undefined;
      const operation = await compileAgentOperation(input, {
        doc: trial.doc,
        ids: trial.ids,
        allocateIds,
        provenance: this.provenance,
        asset: (ref) => this.asset(ref),
        prepareMedia: localMedia,
      });
      const full = { ...operation, clientId: "local-agent", clientSeq: 0, base: 0 } as Intent;
      const result = applyIntentScoped(trial.doc, trial.ids, full);
      if (!result.applied) throw new Error(`${operation.kind} could not apply to the referenced document state`);
      resyncScope(trial.doc, trial.ids, result);
      installPreparedImage(trial.doc, operation, sourceAsset);
      kinds.push(operation.kind);
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
    for (const input of request.operations) {
      const sourceAsset = typeof input.assetRef === "string" ? this.asset(input.assetRef) : undefined;
      const operation = await compileAgentOperation(input, {
        doc: trial.doc,
        ids: trial.ids,
        allocateIds: (count) => target.allocateIds(count),
        provenance: this.provenance,
        asset: (ref) => this.asset(ref),
        prepareMedia: async (bytes) => {
          if (!target.uploadMedia) throw new Error("This collaborative target cannot upload image assets");
          const prepared = await target.uploadMedia(bytes);
          if (!prepared) throw new Error("The collaborative media relay refused the image asset");
          return prepared;
        },
      });
      const result = applyIntentScoped(trial.doc, trial.ids, { ...operation, clientId: "agent-trial", clientSeq: 0, base: 0 } as Intent);
      if (!result.applied) throw new Error(`${operation.kind} could not apply to the referenced document state`);
      resyncScope(trial.doc, trial.ids, result);
      installPreparedImage(trial.doc, operation, sourceAsset);
      compiled.push({ operation, asset: sourceAsset });
    }
    if (request.revision !== this.revision) throw new Error("The document changed while the transaction was validated");
    for (const { operation } of compiled) await target.submit(operation);
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
      {
        name: "word_document_compose",
        description: "Create a complete new document from schema-enforced blocks, tables, media, native objects, headers, and footers.",
        inputSchema: AGENT_COMPOSE_SCHEMA,
        execute: async (input) => this.compose(asObject(input) as unknown as AgentComposeRequest),
      },
      {
        name: "word_document_inspect",
        description: "Inspect a document progressively by overview, story range, search, object, or page geometry.",
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
    uploadMedia: (bytes) => getSession().uploadMedia?.(bytes) ?? Promise.resolve(null),
    getConnectionState: () => {
      const state = getSession().connection;
      return state === "live" ? "live" : state === "reconnecting" ? "reconnecting" : "offline";
    },
  };
}
