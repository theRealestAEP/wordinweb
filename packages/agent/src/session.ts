import { DocxDocument, type XmlElement } from "@wordinweb/core";
import { applyIntentScoped, resyncScope, sha256Hex, unionScopes, type Intent, type IntentBody, type Scope } from "@wordinweb/collab/client";
import { validateIntent } from "@wordinweb/collab/server";
import type { AgentCollaborativeTarget } from "./types.js";

/**
 * A canonical document session for one process. It lets a browser editor and
 * an AgentDocument share one DocxDocument without a server or a DOM dependency.
 */
export class LocalDocumentSession implements AgentCollaborativeTarget {
  readonly doc: DocxDocument;
  private revision = 0;
  private renderRevision = 0;
  private clientSequence = 0;
  private nextId: number;
  private dirtyScope: Scope | null = null;
  private readonly listeners = new Set<() => void>();
  private readonly uploadedMedia = new Map<string, Uint8Array>();

  constructor(source: ArrayBuffer | Uint8Array | DocxDocument) {
    this.doc = source instanceof DocxDocument ? source : DocxDocument.load(source);
    this.nextId = this.doc.enableStableIds().nextId();
  }

  get version(): number {
    return this.revision;
  }

  get renderSignal(): number {
    return this.renderRevision;
  }

  getDocument(): DocxDocument {
    return this.doc;
  }

  getRevision(): number {
    return this.revision;
  }

  getConnectionState(): "local" {
    return "local";
  }

  allocateIds(count: number): number[] {
    if (!Number.isInteger(count) || count < 0) throw new Error("ID count must be a non-negative integer");
    // A suggested insertion creates nodes whose ids come from the id table's
    // own counter (scoped reparse), not from this allocator — resync before
    // handing out numbers, or the next carried id collides with them.
    this.nextId = Math.max(this.nextId, this.doc.enableStableIds().nextId());
    return Array.from({ length: count }, () => this.nextId++);
  }

  allocIds = (count: number): number[] => this.allocateIds(count);

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  takeRenderScope = (): Scope | null => {
    const scope = this.dirtyScope;
    this.dirtyScope = null;
    return scope;
  };

  submit = (operation: IntentBody): void => {
    const intent = this.fullIntent(operation);
    const error = validateIntent(intent);
    if (error) throw new Error(error);
    const ids = this.doc.enableStableIds();
    const result = applyIntentScoped(this.doc, ids, intent);
    if (!result.applied) throw new Error(`${operation.kind} could not apply to the referenced document state`);
    resyncScope(this.doc, ids, result);
    this.installUploadedImage(operation);
    this.dirtyScope = unionScopes(this.dirtyScope, result);
    this.revision++;
    this.renderRevision++;
    for (const listener of this.listeners) listener();
  };

  submitOp = (operation: IntentBody): void => this.submit(operation);

  /** Record an operation that DocxView already applied to this same object. */
  submitPreApplied = (operation: IntentBody): void => {
    const error = validateIntent(this.fullIntent(operation));
    if (error) throw new Error(error);
    this.revision++;
  };

  async uploadMedia(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number }> {
    const blobSha = await sha256Hex(bytes);
    this.uploadedMedia.set(blobSha, new Uint8Array(bytes));
    return { blobSha, bytesLen: bytes.length };
  }

  private fullIntent(operation: IntentBody): Intent {
    return { ...operation, clientId: "local-document-session", clientSeq: ++this.clientSequence, base: this.revision } as Intent;
  }

  private installUploadedImage(operation: IntentBody): void {
    if (operation.kind !== "insertImage") return;
    const bytes = this.uploadedMedia.get(operation.blobSha);
    if (!bytes) return;
    for (const [part, metadata] of this.doc.pendingMedia) {
      if (metadata.sha !== operation.blobSha) continue;
      this.doc.installMedia(part, bytes);
      this.uploadedMedia.delete(operation.blobSha);
      return;
    }
  }
}

export interface LocalDocumentViewBinding {
  subscribe(listener: () => void): () => void;
  getSnapshot(): number;
  doc: DocxDocument;
  submit(operation: IntentBody): void;
  submitOp(operation: IntentBody): void;
  allocIds(count: number): number[];
  uploadMedia(bytes: Uint8Array): Promise<{ blobSha: string; bytesLen: number }>;
  takeRenderScope():
    | { kind: "doc" }
    | { kind: "block"; blocks: XmlElement[] }
    | { kind: "split"; before: XmlElement; after: XmlElement }
    | null;
}

export function localDocumentViewBinding(session: LocalDocumentSession): LocalDocumentViewBinding {
  return {
    subscribe: (listener) => session.subscribe(listener),
    getSnapshot: () => session.renderSignal,
    doc: session.doc,
    submit: session.submitPreApplied,
    submitOp: session.submitOp,
    allocIds: session.allocIds,
    uploadMedia: (bytes) => session.uploadMedia(bytes),
    takeRenderScope: session.takeRenderScope,
  };
}
