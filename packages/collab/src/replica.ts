import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, LogEntry, idempotencyKey } from "./intents.js";
import { transformIntent } from "./transform.js";
import { applyIntent } from "./apply.js";

/**
 * A client-side replica with optimistic apply and rollback-replay (plan doc
 * 03 Stage B). It applies local intents immediately against its live
 * document, and reconciles when the server's canonical broadcasts arrive:
 *
 *  - a broadcast that is the canonical form of our own oldest pending intent
 *    advances the confirmed baseline and drops that pending entry;
 *  - a remote broadcast triggers restore-to-confirmed, apply the canonical
 *    broadcasts, then replay our still-pending intents on top (transformed
 *    through the same shared function the server used).
 *
 * The confirmed baseline is snapshotted with save()/load() — correct and
 * simple; save() is side-effect-free (checkpoint-purity invariant), so
 * snapshotting never perturbs live state. A production client swaps in the
 * lighter confirmed-buffer scheme from doc 03; the observable behavior is the
 * same and is what these tests pin.
 */
export class ClientReplica {
  doc: DocxDocument;
  ids: StableIds;
  /** Server seq the confirmed baseline reflects. */
  confirmedSeq = 0;
  private confirmedBytes: Uint8Array;
  private confirmedSidecar: ReturnType<StableIds["exportSidecar"]>;
  private pending: Intent[] = [];

  constructor(bytes: Uint8Array) {
    this.doc = DocxDocument.load(bytes);
    this.ids = this.doc.enableStableIds();
    this.confirmedBytes = this.doc.save();
    this.confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
  }

  /** Apply a locally produced intent optimistically and enqueue it as pending.
   * The intent's `base` must be the current confirmed seq. */
  submitLocal(intent: Intent): void {
    applyIntent(this.doc, this.ids, intent);
    if (intent.kind === "splitParagraph") this.resync();
    this.pending.push(intent);
  }

  /** Receive the server's canonical broadcast entries (in seq order) and
   * reconcile. Idempotent for already-seen seqs. */
  receive(entries: LogEntry[]): void {
    const fresh = entries.filter((e) => e.seq > this.confirmedSeq);
    if (fresh.length === 0) return;

    const ownAhead = fresh.some((e) => !isOurs(e, this.pending));
    if (!ownAhead) {
      // Every fresh entry is the canonical echo of our own pending, in order:
      // fast-forward the confirmed baseline, drop matching pending.
      for (const e of fresh) this.advanceConfirmed(e);
      this.snapshotConfirmed();
      return;
    }

    // A remote edit interleaves: restore to confirmed, apply the canonical
    // entries, then replay remaining pending on top.
    this.restoreConfirmed();
    for (const e of fresh) this.advanceConfirmed(e);
    this.snapshotConfirmed();
    this.replayPending();
  }

  private advanceConfirmed(e: LogEntry): void {
    if (e.kind === "applied") {
      applyIntent(this.doc, this.ids, e.intent);
      if (e.intent.kind === "splitParagraph") this.resync();
    }
    // Drop a matching pending intent (ours, now confirmed).
    const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
    this.pending = this.pending.filter((p) => idempotencyKey(p) !== key);
    this.confirmedSeq = e.seq;
  }

  private replayPending(): void {
    // Each pending intent was produced against the pre-broadcast state; replay
    // it transformed against the entries confirmed since its base.
    const stillPending = this.pending;
    this.pending = [];
    for (const p of stillPending) {
      const applied = applyIntent(this.doc, this.ids, p);
      if (p.kind === "splitParagraph") this.resync();
      if (applied) this.pending.push(p);
    }
  }

  private restoreConfirmed(): void {
    this.doc = DocxDocument.load(this.confirmedBytes);
    this.ids = this.doc.enableStableIds();
    // Reproduce the exact id table via the sidecar — parse-order alone would
    // renumber split-created carried ids and break address resolution
    // (plan round-2 F1).
    this.ids.importSidecar(this.doc.editableRoots(), this.confirmedSidecar);
  }

  private snapshotConfirmed(): void {
    this.confirmedBytes = this.doc.save();
    this.confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
  }

  private resync(): void {
    this.doc.refresh();
    this.ids.assignFromRoots(this.doc.editableRoots());
  }
}

function isOurs(e: LogEntry, pending: Intent[]): boolean {
  const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
  return pending.some((p) => idempotencyKey(p) === key);
}

/** Re-export the transform for clients that need to pre-transform a pending
 * intent's addresses for display before the canonical form returns. */
export { transformIntent };
