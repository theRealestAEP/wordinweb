import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, LogEntry, idempotencyKey } from "./intents.js";
import { transformIntent } from "./transform.js";
import { applyIntent } from "./apply.js";

// transformIntent is used in replayPending below.

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

  /**
   * Apply a locally produced intent optimistically and enqueue it as pending.
   * The intent's `base` must be the current confirmed seq.
   *
   * CONSTRAINT (one-in-flight): a client should hold at most one un-confirmed
   * pending intent at a time — submit the next only after the previous is
   * confirmed. This is the model production OT servers use (e.g. ShareDB).
   * Multiple concurrent same-client pending whose offsets assume each other
   * cannot be correctly rebased against an interleaved remote edit without
   * operation inverses, which the OT-lite transform deliberately does not
   * implement (plan doc 03 marks adjusted-sibling replay a scoped next step).
   * The transform below is correct for the single-in-flight case; the replay
   * still stores transformed forms so a burst that stays same-client-only
   * (no remote interleave) also composes.
   */
  submitLocal(intent: Intent): void {
    applyIntent(this.doc, this.ids, intent);
    // Keep the parsed model consistent with the mutated XML: op-style
    // optimistic submits (toolbar commands routed through the canonical
    // apply) are rendered directly from this replica's live model.
    this.resync();
    this.pending.push(intent);
  }

  /** Track an intent the caller ALREADY applied to this replica's live doc
   * (the editor-driven path: DocxEditor mutates `doc` through its own command
   * and emits the intent afterwards). Applying it again here would double it —
   * the demo typed "Hello" and got "Hello" + its reversal. Pending tracking
   * and reconciliation behave exactly as for submitLocal. */
  trackLocal(intent: Intent): void {
    this.pending.push(intent);
  }

  /** Number of un-confirmed local intents (for the one-in-flight discipline). */
  get pendingCount(): number {
    return this.pending.length;
  }

  /**
   * Receive the server's canonical broadcast entries (in seq order) and
   * reconcile. The invariant is: live doc = confirmed baseline + remaining
   * pending. So every receive restores to the confirmed baseline, applies the
   * canonical batch (advancing the baseline and dropping matched pending),
   * then replays the still-pending intents transformed against the remote
   * intents in the batch (plan doc 03). Idempotent for already-seen seqs.
   */
  /** True after a receive() that reloaded the doc object (a true conflict
   * reconciliation) — the renderer must re-key/rebind. Stays false for the
   * in-place fast paths so the editor updates without a re-mount. */
  reloaded = false;

  receive(entries: LogEntry[]): void {
    this.reloaded = false;
    const fresh = entries.filter((e) => e.seq > this.confirmedSeq);
    if (fresh.length === 0) return;

    const remoteAhead = fresh
      .filter((e): e is Extract<LogEntry, { kind: "applied" }> => e.kind === "applied" && !isOurs(e, this.pending))
      .map((e) => e.intent);

    // A REJECTION of one of our own pending intents means our optimistic doc
    // contains a mutation the canonical history never will. Dropping the
    // pending entry alone (the old behavior) left that mutation in the doc
    // forever — the client diverged permanently from the server. Roll back:
    // restore the confirmed snapshot, advance through the fresh entries, and
    // replay the surviving pending intents.
    const rejectedOurs = fresh.some((e) => e.kind === "rejected" && isOurs(e, this.pending));
    if (rejectedOurs) {
      this.restoreConfirmed();
      this.reloaded = true;
      for (const e of fresh) this.advanceConfirmed(e, e.kind === "applied");
      this.snapshotConfirmed();
      this.replayPending(remoteAhead);
      this.resync(); // model consistent with the replayed XML for the renderer
      return;
    }

    // FAST PATH — no reload, doc object stays stable (production real-time
    // behavior: apply in place, repaint incrementally, no re-mount/flash):
    //  (a) every fresh entry is our OWN echo: the optimistic doc already has
    //      them — just advance confirmed and drop the matched pending;
    //  (b) nothing is pending: apply the remote entries in place (no risk of
    //      double-applying an optimistic edit).
    if (remoteAhead.length === 0 || this.pending.length === 0) {
      const applyToDoc = remoteAhead.length > 0;
      for (const e of fresh) {
        // Own echoes are already applied optimistically; only apply entries
        // that aren't our own pending, to avoid double-apply.
        this.advanceConfirmed(e, /*applyToDoc*/ applyToDoc);
      }
      // apply.ts mutates the XML tree only; the parsed model (Run.content, the
      // line-break inputs the renderer reads) is rebuilt by refresh(). The old
      // render path reloaded from bytes so it never saw the stale model, but
      // rendering the live doc object directly does — so refresh once here when
      // we mutated in place, keeping the model consistent with the XML (and
      // bumping modelVersion so the view repaints).
      if (applyToDoc) this.resync();
      this.snapshotConfirmed();
      return;
    }

    // SLOW PATH (true conflict: pending + interleaved remote) — reload the
    // confirmed baseline and replay pending transformed against the remote.
    this.restoreConfirmed();
    this.reloaded = true;
    for (const e of fresh) this.advanceConfirmed(e, true);
    this.snapshotConfirmed();
    this.replayPending(remoteAhead);
    this.resync(); // model consistent with the replayed XML for the renderer
  }

  private advanceConfirmed(e: LogEntry, applyToDoc: boolean): void {
    if (e.kind === "applied" && applyToDoc) {
      applyIntent(this.doc, this.ids, e.intent);
      if (e.intent.kind === "splitParagraph") this.resync();
    }
    // Drop a matching pending intent (ours, now confirmed).
    const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
    this.pending = this.pending.filter((p) => idempotencyKey(p) !== key);
    this.confirmedSeq = e.seq;
  }

  private replayPending(remoteAhead: Intent[]): void {
    // Each remaining pending intent is transformed through the shared function
    // against the interleaved remote intents (the same transform the server
    // applied), then re-applied on the restored baseline. The TRANSFORMED form
    // is stored back so successive remote batches compose cumulatively; its
    // idempotency key is unchanged, so the eventual canonical echo still drops
    // it. A pending that no longer applies is dropped (clean failure).
    const stillPending = this.pending;
    this.pending = [];
    for (const p of stillPending) {
      const transformed = transformIntent(p, remoteAhead);
      const applied = applyIntent(this.doc, this.ids, transformed);
      if (transformed.kind === "splitParagraph") this.resync();
      if (applied) this.pending.push(transformed);
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
