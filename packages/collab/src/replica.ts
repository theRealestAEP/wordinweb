import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, LogEntry, idempotencyKey } from "./intents.js";
import { transformIntent } from "./transform.js";
import { applyIntentScoped, resyncScope, unionScopes, type Scope } from "./apply.js";

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
  /** Confirmed entries NOT yet folded into confirmedBytes: own-echoes
   * deferred mid-burst (the original role) AND — since the lazy-baseline
   * change — every fast-path confirmed entry, so the O(doc) save() stays
   * off the receive hot path entirely. restoreConfirmed replays the tail
   * onto the baseline; exportBundleState folds it for persistence; the
   * fast path folds eagerly once the tail reaches FOLD_TAIL_AT (bounding
   * conflict-replay cost while amortizing saves to ~1/FOLD_TAIL_AT ops). */
  private confirmedTail: LogEntry[] = [];
  /** Fold cadence for the lazy confirmed baseline (see confirmedTail). */
  private static FOLD_TAIL_AT = 100;
  /**
   * Bumped every time THIS CLASS mutates the live doc (optimistic apply,
   * remote apply, replay, restore) — and deliberately NOT on bookkeeping
   * that leaves the doc untouched (trackLocal, own-echo confirmation).
   *
   * This is the renderer's repaint gate. The react layer used to repaint on
   * EVERY onChange, but for the editor-driven typing path (trackLocal + the
   * own-echo fast path) the doc was already mutated AND painted by the
   * editor, so each keystroke queued a redundant whole-document relayout —
   * invisible on 5 pages, catastrophic past the background-layout threshold
   * (500 pages: the layout ran async with the container inert until it
   * finished, blocking further typing for seconds per keystroke).
   */
  docVersion = 0;

  /**
   * Union of the dirty scopes behind the docVersion bumps not yet consumed by
   * a repaint (see takeRenderScope). Every path that bumps docVersion also
   * records what it disturbed: applyAndResync its intent's ApplyScope,
   * restoreConfirmed doc scope (the doc OBJECT was replaced). The renderer
   * drains this when it repaints, so a coalesced repaint covering several
   * bumps sees the union — the batching rule lives in unionScopes.
   */
  private renderScope: Scope | null = null;

  /** Drain the dirty scope accumulated since the last take (null when no
   * doc-mutating change was recorded). The caller repaints exactly this. */
  takeRenderScope(): Scope | null {
    const s = this.renderScope;
    this.renderScope = null;
    return s;
  }

  constructor(bytes: Uint8Array, sidecar?: ReturnType<StableIds["exportSidecar"]>) {
    this.doc = DocxDocument.load(bytes);
    this.ids = this.doc.enableStableIds();
    // A snapshot from a session whose history contains splits carries
    // non-sequential ids that parse-order assignment cannot reproduce
    // (round-2 F1) — when the welcome supplies the session's sidecar, install
    // it so every later intent's carried ids resolve to the same nodes here
    // as on the server. Without one (fresh doc, legacy tests) parse order is
    // exact by construction.
    if (sidecar) this.ids.importSidecar(this.doc.editableRoots(), sidecar);
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
    // Keep the parsed model consistent with the mutated XML: op-style
    // optimistic submits (toolbar commands routed through the canonical
    // apply) are rendered directly from this replica's live model.
    this.applyAndResync(intent);
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

  /** Identity of the OLDEST un-confirmed local intent — the delivery-retry
   * probe (a stuck front pending means the op or its echo was lost). */
  firstPending(): { clientId: string; clientSeq: number } | null {
    const p = this.pending[0];
    return p ? { clientId: p.clientId, clientSeq: p.clientSeq } : null;
  }

  /** Copies of every pending intent in submit order (their reconciliation-
   * transformed forms) — the rate-limit re-drive resends the WHOLE queue:
   * chained bursts (each op addressing ids a previous op allocated) strand
   * everything behind one refused op, and dedup makes resends free. */
  pendingCopies(): Intent[] {
    return [...this.pending];
  }

  /** The current pending copy of an intent by idempotency key, or null if it
   * was confirmed or dropped. This is the RECONCILIATION-TRANSFORMED form —
   * replayPending rewrites pending against every remote entry ingested since
   * submission — and therefore the only correct body to re-seal on a
   * stale-base retry: resealing the raw original resubmits untransformed
   * coordinates that the server then applies verbatim, while the optimistic
   * doc holds the transformed form — a silently divergent baseline once the
   * own-echo fast path snapshots it (the checkpoint-boundary divergence). */
  pendingIntent(clientId: string, clientSeq: number): Intent | null {
    return this.pending.find((p) => p.clientId === clientId && p.clientSeq === clientSeq) ?? null;
  }

  /**
   * The replica's durable state — exactly the fields the doc-12 §4 bundle
   * persists (the persistence layer is a thin observer over this). The bytes
   * and sidecar are the CONFIRMED baseline, never the live optimistic doc:
   * a bundle must be replayable (confirmed + pending re-submit), and
   * snapshotting optimistic state would bake unsequenced edits into what a
   * re-seed later presents as canonical (doc 03 re-capture rule).
   * `pending` is a copy — callers can't mutate reconciliation state.
   */
  exportBundleState(): {
    confirmedSeq: number;
    confirmedBytes: Uint8Array;
    confirmedSidecar: ReturnType<StableIds["exportSidecar"]>;
    pending: Intent[];
    mediaMeta: [string, { sha: string; iv?: string; genesisId?: string }][];
  } {
    // Fold any deferred confirmed entries into the baseline so the bundle's
    // confirmedBytes is COMPLETE — they are already confirmed and are not in
    // `pending`, so without this a bundle exported mid-burst would silently
    // drop them (they'd be recoverable from neither field).
    let confirmedBytes = this.confirmedBytes;
    let confirmedSidecar = this.confirmedSidecar;
    if (this.confirmedTail.length) {
      if (this.pending.length === 0) {
        // Quiescent: the live doc IS baseline+tail exactly — save IT. This
        // also preserves out-of-band state no intent stream can rebuild
        // (installed media pixels live on the doc, not in the log).
        confirmedBytes = this.doc.save();
        confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
      } else {
        const base = DocxDocument.load(this.confirmedBytes);
        const baseIds = base.enableStableIds();
        baseIds.importSidecar(base.editableRoots(), this.confirmedSidecar);
        for (const e of this.confirmedTail) {
          if (e.kind !== "applied") continue;
          // Reconcile after each one, exactly as the live paths do: a tail
          // that splits a paragraph and then edits the new one needs the
          // split's nodes in THIS doc's model before the next intent resolves
          // them, or the entry silently fails to apply and the exported
          // baseline is missing a confirmed edit.
          const res = applyIntentScoped(base, baseIds, e.intent);
          if (res.applied) resyncScope(base, baseIds, res);
        }
        // Replay re-REGISTERS media (the intent half) but cannot restore
        // PIXELS — carry every installed part over from the live doc.
        this.carryInstalledMedia(this.doc, base);
        confirmedBytes = base.save();
        confirmedSidecar = baseIds.exportSidecar(base.editableRoots());
      }
    }
    return {
      confirmedSeq: this.confirmedSeq,
      confirmedBytes,
      confirmedSidecar,
      pending: [...this.pending],
      mediaMeta: [...this.doc.mediaMeta] as [string, { sha: string; iv?: string; genesisId?: string }][],
    };
  }

  async exportBundleStateAsync(): Promise<ReturnType<ClientReplica["exportBundleState"]>> {
    let confirmedBytes = this.confirmedBytes;
    let confirmedSidecar = this.confirmedSidecar;
    if (this.confirmedTail.length) {
      // Macrotask hop before the heavy work: folding the confirmed tail
      // re-parses the confirmed docx and replays the tail (~50-150ms at 500
      // pages) synchronously. This method starts as a promise continuation
      // (the persister's serial write chain), which can drain at listener
      // boundaries inside an in-flight input dispatch — landing the cost in
      // a keystroke's handler on the persister's 1s cadence. In its own task
      // it runs between dispatches. Taken only on this branch so exports of
      // an unchanged document (blank-doc tests, idle sessions) keep their
      // exact timing.
      await new Promise((resolve) => setTimeout(resolve, 0));
      if (this.pending.length === 0) {
        confirmedBytes = await this.doc.saveAsync();
        confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
      } else {
        const base = DocxDocument.load(this.confirmedBytes);
        const baseIds = base.enableStableIds();
        baseIds.importSidecar(base.editableRoots(), this.confirmedSidecar);
        for (const e of this.confirmedTail) {
          if (e.kind !== "applied") continue;
          const res = applyIntentScoped(base, baseIds, e.intent);
          if (res.applied) resyncScope(base, baseIds, res);
        }
        this.carryInstalledMedia(this.doc, base);
        confirmedBytes = await base.saveAsync();
        confirmedSidecar = baseIds.exportSidecar(base.editableRoots());
      }
    }
    return {
      confirmedSeq: this.confirmedSeq,
      confirmedBytes,
      confirmedSidecar,
      pending: [...this.pending],
      mediaMeta: [...this.doc.mediaMeta] as [string, { sha: string; iv?: string; genesisId?: string }][],
    };
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
    // Do these fresh own-echoes confirm EVERY pending op? A fast typing burst
    // puts several ops in flight; an echo batch confirms only the oldest, so
    // the optimistic doc still holds the un-confirmed tail. The fast path below
    // snapshots the optimistic doc as the confirmed baseline — correct ONLY
    // when the doc equals the new baseline (nothing pending, or these echoes
    // drain it). If a burst leaves pending un-confirmed, snapshotting here would
    // bake that tail into the baseline and replay it AGAIN on the next remote
    // batch — the duplication/divergence real typing hits ("SABABAB" on the
    // server vs "SABABABAAA" locally). Route that to the reload path, which
    // rebuilds the baseline from the true confirmed state.
    const freshKeys = new Set(
      fresh.map((e) => (e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`)),
    );
    const drainsAllPending = this.pending.every((p) => freshKeys.has(idempotencyKey(p)));
    // FAST PATH — no reload, doc object stays stable (and out-of-band state like
    // installed media pixels, which live on the doc and can't be rebuilt from
    // intents, is preserved by snapshotting the live doc):
    //  (b) nothing pending: apply the remote entries in place; or
    //  (a) every fresh entry is our own echo AND drains all pending: the
    //      optimistic doc already equals the new confirmed baseline.
    if (this.pending.length === 0 || (remoteAhead.length === 0 && drainsAllPending)) {
      const applyToDoc = remoteAhead.length > 0;
      for (const e of fresh) {
        // Own echoes are already applied optimistically; only apply entries
        // that aren't our own pending, to avoid double-apply.
        this.advanceConfirmed(e, /*applyToDoc*/ applyToDoc);
        // LAZY BASELINE (perf B9/B10): the entry is confirmed and lives in
        // the optimistic doc, but is NOT re-serialized into confirmedBytes
        // here — snapshotConfirmed() ran doc.save() on EVERY drain, an
        // O(document) serialize per receive batch that made per-edit cost
        // grow with document size (the 6× seeding curve) and starved the
        // UI under floods. The tail records it; restoreConfirmed replays
        // it on the (rare) conflict path, exportBundleState folds it for
        // persistence, and the periodic fold below bounds both.
        if (e.kind === "applied") this.confirmedTail.push(e);
      }
      // There is deliberately no full resync() here. It used to close this
      // loop — one O(document) model reparse per incoming batch, the stall a
      // flood of remote edits produced (B10). advanceConfirmed now reconciles
      // each entry against what THAT entry touched, so the full refresh is
      // paid only by entries that actually report document scope.
      //
      // Pending is empty on both fast-path shapes, so live == baseline+tail
      // exactly — folding here (an O(doc) save) is CORRECT whenever we
      // choose to pay it; paying it every FOLD_TAIL_AT entries amortizes
      // the save to ~1% of ops instead of every drain.
      if (this.confirmedTail.length >= ClientReplica.FOLD_TAIL_AT) this.snapshotConfirmed();
      return;
    }

    // OWN ECHOES ONLY, but a typing burst left later ops un-confirmed. The
    // optimistic doc already holds every pending op and stays correct as-is —
    // so DON'T reload it (a reload swaps the doc object mid-typing, breaking the
    // editor's live binding and dropping the caret). Advance only the CONFIRMED
    // BASELINE by the echoed ops, in place: apply them to a copy of the baseline
    // and re-snapshot, leaving the un-confirmed tail solely in the optimistic
    // doc. (Snapshotting the optimistic doc here — the old bug — baked that tail
    // into the baseline and replayed it again on the next remote → duplication.)
    if (remoteAhead.length === 0) {
      // DEFER these confirmed own-echoes: this.doc already holds them (applied
      // optimistically) and MUST NOT be reloaded mid-typing (a reload swaps the
      // doc object, breaking the editor binding + caret). Record them in
      // confirmedTail — restoreConfirmed replays them onto the baseline, and the
      // next snapshot (when the burst drains) folds them into confirmedBytes.
      // Snapshotting the optimistic doc here (the old bug) baked the STILL-
      // pending tail into the baseline and replayed it again on the next remote
      // batch → the duplication/divergence real typing hit.
      for (const e of fresh) {
        if (e.kind === "applied") this.confirmedTail.push(e);
        const key = e.kind === "applied" ? idempotencyKey(e.intent) : `${e.clientId}:${e.clientSeq}`;
        this.pending = this.pending.filter((p) => idempotencyKey(p) !== key);
        this.confirmedSeq = e.seq;
      }
      return; // this.doc untouched — no reload, no re-mount, caret kept
    }

    // SLOW PATH — a true conflict (pending + interleaved remote): reload the
    // confirmed baseline, advance through the fresh entries, then replay the
    // still-pending intents transformed against the remote ones.
    this.restoreConfirmed();
    this.reloaded = true;
    for (const e of fresh) this.advanceConfirmed(e, true);
    this.snapshotConfirmed();
    this.replayPending(remoteAhead);
    this.resync(); // model consistent with the replayed XML for the renderer
  }

  /**
   * Apply one intent to the live doc and reconcile ONLY what it touched
   * (perf B9/B10). Every applied intent used to cost a full model reparse plus
   * a full id walk somewhere on this path, so a keystroke arriving into a
   * 600-paragraph document cost 600 paragraphs of work — the seeding curve and
   * the stall under an incoming flood. Text-level intents now reparse their
   * own paragraph; anything structural or unverifiable still takes the full
   * refresh (see resyncScope). Returns false when the intent didn't apply.
   */
  private applyAndResync(intent: Intent): boolean {
    const res = applyIntentScoped(this.doc, this.ids, intent);
    if (!res.applied) return false;
    resyncScope(this.doc, this.ids, res);
    this.docVersion++;
    this.renderScope = unionScopes(this.renderScope, res);
    return true;
  }

  private advanceConfirmed(e: LogEntry, applyToDoc: boolean): void {
    if (e.kind === "applied" && applyToDoc) this.applyAndResync(e.intent);
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
      if (this.applyAndResync(transformed)) this.pending.push(transformed);
    }
  }

  /** Copy installed media pixels (out-of-band state — bytes arrive via the
   * relay, not the intent log) from one doc into another whose matching
   * parts are still pending. */
  private carryInstalledMedia(from: DocxDocument, to: DocxDocument): void {
    for (const [part] of from.mediaMeta) {
      if (from.pendingMedia.has(part)) continue; // no pixels held locally
      const bytes = from.pkg.binary(part);
      if (bytes) to.installMedia(part, bytes);
    }
  }

  private restoreConfirmed(): void {
    // The doc OBJECT is replaced below — a repaint is needed even when the
    // tail replay applies nothing (applyAndResync bumps per applied entry).
    // Every retained layout/page binds parsed objects of the OLD doc, so no
    // narrow scope can describe this: the repaint must be document-wide.
    this.docVersion++;
    this.renderScope = { kind: "doc" };
    // The lazy baseline can be up to FOLD_TAIL_AT entries stale, so pixels
    // installed during the tail window exist only on the LIVE doc — capture
    // them before it is replaced, re-install after the replay.
    const pixelSource = this.doc;
    this.doc = DocxDocument.load(this.confirmedBytes);
    this.ids = this.doc.enableStableIds();
    // Reproduce the exact id table via the sidecar — parse-order alone would
    // renumber split-created carried ids and break address resolution
    // (plan round-2 F1).
    this.ids.importSidecar(this.doc.editableRoots(), this.confirmedSidecar);
    // Fold in own-echoes confirmed since the last snapshot (deferred to avoid
    // reloading during a typing burst). These are already confirmed, so they
    // belong to the baseline before pending is replayed on top. A replayed
    // split mutates the raw XML without reparsing doc.sections, so resync —
    // exactly like advanceConfirmed/replayPending — or the NEXT apply in this
    // same receive() builds its run map from the stale model and silently
    // drops a remote edit addressing the split's new run (permanent, silent
    // divergence under fast-typing+Enter bursts).
    for (const e of this.confirmedTail) {
      if (e.kind === "applied") this.applyAndResync(e.intent);
    }
    // Re-install pixels held by the replaced live doc (see pixelSource note).
    this.carryInstalledMedia(pixelSource, this.doc);
  }

  private snapshotConfirmed(): void {
    this.confirmedBytes = this.doc.save();
    this.confirmedSidecar = this.ids.exportSidecar(this.doc.editableRoots());
    // this.doc == the full confirmed baseline here (callers snapshot only when
    // nothing pending remains, or after a restore+advance), so the deferred
    // tail is now baked into confirmedBytes — clear it.
    this.confirmedTail = [];
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
