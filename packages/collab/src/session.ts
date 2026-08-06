import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, IntentBody, LogEntry, idempotencyKey } from "./intents.js";

/** The ID sidecar carried in a checkpoint bundle (plan doc 03). */
export type IdSidecar = ReturnType<StableIds["exportSidecar"]>;
import { transformIntent } from "./transform.js";
import { applyIntentScoped, resyncScope } from "./apply.js";
import { invertIntent } from "./invert.js";
import { validateIntent } from "./validate.js";

/**
 * The authoritative document session (plan doc 06), transport-free. It owns
 * the live document, the stable-id table, and the intent log; it sequences
 * each submitted intent through the canonical pipeline and emits broadcasts.
 * No sockets, no storage, no timers — a pure state machine so the whole
 * protocol is unit-testable with simulated clients (doc 07/09 H4).
 *
 * Pipeline per submit (doc 06): transform against (base, seq] → apply on the
 * live tree → append a sequenced log entry (applied or rejected) → return the
 * broadcast (the canonical, transformed form) for the transport to fan out.
 */
export interface Broadcast {
  entries: LogEntry[];
}

/** Undo-issued intents use clientSeq at or above this so they aren't
 * re-pushed onto the undo stack (avoids undo-of-undo growth). */
const UNDO_CLIENT_SEQ_BASE = 1_000_000_000;

/**
 * What the top of a client's undo stack offers. The three cases are
 * deliberately distinct because the UI says something different for each:
 * nothing to undo (button disabled), the last action can't be reversed yet
 * (button disabled WITH a reason), or here is the inverse to submit.
 */
/**
 * What a collaborative undo attempt did, as the UI needs to distinguish it:
 *
 *  undone         the inverse was submitted and painted optimistically
 *  cannot-undo    the last action has no inverse yet — a HARD STOP; undo does
 *                 not reach past it (the cure is more inverses, not skipping)
 *  changed-since  the target is already gone in canonical state, so the undo
 *                 could never land. The entry IS consumed: what the user
 *                 asked for has effectively happened, so their next press
 *                 moving to the previous action cannot surprise them
 *  nothing-to-undo  the stack is empty
 *  unavailable    this connection has no collaborative undo (plaintext rooms,
 *                 whose authority is the server)
 *  queued         the press is HELD because this client's own edits are still
 *                 in flight and the undo stack cannot see them yet. It runs
 *                 when they confirm, or is dropped at the deadline — either
 *                 way onUndoQueue reports it (see
 *                 EncryptedCollabConnection.undoLast)
 */
export type UndoOutcome = "undone" | "cannot-undo" | "changed-since" | "nothing-to-undo" | "unavailable" | "queued";

export type UndoCandidate =
  | { kind: "undoable"; seq: number; inverse: IntentBody }
  | { kind: "cannot-undo" }
  | { kind: "nothing-to-undo" };

export class DocumentSession {
  readonly doc: DocxDocument;
  readonly ids: StableIds;
  private log: LogEntry[] = [];
  private seen = new Set<string>();
  /** Base sequence number when the session was seeded from a mid-history
   * checkpoint (E2EE mirrors, doc 13 §3): entries 1..floor are baked into
   * the seed bytes; numbering continues from here. */
  private seqFloor = 0;
  /**
   * Undo stack per client: EVERY applied intent that client issued, newest
   * last, with its pre-computed inverse — or `null` when the kind has no
   * inverse yet.
   *
   * The null MARKERS are load-bearing, not bookkeeping. Stacking only the
   * invertible intents makes undo mean "your last action THAT HAPPENS TO BE
   * INVERTIBLE", which silently reaches past the others: type "hello", bold
   * it, press undo, and the bold stays while the typing disappears (measured,
   * not theorised — the bold applies and the depth never moves). Recording
   * every action lets undo mean "your last action" and answer honestly when
   * that action cannot be reversed yet.
   */
  private undoStacks = new Map<string, { seq: number; inverse: IntentBody | null }[]>();
  private undoSeq = new Map<string, number>();

  constructor(doc: DocxDocument) {
    this.doc = doc;
    this.ids = doc.enableStableIds();
  }

  /** Newest assigned sequence number (0 before any entry). Derived from the
   * last log entry so it stays correct after a snapshot prunes the log
   * prefix (rehydration). */
  get seq(): number {
    return this.log.length === 0 ? this.seqFloor : this.log[this.log.length - 1].seq;
  }

  /** Declare that this session's document already reflects seqs 1..n (it
   * was seeded from a checkpoint at n) — numbering continues from there. */
  setSeqFloor(n: number): void {
    this.seqFloor = n;
  }

  entriesSince(base: number): LogEntry[] {
    // Entries strictly after `base` by sequence number (log prefix may be
    // pruned past a snapshot, so index math would be wrong).
    return this.log.filter((e) => e.seq > base);
  }

  /**
   * A checkpoint bundle (plan doc 03): the snapshot docx bytes at the current
   * seq plus the ID sidecar, so a client bootstrapping from it reproduces the
   * exact id table (parse order alone cannot, because split-created nodes
   * carry non-sequential ids).
   */
  checkpoint(): { seq: number; docx: Uint8Array; sidecar: IdSidecar } {
    return {
      seq: this.seq,
      docx: this.doc.save(),
      sidecar: this.ids.exportSidecar(this.doc.editableRoots()),
    };
  }

  /** Install an ID sidecar (from a checkpoint bundle) onto the current
   * document — used when rehydrating from a snapshot so subsequent tail
   * entries resolve their carried ids correctly. */
  installSidecar(sidecar: IdSidecar): void {
    this.ids.importSidecar(this.doc.editableRoots(), sidecar);
  }

  /**
   * Rehydrate from a persisted log tail whose entries are already sequenced
   * and canonical (their positions were transformed when first submitted).
   * Applies them in order WITHOUT re-transforming, and restores the log,
   * dedup set, and seq. The document passed to the constructor must be the
   * snapshot the tail continues from.
   *
   * NOTE (plan round-2 F1): this reconstructs stable ids by the snapshot's
   * parse order plus the carried ids in split entries. That is exact for
   * histories without pre-snapshot splits; a snapshot taken after splits
   * needs the ID sidecar (documented next step) to reproduce the id table.
   */
  loadCanonical(tail: LogEntry[]): void {
    for (const e of tail) {
      if (e.kind === "applied") {
        const res = applyIntentScoped(this.doc, this.ids, e.intent);
        if (res.applied && res.kind === "split") resyncScope(this.doc, this.ids, res);
        this.seen.add(idempotencyKey(e.intent));
      } else {
        this.seen.add(`${e.clientId}:${e.clientSeq}`);
      }
      this.log.push(e);
    }
  }

  /**
   * Submit an intent from a client. Deduplicates re-sends by idempotency key,
   * transforms against everything sequenced since the intent's base, applies
   * (or rejects), logs, and returns the sequenced entry(ies) to broadcast.
   */
  submit(intent: Intent): LogEntry {
    const key = idempotencyKey(intent);
    if (this.seen.has(key)) {
      // Already sequenced: return its existing entry (idempotent re-send).
      const prior = this.log.find(
        (e) =>
          (e.kind === "applied" && idempotencyKey(e.intent) === key) ||
          (e.kind === "rejected" && `${e.clientId}:${e.clientSeq}` === key),
      );
      if (prior) return prior;
    }
    this.seen.add(key);

    // Cheap structural validation before any work (doc 06 / doc 11 F9): reject
    // malformed/oversized intents on the hot path, before transform/apply.
    const invalid = validateIntent(intent);
    if (invalid) return this.reject(intent, invalid);

    if (intent.base < 0 || intent.base > this.seq) {
      return this.reject(intent, "invalid base");
    }
    // Transform against CONCURRENT intents only: entries after the client's
    // base authored by OTHER clients. The client's own later-sequenced intents
    // are not concurrent — it authored this intent on a doc that already
    // contained them (client-FIFO order), so transforming against them would
    // double-count every one of its own edits. This is what broke burst
    // typing over a real socket: keystroke N carried base = last echo, the
    // server shifted it by the client's own N-1 in-flight inserts, and
    // offsets drifted until every subsequent apply failed.
    const ahead = this.log
      .filter(
        (e): e is Extract<LogEntry, { kind: "applied" }> =>
          e.kind === "applied" && e.seq > intent.base && e.intent.clientId !== intent.clientId,
      )
      .map((e) => e.intent);
    const canonical = transformIntent(intent, ahead);

    // Compute the inverse BEFORE applying, while the pre-state is intact
    // (e.g. a delete's removed text is still present). Null when not undoable.
    const inverse = invertIntent(this.doc, this.ids, canonical);

    let result: ReturnType<typeof applyIntentScoped>;
    try {
      result = applyIntentScoped(this.doc, this.ids, canonical);
    } catch {
      result = { kind: "doc", applied: false };
    }
    if (!result.applied) return this.reject(intent, "apply failed");

    // Structural intents changed the tree shape; rebuild the model for what
    // they touched and re-key survivors. A split is the only intent whose new
    // nodes the session must see in the MODEL before the next apply resolves
    // them (everything else either refreshes inside the mutation or addresses
    // nodes the model already holds) — so it is the only scope acted on here,
    // and it now reparses two paragraphs instead of the document (perf B9).
    if (result.kind === "split") resyncScope(this.doc, this.ids, result);

    const entry: LogEntry = { seq: this.seq + 1, kind: "applied", intent: canonical };
    this.log.push(entry);
    // Record the action on the originating client's undo stack — WITH a null
    // inverse when the kind isn't invertible yet, so undo can stop there
    // instead of reaching past it (see undoStacks). Intents that are
    // themselves undos are excluded via the reserved clientSeq range.
    if (!this.isUndoIntent(intent)) {
      const stack = this.undoStacks.get(intent.clientId) ?? [];
      stack.push({ seq: entry.seq, inverse });
      this.undoStacks.set(intent.clientId, stack);
    }
    return entry;
  }

  private isUndoIntent(intent: Intent): boolean {
    // Undo-issued intents use clientSeq in a reserved high range so they don't
    // themselves get pushed onto the undo stack (no undo-of-undo here).
    return intent.clientSeq >= UNDO_CLIENT_SEQ_BASE;
  }

  /**
   * Take this client's most recent action off its undo stack WITHOUT applying
   * anything — the caller decides how to submit the inverse.
   *
   * This exists because `undo()` below both pops AND submits, and an
   * ENCRYPTED client must never do the second half here: its mirror is a
   * local re-derivation of the canonical log and may only advance by
   * ingesting sequenced envelopes. Applying an inverse directly to the mirror
   * would desynchronise it from every other client's mirror — the one thing
   * the mirror architecture exists to prevent. So the encrypted path takes
   * the inverse, seals it, and submits it as an ordinary intent; it arrives
   * back through ingest like anything else.
   *
   * `cannot-undo` does NOT pop: undo stops at an action it cannot reverse
   * rather than skipping to an older one. That is a deliberate dead end — the
   * cure is implementing more inverses, not quietly undoing something the
   * user did not ask about.
   */
  takeUndo(clientId: string): UndoCandidate {
    const stack = this.undoStacks.get(clientId);
    if (!stack || stack.length === 0) return { kind: "nothing-to-undo" };
    const top = stack[stack.length - 1];
    if (!top.inverse) return { kind: "cannot-undo" };
    stack.pop();
    return { kind: "undoable", seq: top.seq, inverse: top.inverse };
  }

  /** The clientSeq an undo intent must carry so it is not itself stacked.
   * Exposed because the encrypted client builds its own undo intents. */
  nextUndoClientSeq(clientId: string): number {
    const n = (this.undoSeq.get(clientId) ?? 0) + 1;
    this.undoSeq.set(clientId, n);
    return UNDO_CLIENT_SEQ_BASE + n;
  }

  /**
   * Selective per-user undo (plan doc 03 Phase 8): revert the given client's
   * most recent action by submitting its inverse with `base` set to the
   * ORIGINAL intent's seq — the canonical transform then rebases the inverse
   * against everything sequenced since, which is what makes undo correct under
   * concurrency. Returns the resulting log entry, or null when there is
   * nothing to undo OR the last action has no inverse yet (use `takeUndo` when
   * the difference matters to the caller). The inverse is applied through the
   * normal pipeline, so it converges and broadcasts like any intent.
   *
   * PLAINTEXT path only — see takeUndo for why an encrypted client cannot use
   * this method.
   */
  undo(clientId: string): LogEntry | null {
    const candidate = this.takeUndo(clientId);
    if (candidate.kind !== "undoable") return null;
    const undoIntent = {
      ...candidate.inverse,
      clientId,
      clientSeq: this.nextUndoClientSeq(clientId),
      base: candidate.seq, // rebase the inverse through everything applied since.
    } as Intent;
    return this.submit(undoIntent);
  }

  /** How many of this client's actions are on its undo stack — including ones
   * whose inverse doesn't exist yet, because they are still actions the user
   * took. Ask `takeUndo` (or peek via `undoState`) for reversibility. */
  undoDepth(clientId: string): number {
    return this.undoStacks.get(clientId)?.length ?? 0;
  }

  /**
   * Do the nodes an inverse addresses still exist in CANONICAL state?
   *
   * A cheap NECESSARY condition, deliberately not a full validation: every
   * inverse this session produces is id-addressed, so a retired id means the
   * target is definitively gone (someone else's delete already sequenced) and
   * the undo cannot possibly land. An encrypted client checks this against
   * its mirror before painting optimistically — declining up front is both
   * honest and flash-free, where submitting would paint the undo and then
   * visibly revert it on the rejection.
   *
   * It does NOT prove the inverse WILL apply: a target that still exists may
   * have changed in ways only the real apply can reject (a text range that no
   * longer spans what it did). Those fall through to the ordinary optimistic
   * rollback every other op already has. Sound about "no", silent about "yes"
   * — which is the right asymmetry for a pre-flight check, since a false
   * "can't" would block a legitimate undo while a false "can" only costs the
   * flash we already tolerate.
   */
  targetsResolve(body: IntentBody): boolean {
    const ids: unknown[] = [];
    const b = body as unknown as Record<string, unknown>;
    for (const key of ["blockId", "runId", "cellParagraphId"]) {
      if (b[key] !== undefined) ids.push(b[key]);
    }
    const at = b.at as { blockId?: unknown; runId?: unknown } | undefined;
    if (at) {
      if (at.blockId !== undefined) ids.push(at.blockId);
      if (at.runId !== undefined) ids.push(at.runId);
    }
    // An inverse addressing nothing (none exist today) is not provably dead.
    return ids.every((id) => typeof id === "number" && this.ids.elOf(id) !== undefined);
  }

  /** What the next undo WOULD do, without consuming anything — for enabling
   * a toolbar button and choosing its tooltip. */
  undoState(clientId: string): UndoCandidate["kind"] {
    const stack = this.undoStacks.get(clientId);
    if (!stack || stack.length === 0) return "nothing-to-undo";
    return stack[stack.length - 1].inverse ? "undoable" : "cannot-undo";
  }

  /**
   * E2EE mode (doc 13 §2): an envelope that fails to open (garbage from a
   * malicious participant, or any tamper) still consumed a sequence number
   * — every honest client must agree that seq is a no-op, deterministically,
   * because the applied/rejected verdict feeds the transform's `ahead` set.
   * GCM authentication IS deterministic (same bytes fail for everyone), so
   * ingesting the failure as a sequenced rejection keeps all mirrors
   * byte-identical. Client-side use only; the blind server never calls this.
   */
  ingestOpaqueFailure(clientId: string, clientSeq: number): LogEntry {
    this.seen.add(`${clientId}:${clientSeq}`);
    return this.reject({ clientId, clientSeq } as Intent, "undecryptable");
  }

  private reject(intent: Intent, reason: string): LogEntry {
    const entry: LogEntry = {
      seq: this.seq + 1,
      kind: "rejected",
      clientId: intent.clientId,
      clientSeq: intent.clientSeq,
      reason,
    };
    this.log.push(entry);
    return entry;
  }
}
