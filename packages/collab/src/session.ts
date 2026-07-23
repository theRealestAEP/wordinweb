import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, IntentBody, LogEntry, idempotencyKey } from "./intents.js";

/** The ID sidecar carried in a checkpoint bundle (plan doc 03). */
export type IdSidecar = ReturnType<StableIds["exportSidecar"]>;
import { transformIntent } from "./transform.js";
import { applyIntent } from "./apply.js";
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

export class DocumentSession {
  readonly doc: DocxDocument;
  readonly ids: StableIds;
  private log: LogEntry[] = [];
  private seen = new Set<string>();
  /** Base sequence number when the session was seeded from a mid-history
   * checkpoint (E2EE mirrors, doc 13 §3): entries 1..floor are baked into
   * the seed bytes; numbering continues from here. */
  private seqFloor = 0;
  /** Undo stack per client: applied intents (seq + their pre-computed inverse)
   * not yet undone. Enables selective per-user undo (plan doc 03 Phase 8). */
  private undoStacks = new Map<string, { seq: number; inverse: IntentBody }[]>();
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
        applyIntent(this.doc, this.ids, e.intent);
        if (e.intent.kind === "splitParagraph") {
          this.doc.refresh();
          this.ids.assignFromRoots(this.doc.editableRoots());
        }
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

    let applied: boolean;
    try {
      applied = applyIntent(this.doc, this.ids, canonical);
    } catch {
      applied = false;
    }
    if (!applied) return this.reject(intent, "apply failed");

    // Structural intents changed the tree shape; refresh and re-key survivors.
    if (canonical.kind === "splitParagraph") {
      this.doc.refresh();
      this.ids.assignFromRoots(this.doc.editableRoots());
    }

    const entry: LogEntry = { seq: this.seq + 1, kind: "applied", intent: canonical };
    this.log.push(entry);
    // Record the inverse on the originating client's undo stack (skip intents
    // that are themselves undo/redo — tracked via clientSeq marker below).
    if (inverse && !this.isUndoIntent(intent)) {
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
   * Selective per-user undo (plan doc 03 Phase 8): revert the given client's
   * most recent not-yet-undone intent by submitting its inverse with `base`
   * set to the ORIGINAL intent's seq — the canonical transform then rebases
   * the inverse against everything sequenced since, which is what makes undo
   * correct under concurrency. Returns the resulting log entry, or null if the
   * client has nothing undoable. The inverse is applied through the normal
   * pipeline, so it converges and broadcasts like any intent.
   */
  undo(clientId: string): LogEntry | null {
    const stack = this.undoStacks.get(clientId);
    if (!stack || stack.length === 0) return null;
    const { seq, inverse } = stack.pop()!;
    const n = (this.undoSeq.get(clientId) ?? 0) + 1;
    this.undoSeq.set(clientId, n);
    const undoIntent = {
      ...inverse,
      clientId,
      clientSeq: UNDO_CLIENT_SEQ_BASE + n,
      base: seq, // rebase the inverse through everything applied since.
    } as Intent;
    return this.submit(undoIntent);
  }

  /** Number of undoable intents for a client (for tests/UI). */
  undoDepth(clientId: string): number {
    return this.undoStacks.get(clientId)?.length ?? 0;
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
