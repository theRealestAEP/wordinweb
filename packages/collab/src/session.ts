import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, LogEntry, idempotencyKey } from "./intents.js";

/** The ID sidecar carried in a checkpoint bundle (plan doc 03). */
export type IdSidecar = ReturnType<StableIds["exportSidecar"]>;
import { transformIntent } from "./transform.js";
import { applyIntent } from "./apply.js";

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

export class DocumentSession {
  readonly doc: DocxDocument;
  readonly ids: StableIds;
  private log: LogEntry[] = [];
  private seen = new Set<string>();

  constructor(doc: DocxDocument) {
    this.doc = doc;
    this.ids = doc.enableStableIds();
  }

  /** Newest assigned sequence number (0 before any entry). Derived from the
   * last log entry so it stays correct after a snapshot prunes the log
   * prefix (rehydration). */
  get seq(): number {
    return this.log.length === 0 ? 0 : this.log[this.log.length - 1].seq;
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

    if (intent.base < 0 || intent.base > this.seq) {
      return this.reject(intent, "invalid base");
    }
    const ahead = this.log
      .filter((e): e is Extract<LogEntry, { kind: "applied" }> => e.kind === "applied" && e.seq > intent.base)
      .map((e) => e.intent);
    const canonical = transformIntent(intent, ahead);

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
    return entry;
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
