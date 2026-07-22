import { DocxDocument, StableIds } from "@wordinweb/core";
import { Intent, LogEntry, idempotencyKey } from "./intents.js";
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

  /** Newest assigned sequence number (0 before any entry). */
  get seq(): number {
    return this.log.length;
  }

  entriesSince(base: number): LogEntry[] {
    return this.log.slice(base);
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

    if (intent.base < 0 || intent.base > this.log.length) {
      return this.reject(intent, "invalid base");
    }
    const ahead = this.log
      .slice(intent.base)
      .filter((e): e is Extract<LogEntry, { kind: "applied" }> => e.kind === "applied")
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

    const entry: LogEntry = { seq: this.log.length + 1, kind: "applied", intent: canonical };
    this.log.push(entry);
    return entry;
  }

  private reject(intent: Intent, reason: string): LogEntry {
    const entry: LogEntry = {
      seq: this.log.length + 1,
      kind: "rejected",
      clientId: intent.clientId,
      clientSeq: intent.clientSeq,
      reason,
    };
    this.log.push(entry);
    return entry;
  }
}
