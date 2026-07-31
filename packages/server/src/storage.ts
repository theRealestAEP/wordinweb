import type { IdSidecar, LogEntry } from "@wordinweb/collab/server";

/**
 * Storage driver (plan doc 07): the narrow seam a host implements against its
 * own infrastructure. Deliberately small — append/read the intent log, and
 * save/load a snapshot. Production drivers back this with Postgres + S3 (the
 * durable default) or per-document SQLite (the ephemeral tier); the in-memory
 * driver below is for dev/tests. Media is a separate seam (not modeled here
 * yet — the current intents carry no media).
 */
export interface StorageDriver {
  /** Append sequenced log entries for a document (durably, before ack — the
   * caller relies on this having committed). */
  appendEntries(docId: string, entries: LogEntry[]): Promise<void>;
  /** Read the log tail from (exclusive) a sequence number. */
  readLog(docId: string, sinceSeq: number): Promise<LogEntry[]>;
  /** Persist a snapshot at a sequence number. A snapshot is always the full
   * checkpoint bundle `(docx, seq, sidecar)` (plan doc 12 §2, round-4 F10):
   * the id sidecar is not optional metadata — rehydrating from bytes alone
   * re-derives ids by parse order, which is wrong once history contains
   * split-created carried ids (round-2 F1) and silently mis-addresses every
   * subsequent intent. */
  saveSnapshot(docId: string, seq: number, docx: Uint8Array, sidecar: IdSidecar): Promise<void>;
  /** Load the latest snapshot bundle, or null if none exists. */
  loadSnapshot(docId: string): Promise<{ seq: number; docx: Uint8Array; sidecar: IdSidecar } | null>;
}

/** In-memory driver for dev/tests: everything lives in maps and is lost on
 * restart (the ephemeral contract, plan doc 10 — writ small). */
export class InMemoryStorage implements StorageDriver {
  private logs = new Map<string, LogEntry[]>();
  private snaps = new Map<string, { seq: number; docx: Uint8Array; sidecar: IdSidecar }>();

  async appendEntries(docId: string, entries: LogEntry[]): Promise<void> {
    const log = this.logs.get(docId) ?? [];
    log.push(...entries);
    this.logs.set(docId, log);
  }

  async readLog(docId: string, sinceSeq: number): Promise<LogEntry[]> {
    return (this.logs.get(docId) ?? []).filter((e) => e.seq > sinceSeq);
  }

  async saveSnapshot(docId: string, seq: number, docx: Uint8Array, sidecar: IdSidecar): Promise<void> {
    this.snaps.set(docId, { seq, docx, sidecar });
    // Prune log entries the snapshot now covers.
    const log = this.logs.get(docId);
    if (log) this.logs.set(docId, log.filter((e) => e.seq > seq));
  }

  async loadSnapshot(docId: string): Promise<{ seq: number; docx: Uint8Array; sidecar: IdSidecar } | null> {
    return this.snaps.get(docId) ?? null;
  }
}
