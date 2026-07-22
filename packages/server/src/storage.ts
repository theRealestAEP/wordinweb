import type { LogEntry } from "@wordinweb/collab/server";

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
  /** Persist a snapshot (docx bytes) at a sequence number. */
  saveSnapshot(docId: string, seq: number, docx: Uint8Array): Promise<void>;
  /** Load the latest snapshot, or null if none exists. */
  loadSnapshot(docId: string): Promise<{ seq: number; docx: Uint8Array } | null>;
}

/** In-memory driver for dev/tests: everything lives in maps and is lost on
 * restart (the ephemeral contract, plan doc 10 — writ small). */
export class InMemoryStorage implements StorageDriver {
  private logs = new Map<string, LogEntry[]>();
  private snaps = new Map<string, { seq: number; docx: Uint8Array }>();

  async appendEntries(docId: string, entries: LogEntry[]): Promise<void> {
    const log = this.logs.get(docId) ?? [];
    log.push(...entries);
    this.logs.set(docId, log);
  }

  async readLog(docId: string, sinceSeq: number): Promise<LogEntry[]> {
    return (this.logs.get(docId) ?? []).filter((e) => e.seq > sinceSeq);
  }

  async saveSnapshot(docId: string, seq: number, docx: Uint8Array): Promise<void> {
    this.snaps.set(docId, { seq, docx });
    // Prune log entries the snapshot now covers.
    const log = this.logs.get(docId);
    if (log) this.logs.set(docId, log.filter((e) => e.seq > seq));
  }

  async loadSnapshot(docId: string): Promise<{ seq: number; docx: Uint8Array } | null> {
    return this.snaps.get(docId) ?? null;
  }
}
