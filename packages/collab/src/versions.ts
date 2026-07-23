import type { IdSidecar } from "./session.js";
import type { CollabConnection } from "./connection.js";

/**
 * Client-side version history (plan doc 14 §1) — a zero-custody feature:
 * the server keeps nothing, so restore points live beside the live bundle
 * in the participant's own storage.
 *
 * A version is a frozen CONFIRMED state (never optimistic — same doc 03
 * re-capture rule the bundle follows), with the sidecar riding along so a
 * restored version can re-seed a working session (F1 invariant).
 *
 * Retention (round-4 F15 — bounded, never "kept until deleted"):
 *  - autos: ring of `autoCap` (default 5, user-settable 1–20);
 *  - labeled: capped at `labeledCap` (default 20), oldest evicted WITH a
 *    warning surfaced to the caller (the UI shows it);
 *  - under browser quota pressure the POLICY is versions-before-live-bundle
 *    — implemented by the app calling `evictOldest()` on quota errors
 *    before ever touching the bundle store.
 *
 * Restore is NOT implemented here as a mutation: a version restores by
 * opening as a NEW local draft (doc 14 §1 — "back to then", never "merge
 * then into now"); the caller feeds `docx` to a fresh editor/go-live flow.
 */
export interface DocVersion {
  docId: string;
  genesisId: string;
  seq: number;
  docx: Uint8Array;
  sidecar: IdSidecar;
  savedAt: number;
  label?: string;
  auto: boolean;
}

export interface VersionStore {
  list(docId: string): Promise<DocVersion[]>;
  put(v: DocVersion): Promise<void>;
  remove(docId: string, savedAt: number): Promise<void>;
}

export class InMemoryVersionStore implements VersionStore {
  private byDoc = new Map<string, DocVersion[]>();
  async list(docId: string): Promise<DocVersion[]> {
    return [...(this.byDoc.get(docId) ?? [])].sort((a, b) => a.savedAt - b.savedAt);
  }
  async put(v: DocVersion): Promise<void> {
    const l = this.byDoc.get(v.docId) ?? [];
    l.push(v);
    this.byDoc.set(v.docId, l);
  }
  async remove(docId: string, savedAt: number): Promise<void> {
    this.byDoc.set(docId, (this.byDoc.get(docId) ?? []).filter((v) => v.savedAt !== savedAt));
  }
}

export interface VersionRingOptions {
  /** Auto-version ring size (doc 14 §1: default 5, user-set 1–20). */
  autoCap?: number;
  /** Labeled/manual cap (F15: 20, oldest evicts with a warning). */
  labeledCap?: number;
  now?: () => number;
}

export class VersionRing {
  constructor(
    private store: VersionStore,
    private opts: VersionRingOptions = {},
  ) {}
  private get autoCap(): number {
    return Math.max(1, Math.min(20, this.opts.autoCap ?? 5));
  }
  private get labeledCap(): number {
    return this.opts.labeledCap ?? 20;
  }
  private now(): number {
    return (this.opts.now ?? Date.now)();
  }

  /**
   * Capture the connection's confirmed state. Auto captures ring-evict;
   * labeled captures evict their oldest past the cap and REPORT it — the
   * return value carries any eviction so the UI can say so (silent caps
   * read as "covered everything" when they didn't).
   */
  async capture(
    conn: CollabConnection,
    docId: string,
    label?: string,
  ): Promise<{ saved: DocVersion | null; evicted: DocVersion[] }> {
    const bundle = conn.exportBundle(docId);
    if (!bundle) return { saved: null, evicted: [] };
    const v: DocVersion = {
      docId,
      genesisId: bundle.genesisId,
      seq: bundle.confirmedSeq,
      docx: bundle.confirmedBytes,
      sidecar: bundle.confirmedSidecar,
      savedAt: this.now(),
      label,
      auto: label === undefined,
    };
    await this.store.put(v);
    const evicted = await this.enforceCaps(docId);
    return { saved: v, evicted };
  }

  /** Ring/cap enforcement — oldest evicts first WITHIN each class; labeled
   * versions never count against the auto ring and vice versa. */
  private async enforceCaps(docId: string): Promise<DocVersion[]> {
    const all = await this.store.list(docId);
    const evicted: DocVersion[] = [];
    const autos = all.filter((v) => v.auto);
    for (let i = 0; i < autos.length - this.autoCap; i++) {
      await this.store.remove(docId, autos[i].savedAt);
      evicted.push(autos[i]);
    }
    const labeled = all.filter((v) => !v.auto);
    for (let i = 0; i < labeled.length - this.labeledCap; i++) {
      await this.store.remove(docId, labeled[i].savedAt);
      evicted.push(labeled[i]);
    }
    return evicted;
  }

  /** Quota-pressure relief (F15 ordering: versions die before the live
   * bundle, oldest-auto first, then oldest-labeled). Returns what was
   * dropped, or empty when there is nothing left to give. */
  async evictOldest(docId: string): Promise<DocVersion[]> {
    const all = await this.store.list(docId);
    const target = all.find((v) => v.auto) ?? all[0];
    if (!target) return [];
    await this.store.remove(docId, target.savedAt);
    return [target];
  }
}
