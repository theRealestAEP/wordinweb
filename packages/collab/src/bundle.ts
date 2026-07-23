import type { Intent } from "./intents.js";
import type { IdSidecar } from "./session.js";
import type { LineageHead } from "./protocol.js";
import type { CollabConnection } from "./connection.js";

/**
 * The client-side document bundle (plan doc 12 §4) — under zero custody this
 * is THE durable copy of a document anywhere: the server deletes everything
 * at session end, so each participating browser persists its own bundle and
 * any holder can re-seed the same link later.
 *
 * Contents are the replica's CONFIRMED state plus the pending queue:
 * `confirmedBytes` is a complete docx at `confirmedSeq`; `confirmedSidecar`
 * rides with it always (round-4 F10 — bytes alone cannot reproduce the id
 * table across split-created ids); `pending` is idempotency-keyed and safe
 * to replay after a crash (the server dedups by (clientId, clientSeq)).
 */
export interface DocBundle {
  docId: string;
  /** Epoch of the session this state came from (doc 12 §5): on rejoin,
   * same ⇒ seamless resume; different ⇒ someone re-seeded while away. */
  genesisId: string;
  confirmedSeq: number;
  confirmedBytes: Uint8Array;
  confirmedSidecar: IdSidecar;
  /** Local intents not yet confirmed when the bundle was written — replayed
   * verbatim on resume; idempotent by (clientId, clientSeq). */
  pending: Intent[];
  /**
   * The clientSeq watermark at write time. Resume MUST restore this before
   * submitting anything new: a fresh connection restarting from 1 would
   * reuse already-sequenced (clientId, clientSeq) keys and the server would
   * dedup its NEW edits as re-sends — self-inflicted silent edit loss (the
   * same failure shape doc 12 §7's single-tab rule guards against).
   */
  clientSeq: number;
  savedAt: number;
  /**
   * The lineage chain (doc 15 §1): every epoch this copy has passed
   * through, newest last — `[..., {this bundle's genesisId, seq, hash}]`.
   * Rides with re-seeds so returning holders can prove ancestry and
   * fast-forward instead of forking. Capped; ancestry checks are O(chain).
   */
  lineage: LineageHead[];
  /** Per-part media metadata (doc 16 §5.3): the sha (and E2EE iv/epoch)
   * of every media part this copy knows. The docx bytes carry the PIXELS;
   * this carries the ADDRESSES — without it a resumed holder couldn't
   * answer re-supply requests or verify fetches (metadata is in-memory on
   * the doc and would otherwise die with the session). */
  mediaMeta?: [string, { sha: string; iv?: string; genesisId?: string }][];
}

/**
 * Storage seam for bundles. The browser implementation is IndexedDB (the
 * only storage that fits multi-MB binary docs); tests and Node use the
 * in-memory implementation — the doc-12 test plan's "in-memory IndexedDB
 * stub", which keeps every bundle test deterministic and dependency-free.
 */
export interface BundleStore {
  get(docId: string): Promise<DocBundle | null>;
  put(bundle: DocBundle): Promise<void>;
  delete(docId: string): Promise<void>;
}

export class InMemoryBundleStore implements BundleStore {
  private bundles = new Map<string, DocBundle>();
  /** Write count — lets tests assert the throttle's coalescing precisely. */
  writes = 0;

  async get(docId: string): Promise<DocBundle | null> {
    return this.bundles.get(docId) ?? null;
  }
  async put(bundle: DocBundle): Promise<void> {
    this.writes++;
    this.bundles.set(bundle.docId, bundle);
  }
  async delete(docId: string): Promise<void> {
    this.bundles.delete(docId);
  }
}

/**
 * Persists a connection's state to a BundleStore on a THROTTLE — not a
 * debounce (round-4 F8): a debounce resets per event, so sustained typing
 * defers the write for the whole burst and an OS kill loses the burst; a
 * throttle bounds data loss to the throttle window (default 1s), which is
 * the actual durability guarantee doc 12 §4 states ("RPO ≈ your last bundle
 * write ≈ the throttle window").
 *
 * Wiring: call `notify()` from the connection's onChange (every local edit,
 * echo, and remote broadcast); call `flush()` from pagehide/visibilitychange
 * — best-effort (IndexedDB has no synchronous API; the throttle is the real
 * mechanism, the flush just narrows the tail).
 *
 * The scheduler (setTimeout/clearTimeout) and clock are injectable so tests
 * drive time deterministically.
 */
export class BundlePersister {
  private lastWrite = -Infinity;
  private trailing: unknown = null;
  private stopped = false;

  constructor(
    private conn: CollabConnection,
    private store: BundleStore,
    private docId: string,
    private opts: {
      throttleMs?: number;
      now?: () => number;
      setTimer?: (fn: () => void, ms: number) => unknown;
      clearTimer?: (t: unknown) => void;
    } = {},
  ) {}

  private get throttleMs(): number {
    return this.opts.throttleMs ?? 1000;
  }
  private now(): number {
    return (this.opts.now ?? Date.now)();
  }

  /** State changed — write now if the window allows, else arm ONE trailing
   * write for the window's end (never more than one timer in flight; N
   * notifies inside a window coalesce into a single trailing write). */
  notify(): void {
    if (this.stopped) return;
    const elapsed = this.now() - this.lastWrite;
    if (elapsed >= this.throttleMs) {
      void this.write();
      return;
    }
    if (this.trailing === null) {
      const set = this.opts.setTimer ?? ((fn, ms) => setTimeout(fn, ms));
      this.trailing = set(() => {
        this.trailing = null;
        void this.write();
      }, this.throttleMs - elapsed);
    }
  }

  /** Immediate best-effort write (pagehide/visibilitychange/session-end). */
  async flush(): Promise<void> {
    if (this.stopped) return;
    if (this.trailing !== null) {
      (this.opts.clearTimer ?? clearTimeout)(this.trailing as never);
      this.trailing = null;
    }
    await this.write();
  }

  /** Detach (unmount): cancel the trailing timer; no further writes. */
  stop(): void {
    this.stopped = true;
    if (this.trailing !== null) {
      (this.opts.clearTimer ?? clearTimeout)(this.trailing as never);
      this.trailing = null;
    }
  }

  private async write(): Promise<void> {
    const bundle = this.conn.exportBundle(this.docId);
    if (!bundle) return; // not welcomed yet — nothing durable to say
    this.lastWrite = this.now();
    bundle.savedAt = this.lastWrite;
    // Maintain the lineage chain (doc 15 §1): the PREVIOUS bundle's chain
    // carries forward; the head for the current epoch is refreshed in
    // place (same genesisId) or appended (first write in a new epoch).
    const prior = await this.store.get(this.docId);
    const chain = [...(prior?.lineage ?? [])];
    const digest = await crypto.subtle.digest("SHA-256", bundle.confirmedBytes as BufferSource);
    let hash = "";
    for (const b of new Uint8Array(digest)) hash += b.toString(16).padStart(2, "0");
    const head = { genesisId: bundle.genesisId, seq: bundle.confirmedSeq, docHash: hash };
    if (chain.length && chain[chain.length - 1].genesisId === bundle.genesisId) chain[chain.length - 1] = head;
    else chain.push(head);
    bundle.lineage = chain.slice(-50); // bounded; ancient epochs age out
    await this.store.put(bundle);
  }
}
