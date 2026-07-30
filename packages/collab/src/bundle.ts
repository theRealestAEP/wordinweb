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
  /**
   * Local intents recorded while the session was UNREACHABLE (doc 15 §2 /
   * doc 12 §5 "offline mode is first-class"): the user kept editing after
   * the socket died. On rejoin these are the offline tail the arrival
   * ladder reconciles — replayed as ordinary intents (same epoch), as
   * suggestions (diverged, small), or parked in a draft (diverged, large).
   * Distinct from `pending` (in-flight when a LIVE session was
   * interrupted): tail entries carry NO wire bookkeeping — clientId /
   * clientSeq / base are assigned at replay time by the ordinary submit
   * path, so the stored objects are the payload half of an Intent only.
   */
  offlineTail?: Intent[];
  /**
   * The genesisId of the epoch the offline tail was recorded AGAINST. This
   * is what lets a resume decide the arrival rung without guessing: same
   * epoch ⇒ the tail is morally a large pending queue and replays silently;
   * different (or absent, for a tail written by an older build) ⇒ true
   * divergence, offer suggest/draft — never silently bulldoze.
   */
  offlineTailEpoch?: string;
  /** Per-part media metadata (doc 16 §5.3): the sha (and E2EE iv/epoch)
   * of every media part this copy knows. The docx bytes carry the PIXELS;
   * this carries the ADDRESSES — without it a resumed holder couldn't
   * answer re-supply requests or verify fetches (metadata is in-memory on
   * the doc and would otherwise die with the session). */
  mediaMeta?: [string, { sha: string; iv?: string; genesisId?: string }][];
}

/**
 * THE KEY CONVENTION, in one place. Everything a browser stores rides the
 * bundle store under a key derived from the base docId:
 *
 *  - `<docId>`                          — the live bundle
 *  - `<docId>#version-<ts>[-<label>]`   — a frozen restore point
 *  - `<docId>#draft-<genesis>`          — an offline copy parked on divergence
 *  - `<docId>#superseded-<genesis>`     — a banked copy after a fast-forward
 *  - `local:…`                          — a document that never went live
 *
 * The builders and the parser below are the only code that knows these
 * shapes; UIs render from `parseBundleKey` instead of re-deriving them.
 */
export type StoredDocKind = "live" | "version" | "draft" | "superseded" | "local" | "unknown";

export interface ParsedBundleKey {
  /** The base document id (the part before any `#` suffix). */
  docId: string;
  kind: StoredDocKind;
  /** A version's human label, when the key carries one. */
  label?: string;
  /** A version's freeze time (ms), parsed from the key. */
  versionSavedAt?: number;
}

export function versionKey(docId: string, savedAt: number, label?: string): string {
  return `${docId}#version-${savedAt}${label ? `-${label}` : ""}`;
}
export function draftKey(docId: string, genesisId: string): string {
  return `${docId}#draft-${genesisId}`;
}
export function supersededKey(docId: string, genesisId: string): string {
  return `${docId}#superseded-${genesisId}`;
}

/** Parse any store key. An unrecognised `#` suffix degrades to `unknown`
 * rather than throwing — a future key shape must not break today's listing. */
export function parseBundleKey(key: string): ParsedBundleKey {
  if (key.startsWith("local:")) return { docId: key, kind: "local" };
  const hash = key.indexOf("#");
  if (hash === -1) return { docId: key, kind: "live" };
  const docId = key.slice(0, hash);
  const suffix = key.slice(hash + 1);
  const version = /^version-(\d+)(?:-(.*))?$/.exec(suffix);
  if (version) return { docId, kind: "version", versionSavedAt: Number(version[1]), label: version[2] };
  if (suffix.startsWith("draft-")) return { docId, kind: "draft" };
  if (suffix.startsWith("superseded-")) return { docId, kind: "superseded" };
  return { docId, kind: "unknown" };
}

/**
 * One stored entry, as METADATA ONLY — everything a list needs to render
 * (what it is, when, how big) without the document bytes. A 500-page doc is
 * megabytes; listing ten of them must not deserialise all ten, so `list()`
 * returns these and `get(key)` fetches bytes only when one is opened.
 */
export interface StoredDocSummary extends ParsedBundleKey {
  /** The full store key — what `get()`/`delete()` take. */
  key: string;
  /** When this entry was written (bundle.savedAt, ms). */
  savedAt: number;
  /** Size of the stored document bytes. */
  byteLength: number;
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
  /** Enumerate everything stored, metadata only (see StoredDocSummary). */
  list(): Promise<StoredDocSummary[]>;
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
  async list(): Promise<StoredDocSummary[]> {
    return [...this.bundles.values()].map((b) => ({
      ...parseBundleKey(b.docId),
      key: b.docId,
      savedAt: b.savedAt,
      byteLength: b.confirmedBytes.byteLength,
    }));
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
  /** Writes are async (digest + store I/O) and must not overlap on one
   * store; a serial chain guarantees ordering and lets flush() await
   * every in-flight write (round-4 F8: the flush is only meaningful if it
   * durably lands the latest state before pagehide). */
  private chain: Promise<void> = Promise.resolve();

  constructor(
    private conn: CollabConnection,
    private store: BundleStore,
    private docId: string,
    private opts: {
      throttleMs?: number;
      now?: () => number;
      setTimer?: (fn: () => void, ms: number) => unknown;
      clearTimer?: (t: unknown) => void;
      /**
       * A persistence write FAILED. Wire this — in a zero-custody design the
       * browser's stored bundle IS the durable copy, so a swallowed quota or
       * blocked-storage error means the user's only durable copy silently
       * stops updating while the editor looks perfectly healthy. That is the
       * most dangerous silence in this codebase, which is why it is now a
       * callback instead of an empty catch.
       */
      onError?: (err: unknown) => void;
      /**
       * The current offline tail (doc 15 §2), read at write time. The
       * connection's exportBundle cannot know it (the tail is recorded by
       * the layer ABOVE the connection, across connection rebuilds), so the
       * owner of the tail injects a getter. Returning an empty tail (or
       * omitting the getter) ERASES any stored tail — which is exactly
       * right once a replay has drained it.
       */
      offlineTail?: () => { tail: Intent[]; epoch?: string } | null;
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
      this.lastWrite = this.now(); // claim the slot SYNCHRONOUSLY so a
      this.enqueueWrite();         // second notify in the same tick arms
      return;                      // the trailing timer, not a 2nd write
    }
    if (this.trailing === null) {
      const set = this.opts.setTimer ?? ((fn, ms) => setTimeout(fn, ms));
      this.trailing = set(() => {
        this.trailing = null;
        this.lastWrite = this.now();
        this.enqueueWrite();
      }, this.throttleMs - elapsed);
    }
  }

  /** Immediate best-effort write (pagehide/visibilitychange/session-end).
   * Awaits ALL in-flight writes plus this one so the latest state is
   * durably landed before the tab goes away. */
  async flush(): Promise<void> {
    if (this.stopped) return;
    if (this.trailing !== null) {
      (this.opts.clearTimer ?? clearTimeout)(this.trailing as never);
      this.trailing = null;
    }
    this.lastWrite = this.now();
    this.enqueueWrite();
    await this.chain;
  }

  /** Append a write to the serial chain (never overlapping). */
  private enqueueWrite(): void {
    this.chain = this.chain
      .then(() => this.write())
      .catch((err: unknown) => {
        // The chain must not wedge — a failed write cannot block every later
        // one — but it must not be SILENT either (see onError above). The
        // next throttled write retries naturally; what the caller needs is to
        // know that "saved" is currently a lie.
        this.opts.onError?.(err);
      });
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
    // Macrotask hop before the heavy sync half. exportBundleAsync's export
    // re-parses the confirmed docx and replays the tail (~50-150ms on a
    // 500-page doc) BEFORE its first internal await, and this method starts
    // as a promise continuation (the serial chain / a prior write's store
    // completion) — which drains at listener boundaries inside an in-flight
    // input dispatch, so the cost landed in a keystroke's measured handler
    // (observed as ~70-85ms spikes on the persister's 1s cadence). In its
    // own task it runs between dispatches; the chain still serializes.
    await new Promise((resolve) => setTimeout(resolve, 0));
    const bundle = await this.conn.exportBundleAsync(this.docId);
    if (!bundle) return; // not welcomed yet — nothing durable to say
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
    // Offline tail (doc 15 §2): rides every write so a refresh mid-offline
    // (or mid-replay) keeps the un-replayed remainder. An empty tail clears
    // the stored one — a drained replay must not be re-offered.
    const off = this.opts.offlineTail?.();
    if (off && off.tail.length) {
      bundle.offlineTail = off.tail;
      bundle.offlineTailEpoch = off.epoch;
    }
    await this.store.put(bundle);
  }
}
