import { parseBundleKey, type BundleStore, type DocBundle, type StoredDocSummary } from "@wordinweb/collab/client";

/** The tiny record `list()` reads instead of the bundle (see below). */
interface MetaRecord {
  docId: string;
  savedAt: number;
  byteLength: number;
}

/**
 * Browser BundleStore over IndexedDB — the only browser storage that fits
 * multi-MB binary docs (doc 12 §4). Bundles are stored structured-clone-ably
 * (Uint8Array clones natively) in `bundles`, keyed by docId. A second store,
 * `meta`, holds a per-key {savedAt, byteLength} record written in the SAME
 * transaction as each bundle: `list()` reads only key names and these tiny
 * records, so enumerating ten multi-MB documents never deserialises their
 * bytes. Bundles written by the v1 schema have no meta record; `list()`
 * reads those once and backfills, so the cost is paid one time per legacy
 * record, not per listing.
 *
 * Every method opens lazily so constructing the store never touches the API
 * (SSR-safe); callers on platforms without IndexedDB should inject the
 * in-memory store instead.
 *
 * Durability honesty (doc 12 §4, round-4 F8): IndexedDB is evictable —
 * `navigator.storage.persist()` is requested once per store on first write,
 * but browsers may deny it; Safari time-boxes script-writable storage and
 * private windows drop it on close. The download-.docx escape hatch in the
 * UI, not this class, is the user's guarantee. Failures here REJECT — a
 * swallowed quota error would let the user believe work is saved when it is
 * not, so surfacing them is the caller's contract, not an option.
 */
/**
 * Legacy records measured per `list()` call. Small on purpose: each costs a
 * full bundle read, and the cost is paid on a page load.
 */
const LEGACY_BACKFILL_PER_LIST = 4;

export class IndexedDbBundleStore implements BundleStore {
  private db: Promise<IDBDatabase> | null = null;
  private persistRequested = false;

  constructor(private dbName = "wordinweb-bundles") {}

  private open(): Promise<IDBDatabase> {
    if (!this.db) {
      this.db = new Promise((resolve, reject) => {
        const req = indexedDB.open(this.dbName, 2);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains("bundles")) db.createObjectStore("bundles", { keyPath: "docId" });
          if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta", { keyPath: "docId" });
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return this.db;
  }

  /** One request in one transaction, as a promise. */
  private tx<T>(stores: string | string[], mode: IDBTransactionMode, run: (t: IDBTransaction) => IDBRequest<T>): Promise<T> {
    return this.open().then(
      (db) =>
        new Promise<T>((resolve, reject) => {
          const t = db.transaction(stores, mode);
          const req = run(t);
          req.onsuccess = () => resolve(req.result);
          req.onerror = () => reject(req.error);
          // Quota failures can abort the transaction without a request error.
          t.onabort = () => reject(t.error ?? new Error("IndexedDB transaction aborted"));
        }),
    );
  }

  async get(docId: string): Promise<DocBundle | null> {
    return (await this.tx("bundles", "readonly", (t) => t.objectStore("bundles").get(docId))) ?? null;
  }

  async put(bundle: DocBundle): Promise<void> {
    if (!this.persistRequested) {
      this.persistRequested = true;
      // Best-effort, once: ask the browser not to evict us (doc 12 §4).
      void globalThis.navigator?.storage?.persist?.();
    }
    const db = await this.open();
    // Bundle + meta land atomically: a listing must never describe a bundle
    // that failed to write, or miss one that succeeded.
    await new Promise<void>((resolve, reject) => {
      const t = db.transaction(["bundles", "meta"], "readwrite");
      t.objectStore("bundles").put(bundle);
      const meta: MetaRecord = { docId: bundle.docId, savedAt: bundle.savedAt, byteLength: bundle.confirmedBytes.byteLength };
      t.objectStore("meta").put(meta);
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error ?? new Error("IndexedDB transaction aborted"));
    });
  }

  async delete(docId: string): Promise<void> {
    const db = await this.open();
    await new Promise<void>((resolve, reject) => {
      const t = db.transaction(["bundles", "meta"], "readwrite");
      t.objectStore("bundles").delete(docId);
      t.objectStore("meta").delete(docId);
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error ?? new Error("IndexedDB transaction aborted"));
    });
  }

  async list(): Promise<StoredDocSummary[]> {
    const keys = (await this.tx("bundles", "readonly", (t) => t.objectStore("bundles").getAllKeys())) as string[];
    const metas = (await this.tx("meta", "readonly", (t) => t.objectStore("meta").getAll())) as MetaRecord[];
    const byKey = new Map(metas.map((m) => [m.docId, m]));
    const out: StoredDocSummary[] = [];
    // BACKFILL IS BOUNDED, and that bound is the whole point.
    //
    // Records written before the meta store existed carry no metadata, and the
    // only way to recover a bundle's size is to read the bundle — megabytes,
    // deserialised, to learn two numbers. Doing that for every legacy key made
    // opening the page deserialise the entire store: on a browser holding a
    // pile of leaked versions it took over thirty seconds and a matching spike
    // in memory, before anything rendered. Precisely the browsers that most
    // need the reclaim this listing feeds.
    //
    // So a few per call. The listing stays fast, sizes converge over a handful
    // of loads, and until a record is measured it reports `byteLength: 0` —
    // read by callers as "unknown", never as "empty" (see StoredDocSummary).
    let budget = LEGACY_BACKFILL_PER_LIST;
    for (const key of keys) {
      let meta = byKey.get(key);
      if (!meta && budget > 0) {
        budget--;
        const bundle = await this.get(key);
        if (!bundle) continue; // deleted between the two reads
        meta = { docId: key, savedAt: bundle.savedAt, byteLength: bundle.confirmedBytes.byteLength };
        await this.tx("meta", "readwrite", (t) => t.objectStore("meta").put(meta));
      }
      out.push({
        ...parseBundleKey(key),
        key,
        // Unmeasured legacy record: 0 means UNKNOWN. A byte-budget caller must
        // not treat it as free, and must not delete it for being large either.
        savedAt: meta?.savedAt ?? 0,
        byteLength: meta?.byteLength ?? 0,
      });
    }
    return out;
  }
}
