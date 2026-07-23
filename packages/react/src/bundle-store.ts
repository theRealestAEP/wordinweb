import type { BundleStore, DocBundle } from "@wordinweb/collab/client";
/**
 * Browser BundleStore over IndexedDB — the only browser storage that fits
 * multi-MB binary documents (doc 12 §4). One object store keyed by docId;
 * bundles are stored structured-clone-ably (Uint8Array clones natively).
 * Every method opens lazily so constructing the store never touches the API
 * (SSR-safe); callers on platforms without IndexedDB should inject the
 * in-memory store instead.
 *
 * Durability honesty (doc 12 §4, round-4 F8): IndexedDB is evictable —
 * `navigator.storage.persist()` is requested once per store on first write,
 * but browsers may deny it; Safari time-boxes script-writable storage and
 * private windows drop it on close. The download-.docx escape hatch in the
 * UI, not this class, is the user's guarantee.
 */
export class IndexedDbBundleStore implements BundleStore {
  private db: Promise<IDBDatabase> | null = null;
  private persistRequested = false;

  constructor(private dbName = "wordinweb-bundles") {}

  private open(): Promise<IDBDatabase> {
    if (!this.db) {
      this.db = new Promise((resolve, reject) => {
        const req = indexedDB.open(this.dbName, 1);
        req.onupgradeneeded = () => req.result.createObjectStore("bundles", { keyPath: "docId" });
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return this.db;
  }

  private tx<T>(mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
    return this.open().then(
      (db) =>
        new Promise<T>((resolve, reject) => {
          const t = db.transaction("bundles", mode);
          const req = run(t.objectStore("bundles"));
          req.onsuccess = () => resolve(req.result);
          req.onerror = () => reject(req.error);
        }),
    );
  }

  async get(docId: string): Promise<DocBundle | null> {
    return (await this.tx("readonly", (s) => s.get(docId))) ?? null;
  }

  async put(bundle: DocBundle): Promise<void> {
    if (!this.persistRequested) {
      this.persistRequested = true;
      // Best-effort, once: ask the browser not to evict us (doc 12 §4).
      void globalThis.navigator?.storage?.persist?.();
    }
    await this.tx("readwrite", (s) => s.put(bundle));
  }

  async delete(docId: string): Promise<void> {
    await this.tx("readwrite", (s) => s.delete(docId));
  }
}
