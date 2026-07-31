// @vitest-environment jsdom
/**
 * IndexedDbBundleStore enumeration (bundle-store.ts): the browser store's
 * `list()` must be METADATA-ONLY — a 500-page document is megabytes, and
 * listing ten of them must not deserialise all ten. The fake IndexedDB here
 * records every read against the `bundles` store, so the central claim is
 * observable: list() touches key names and `meta` records, never bundle
 * values (except the one-time backfill of a v1-era record). Write failures
 * must REJECT — a swallowed quota error tells the user work is saved when
 * it is not.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { IndexedDbBundleStore } from "../src/bundle-store.js";
import type { DocBundle } from "@wordinweb/collab/client";

/* ---------------- minimal fake IndexedDB (keyPath stores) ---------------- */

type Handler = (() => void) | null;

function makeRequest(exec: () => unknown, settled?: () => void) {
  const r = { result: undefined as unknown, error: null as unknown, onsuccess: null as Handler, onerror: null as Handler, onupgradeneeded: null as Handler };
  setTimeout(() => {
    try {
      r.result = exec();
      r.onsuccess?.();
    } catch (e) {
      r.error = e;
      r.onerror?.();
    }
    settled?.();
  }, 0);
  return r;
}

class FakeStoreLog {
  gets = 0;
  getAlls = 0;
  getAllKeys = 0;
  puts = 0;
}

class FakeDb {
  data = new Map<string, Map<string, Record<string, unknown>>>();
  log = new Map<string, FakeStoreLog>();
  /** Store names whose put throws (quota simulation). */
  failPuts = new Set<string>();

  objectStoreNames = {
    contains: (name: string) => this.data.has(name),
  };
  createObjectStore(name: string) {
    this.data.set(name, new Map());
    this.log.set(name, new FakeStoreLog());
    return {};
  }
  transaction(_stores: string | string[], _mode: string) {
    const db = this;
    let aborted = false;
    let pending = 0;
    let finished = false;
    // The transaction outcome fires only after every issued request settled
    // — real IndexedDB never completes a transaction under its requests.
    const settle = () => {
      pending--;
      if (finished || pending > 0) return;
      finished = true;
      setTimeout(() => (aborted ? tx.onabort?.() : tx.oncomplete?.()), 0);
    };
    const track = <T>(req: T): T => { pending++; return req; };
    const tx = {
      error: null as unknown,
      oncomplete: null as Handler,
      onerror: null as Handler,
      onabort: null as Handler,
      objectStore(name: string) {
        const rows = db.data.get(name)!;
        const log = db.log.get(name)!;
        return {
          get: (key: string) => { log.gets++; return track(makeRequest(() => rows.get(key), settle)); },
          getAll: () => { log.getAlls++; return track(makeRequest(() => [...rows.values()], settle)); },
          getAllKeys: () => { log.getAllKeys++; return track(makeRequest(() => [...rows.keys()], settle)); },
          put: (value: Record<string, unknown>) => {
            log.puts++;
            return track(makeRequest(() => {
              if (db.failPuts.has(name)) {
                aborted = true;
                tx.error = new DOMException("quota", "QuotaExceededError");
                throw tx.error;
              }
              rows.set(value.docId as string, value);
            }, settle));
          },
          delete: (key: string) => track(makeRequest(() => { rows.delete(key); }, settle)),
        };
      },
    };
    return tx;
  }
}

function installFakeIdb() {
  const db = new FakeDb();
  const idb = {
    open: (_name: string, _version: number) => {
      const req = { result: undefined as unknown, error: null as unknown, onsuccess: null as Handler, onerror: null as Handler, onupgradeneeded: null as Handler };
      // Like real IndexedDB: upgradeneeded fires before success on first open.
      setTimeout(() => {
        req.result = db;
        if (!db.data.has("bundles")) req.onupgradeneeded?.();
        req.onsuccess?.();
      }, 0);
      return req;
    },
  };
  (globalThis as { indexedDB?: unknown }).indexedDB = idb;
  return db;
}

/* ------------------------------------------------------------------------ */

function bundle(key: string, savedAt: number, size: number): DocBundle {
  return {
    docId: key,
    genesisId: "g1",
    confirmedSeq: 1,
    confirmedBytes: new Uint8Array(size),
    confirmedSidecar: { next: 1, entries: [] },
    pending: [],
    clientSeq: 0,
    savedAt,
    lineage: [],
  };
}

let prevIdb: unknown;
beforeEach(() => { prevIdb = (globalThis as { indexedDB?: unknown }).indexedDB; });
afterEach(() => { (globalThis as { indexedDB?: unknown }).indexedDB = prevIdb; });

describe("IndexedDbBundleStore.list", () => {
  it("lists stored entries from metadata without reading any bundle values", async () => {
    const db = installFakeIdb();
    const store = new IndexedDbBundleStore("t");
    await store.put({ ...bundle("doc1", 100, 5000), title: "Client brief" });
    await store.put(bundle("doc1#version-200-v1", 200, 4000));

    const bundles = db.log.get("bundles")!;
    const before = { gets: bundles.gets, getAlls: bundles.getAlls };
    const all = await store.list();

    expect(all).toHaveLength(2);
    expect(all.find((s) => s.key === "doc1")).toMatchObject({ kind: "live", savedAt: 100, byteLength: 5000, title: "Client brief" });
    expect(all.find((s) => s.key === "doc1#version-200-v1")).toMatchObject({ kind: "version", label: "v1", byteLength: 4000 });
    // The load-bearing claim: no bundle VALUE was read — keys only.
    expect(bundles.gets).toBe(before.gets);
    expect(bundles.getAlls).toBe(before.getAlls);
    expect(bundles.getAllKeys).toBeGreaterThan(0);
  });

  it("backfills a v1-era record (no meta) once, then lists it metadata-only", async () => {
    const db = installFakeIdb();
    const store = new IndexedDbBundleStore("t");
    await store.put(bundle("doc1", 100, 5000)); // creates the schema
    // A record written by the v1 schema: bundle present, no meta row.
    db.data.get("bundles")!.set("legacy", bundle("legacy", 50, 1234) as unknown as Record<string, unknown>);

    const bundles = db.log.get("bundles")!;
    const first = await store.list();
    expect(first.find((s) => s.key === "legacy")).toMatchObject({ kind: "live", savedAt: 50, byteLength: 1234 });
    expect(bundles.gets).toBe(1); // exactly the one legacy record

    const second = await store.list();
    expect(second).toHaveLength(2);
    expect(bundles.gets).toBe(1); // backfilled — never read again
  });

  it("rejects put() when the write fails, instead of swallowing it", async () => {
    const db = installFakeIdb();
    const store = new IndexedDbBundleStore("t");
    await store.put(bundle("doc1", 100, 10)); // creates the schema
    db.failPuts.add("bundles");
    await expect(store.put(bundle("doc1", 200, 10))).rejects.toThrow(/quota/i);
  });
});
