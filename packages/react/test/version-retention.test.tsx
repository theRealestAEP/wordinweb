// @vitest-environment jsdom
/**
 * Version retention in the demo (examples/anon-share/src/version-retention.ts).
 *
 * THE BUG UNDER TEST: saveVersion wrote a new stored version on every save
 * and trimmed only the in-memory strip (`.slice(-25)`) — nothing ever
 * deleted a STORED version, so every save leaked a full document copy into
 * IndexedDB until the origin quota ran out. Every assertion here therefore
 * goes to `store.list()` / `store.get()`, never to component state: the
 * store is where the bug lived.
 */
import { afterEach, describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { InMemoryBundleStore, versionKey, type BundleStore, type DocBundle } from "@wordinweb/collab/client";
import { App } from "../../../examples/anon-share/src/app";
import { pruneVersions, versionByteBudget, VERSION_COUNT_CAP } from "../../../examples/anon-share/src/version-retention";

/* ------------------------------ helpers --------------------------------- */

function bundleAt(key: string, savedAt: number, byteLength?: number): DocBundle {
  return {
    docId: key,
    genesisId: "g1",
    confirmedSeq: 1,
    confirmedBytes: byteLength ? new Uint8Array(byteLength) : blankDocxBytes(),
    confirmedSidecar: { next: 1, entries: [] },
    pending: [],
    clientSeq: 0,
    savedAt,
    lineage: [],
  };
}

const versionKeys = async (store: BundleStore, docId: string) =>
  (await store.list()).filter((s) => s.kind === "version" && s.docId === docId).map((s) => s.key);

/* --------------------------- prune (unit) ------------------------------- */

describe("pruneVersions", () => {
  it("byte budget: many small versions survive where few large ones would", async () => {
    const small = new InMemoryBundleStore();
    for (let i = 1; i <= 10; i++) await small.put(bundleAt(versionKey("d", i * 1000), i * 1000, 1_000));
    const droppedSmall = await pruneVersions(small, "d", 50_000);
    expect(droppedSmall).toHaveLength(0);
    expect(await versionKeys(small, "d")).toHaveLength(10);

    const large = new InMemoryBundleStore();
    for (let i = 1; i <= 10; i++) await large.put(bundleAt(versionKey("d", i * 1000), i * 1000, 20_000));
    const dropped = await pruneVersions(large, "d", 50_000);
    // 50 KB budget over 20 KB copies: the newest two fit, the third breaks
    // the budget — the 8 oldest are deleted FROM THE STORE.
    expect(dropped.map((s) => s.versionSavedAt)).toEqual([1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000]);
    expect(await versionKeys(large, "d")).toEqual([versionKey("d", 9000), versionKey("d", 10_000)]);
    // Deleted means deleted: the bytes are unreachable, not just unlisted.
    expect(await large.get(versionKey("d", 1000))).toBeNull();
  });

  it("count cap: bounds even when the byte budget never binds", async () => {
    const store = new InMemoryBundleStore();
    for (let i = 1; i <= VERSION_COUNT_CAP + 5; i++) await store.put(bundleAt(versionKey("d", i * 1000), i * 1000, 10));
    const dropped = await pruneVersions(store, "d", Number.MAX_SAFE_INTEGER);
    expect(dropped).toHaveLength(5);
    expect(await versionKeys(store, "d")).toHaveLength(VERSION_COUNT_CAP);
  });

  it("never deletes the live bundle, the newest version, or another doc's versions", async () => {
    const store = new InMemoryBundleStore();
    await store.put(bundleAt("d", 500, 5_000)); // the live bundle
    await store.put(bundleAt(versionKey("d", 1000), 1000, 5_000));
    await store.put(bundleAt(versionKey("d", 2000), 2000, 5_000));
    await store.put(bundleAt(versionKey("other", 100), 100, 5_000));
    const dropped = await pruneVersions(store, "d", 0); // zero budget: maximum pressure
    expect(dropped.map((s) => s.key)).toEqual([versionKey("d", 1000)]);
    expect(await store.get("d"), "the live bundle is never retention's to delete").toBeTruthy();
    expect(await store.get(versionKey("d", 2000)), "the newest version survives any budget").toBeTruthy();
    expect(await store.get(versionKey("other", 100)), "another doc's versions are out of scope").toBeTruthy();
  });
});

/* ------------------------- budget derivation ---------------------------- */

describe("versionByteBudget", () => {
  const setEstimate = (value: unknown) =>
    Object.defineProperty(globalThis.navigator, "storage", { value, configurable: true });
  afterEach(() => setEstimate(undefined));

  it("takes a quarter of the reported quota, capped at 512 MiB", async () => {
    setEstimate({ estimate: async () => ({ quota: 400_000_000, usage: 0 }) });
    expect(await versionByteBudget()).toBe(100_000_000);
    setEstimate({ estimate: async () => ({ quota: 1e12, usage: 0 }) });
    expect(await versionByteBudget()).toBe(512 * 1024 * 1024);
  });

  it("degrades to the conservative fallback when estimate is absent or throws", async () => {
    setEstimate(undefined);
    expect(await versionByteBudget()).toBe(64 * 1024 * 1024);
    setEstimate({ estimate: async () => { throw new Error("blocked"); } });
    expect(await versionByteBudget()).toBe(64 * 1024 * 1024);
  });
});

/* ------------------------ App save flow (component) ---------------------- */

let mounted: { root: Root; host: HTMLElement }[] = [];
function render(node: ReturnType<typeof createElement>) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  act(() => { root.render(node); });
  mounted.push({ root, host });
  return host;
}
afterEach(() => {
  for (const { root, host } of mounted) {
    act(() => { root.unmount(); });
    host.remove();
  }
  mounted = [];
  Object.defineProperty(globalThis.navigator, "storage", { value: undefined, configurable: true });
});

const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);
function click(el: HTMLElement | null) {
  expect(el).toBeTruthy();
  act(() => { el!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

const provider: DocProvider = { load: () => blankDocxBytes() };

function hubSocketClass(hub: CollabHub) {
  let n = 0;
  return class HubSocket {
    constructor(_url: string) {
      const ls: ((ev: { data: unknown }) => void)[] = [];
      const conn: Connection = {
        id: `vr-c${n++}`,
        send: (m: ServerMessage) => setTimeout(() => ls.forEach((l) => l({ data: JSON.stringify(m) })), 1),
      };
      let opened = false;
      return {
        send: (d: string) => setTimeout(() => { void hub.handle(conn, JSON.parse(d)); }, 1),
        close: () => {},
        addEventListener: (t: string, cb: never) => {
          if (t === "message") ls.push(cb as never);
          else if (t === "open" && !opened) { opened = true; (cb as () => void)(); }
        },
      } as unknown as HubSocket;
    }
  };
}

/** Mount App over a loopback hub and wait for a live bundle + ready session. */
async function mountReady(store: BundleStore, docId: string) {
  const hub = new CollabHub(provider);
  (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub);
  const host = render(createElement(App, {
    url: "ws://loopback/collab", httpBase: "http://loopback", docId, clientId: "vr-client", name: "Ada", store,
  }));
  const saveBtn = () => [...host.querySelectorAll("button")].find((b) => b.textContent === "Save version") ?? null;
  for (let i = 0; i < 200 && (saveBtn()?.disabled !== false || !(await store.get(docId))); i++) await tick();
  expect(saveBtn()?.disabled, "session never became ready").toBe(false);
  expect(await store.get(docId), "the persister never landed a live bundle").toBeTruthy();
  return { host, saveVersion: async () => {
    click(saveBtn());
    click(byId(host, "version-save"));
    await tick(10);
  } };
}

describe("App save-version retention", () => {
  const prevSocket = globalThis.WebSocket;
  afterEach(() => { (globalThis as { WebSocket: unknown }).WebSocket = prevSocket; });

  it("prunes the store after a save and says what was dropped", async () => {
    // Budget = quota/4 = 25 KB. The live blank bundle is ~2.5 KB, so with
    // three seeded 10 KB versions a fresh save must push retention over the
    // byte budget and delete the oldest stored versions.
    Object.defineProperty(globalThis.navigator, "storage", {
      value: { estimate: async () => ({ quota: 100_000, usage: 0 }) }, configurable: true,
    });
    const store = new InMemoryBundleStore();
    for (let i = 1; i <= 3; i++) await store.put(bundleAt(versionKey("vdoc", i * 1000, `old-${i}`), i * 1000, 10_000));

    const { host, saveVersion } = await mountReady(store, "vdoc");
    await saveVersion();

    // The STORE was pruned — oldest first — and the new version landed.
    const left = await versionKeys(store, "vdoc");
    expect(left.some((k) => k.includes("old-1")), "oldest version must be deleted from the store").toBe(false);
    expect(left.length, "newer versions and the fresh save survive").toBeGreaterThanOrEqual(2);
    expect(await store.get("vdoc"), "the live bundle survives").toBeTruthy();
    // The user is told, by name.
    const notice = byId(host, "version-retention-notice");
    expect(notice, "deleting a restore point silently is the failure mode").toBeTruthy();
    expect(notice!.textContent).toContain("old-1");
  });

  it("keeps every version when the budget is not exceeded (no gratuitous deletion)", async () => {
    Object.defineProperty(globalThis.navigator, "storage", {
      value: { estimate: async () => ({ quota: 4e9, usage: 0 }) }, configurable: true,
    });
    const store = new InMemoryBundleStore();
    for (let i = 1; i <= 3; i++) await store.put(bundleAt(versionKey("vdoc", i * 1000, `old-${i}`), i * 1000, 10_000));
    const { host, saveVersion } = await mountReady(store, "vdoc");
    await saveVersion();
    expect(await versionKeys(store, "vdoc")).toHaveLength(4);
    expect(byId(host, "version-retention-notice")).toBeNull();
  });

  it("on a quota failure, frees older versions of this doc and retries — reporting the deletion", async () => {
    Object.defineProperty(globalThis.navigator, "storage", {
      value: { estimate: async () => ({ quota: 100_000, usage: 0 }) }, configurable: true,
    });
    const inner = new InMemoryBundleStore();
    for (let i = 1; i <= 3; i++) await inner.put(bundleAt(versionKey("vdoc", i * 1000, `old-${i}`), i * 1000, 10_000));
    // Version writes fail (quota) until something is deleted — the shape of
    // a store that leaked itself full.
    let full = true;
    const store: BundleStore = {
      get: (k) => inner.get(k),
      list: () => inner.list(),
      delete: async (k) => { full = false; await inner.delete(k); },
      put: async (b) => {
        if (full && b.docId.includes("#version-")) throw new DOMException("quota", "QuotaExceededError");
        await inner.put(b);
      },
    };

    const { host, saveVersion } = await mountReady(store, "vdoc");
    await saveVersion();

    const left = await versionKeys(inner, "vdoc");
    expect(left.some((k) => k.includes("old-1")), "room must be made from the oldest versions").toBe(false);
    expect(left.some((k) => !k.includes("old-")), "the new version must land after the retry").toBe(true);
    expect(byId(host, "version-error"), "a recovered save is not an error").toBeNull();
    expect(byId(host, "version-retention-notice"), "the freed versions are reported").toBeTruthy();
  });

  it("surfaces a quota failure when there is nothing left to free", async () => {
    const inner = new InMemoryBundleStore();
    const store: BundleStore = {
      get: (k) => inner.get(k),
      list: () => inner.list(),
      delete: (k) => inner.delete(k),
      put: async (b) => {
        if (b.docId.includes("#version-")) throw new DOMException("quota", "QuotaExceededError");
        await inner.put(b);
      },
    };
    const { host, saveVersion } = await mountReady(store, "vdoc");
    await saveVersion();
    expect(byId(host, "version-error"), "a quota failure must surface, never be swallowed").toBeTruthy();
    expect(await versionKeys(inner, "vdoc")).toHaveLength(0);
  });
});
