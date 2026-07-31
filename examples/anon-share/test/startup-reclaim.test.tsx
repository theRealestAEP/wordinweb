// @vitest-environment jsdom
/**
 * Startup reclaim in the demo (examples/anon-share/src/startup-reclaim.ts).
 *
 * THE DEBT UNDER TEST: before e8c3188 every "Save version" leaked a full
 * document copy into IndexedDB — version-retention.ts stopped the leak for
 * NEW saves, but nothing reclaimed what had already accumulated, so an owner
 * could stay wedged against the quota on every load. Startup reclaim frees
 * the provably-redundant copies (superseded banks, older versions) and
 * NEVER the unique ones (live bundles, drafts, local archives).
 *
 * Every assertion about what was deleted goes to `store.list()` /
 * `store.get()`, never to component state — the store is where the truth is.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { InMemoryBundleStore, versionKey, draftKey, supersededKey, type BundleStore, type DocBundle } from "@wordinweb/collab/client";
import { App } from "../src/app";
import {
  reclaimStorage,
  storeByteBudget,
  RECLAIM_LIST_FAILED_MESSAGE,
} from "../src/startup-reclaim";
import { FileMenu } from "../src/file-menu";

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

const keysOf = async (store: BundleStore) => (await store.list()).map((s) => s.key).sort();

const setEstimate = (value: unknown) =>
  Object.defineProperty(globalThis.navigator, "storage", { value, configurable: true });

/* ---------------------------- reclaim (unit) ----------------------------- */

describe("reclaimStorage", () => {
  it("reclaims superseded first, then oldest versions, and stops once under budget", async () => {
    const store = new InMemoryBundleStore();
    await store.put(bundleAt("d", 100, 10_000)); // live
    // The superseded bank is NEWER than every version — class order must
    // still put it first (redundancy, not age, is the first sort key).
    await store.put(bundleAt(supersededKey("d", "gX"), 9000, 10_000));
    await store.put(bundleAt(versionKey("d", 1000), 1000, 10_000));
    await store.put(bundleAt(versionKey("d", 2000), 2000, 10_000));
    await store.put(bundleAt(versionKey("d", 3000), 3000, 10_000));

    const r = await reclaimStorage(store, 30_000); // total 50k → free 20k
    expect(r.reclaimed.map((s) => s.key)).toEqual([supersededKey("d", "gX"), versionKey("d", 1000)]);
    expect(r.overBudget).toBe(false);
    // Under budget after two deletions — the still-redundant v2000 is KEPT
    // (bias toward keeping: reclaim stops the moment there is room).
    expect(await keysOf(store)).toEqual(["d", versionKey("d", 2000), versionKey("d", 3000)].sort());
  });

  it("provably leaves drafts, local archives, live bundles, unknown keys and newest versions — even at zero budget", async () => {
    const store = new InMemoryBundleStore();
    const protectedKeys = [
      "d", // live bundle
      draftKey("d", "g1"), // ONLY copy of diverged offline work
      "local:autosave", // local document never gone live
      "local:doc-42", // archived local document
      "d#future-shape", // unknown kind: a future build's data
      versionKey("d", 2000), // newest version of d
      versionKey("e", 500), // newest (only) version of e
    ];
    for (const k of protectedKeys) await store.put(bundleAt(k, 1, 10_000));
    await store.put(bundleAt(supersededKey("d", "g0"), 1, 10_000));
    await store.put(bundleAt(versionKey("d", 1000), 1000, 10_000));

    const r = await reclaimStorage(store, 0); // maximum pressure
    expect(r.reclaimed.map((s) => s.key).sort()).toEqual([supersededKey("d", "g0"), versionKey("d", 1000)].sort());
    // THE assertion that matters: everything unique survived zero budget.
    expect(await keysOf(store)).toEqual([...protectedKeys].sort());
    // Deleted means deleted, kept means reachable.
    expect(await store.get(supersededKey("d", "g0"))).toBeNull();
    expect(await store.get(draftKey("d", "g1"))).toBeTruthy();
    // Over budget after reclaiming everything safe is SAID, never silently
    // fixed by deleting unique work.
    expect(r.overBudget).toBe(true);
  });

  it("reports exactly what the store lost: counts, bytes and list() agree", async () => {
    const store = new InMemoryBundleStore();
    await store.put(bundleAt("d", 1, 5_000));
    await store.put(bundleAt(supersededKey("d", "g0"), 1, 7_000));
    for (let i = 1; i <= 4; i++) await store.put(bundleAt(versionKey("d", i * 1000), i * 1000, 4_000));
    const before = await keysOf(store);

    const r = await reclaimStorage(store, 10_000); // total 28k
    const after = await keysOf(store);
    expect(after.length).toBe(before.length - r.reclaimed.length);
    expect(after).toEqual(before.filter((k) => !r.reclaimed.some((s) => s.key === k)).sort());
    expect(r.reclaimedBytes).toBe(r.reclaimed.reduce((n, s) => n + s.byteLength, 0));
    expect(r.keptCount).toBe(after.length);
    expect(r.keptBytes).toBe((await store.list()).reduce((n, s) => n + s.byteLength, 0));
    for (const s of r.reclaimed) expect(await store.get(s.key)).toBeNull();
  });

  it("deletes nothing when the store is under budget", async () => {
    const store = new InMemoryBundleStore();
    await store.put(bundleAt(supersededKey("d", "g0"), 1, 5_000));
    await store.put(bundleAt(versionKey("d", 1000), 1000, 5_000));
    const r = await reclaimStorage(store, 1_000_000);
    expect(r.reclaimed).toHaveLength(0);
    expect((await store.list()).length).toBe(2);
  });
});

/* ------------------------- budget derivation ---------------------------- */

describe("storeByteBudget", () => {
  afterEach(() => setEstimate(undefined));

  it("takes half the reported quota, capped at 1 GiB", async () => {
    setEstimate({ estimate: async () => ({ quota: 400_000_000, usage: 0 }) });
    expect(await storeByteBudget()).toBe(200_000_000);
    setEstimate({ estimate: async () => ({ quota: 1e12, usage: 0 }) });
    expect(await storeByteBudget()).toBe(1024 * 1024 * 1024);
  });

  it("degrades to the conservative fallback when estimate is absent or throws", async () => {
    setEstimate(undefined);
    expect(await storeByteBudget()).toBe(256 * 1024 * 1024);
    setEstimate({ estimate: async () => { throw new Error("blocked"); } });
    expect(await storeByteBudget()).toBe(256 * 1024 * 1024);
  });
});

/* ------------------------- component plumbing ---------------------------- */

let mounted: { root: Root; host: HTMLElement }[] = [];
function render(node: ReturnType<typeof createElement>) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  act(() => { root.render(node); });
  mounted.push({ root, host });
  return host;
}
function unmountAll() {
  for (const { root, host } of mounted) {
    act(() => { root.unmount(); });
    host.remove();
  }
  mounted = [];
}
afterEach(() => {
  unmountAll();
  setEstimate(undefined);
});

const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);
function click(el: HTMLElement | null) {
  expect(el).toBeTruthy();
  act(() => { el!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function until(cond: () => boolean, what: string) {
  for (let i = 0; i < 200 && !cond(); i++) await tick();
  expect(cond(), what).toBe(true);
}

const sessionReady = () =>
  Boolean((window as typeof window & { __ww?: { ready?: boolean } }).__ww?.ready);

const provider: DocProvider = { load: () => blankDocxBytes() };

function hubSocketClass(hub: CollabHub) {
  let n = 0;
  return class HubSocket {
    constructor(_url: string) {
      const ls: ((ev: { data: unknown }) => void)[] = [];
      const conn: Connection = {
        id: `sr-c${n++}`,
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

function mountApp(store: BundleStore, docId: string) {
  const hub = new CollabHub(provider);
  (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub);
  return render(createElement(App, {
    url: "ws://loopback/collab", httpBase: "http://loopback", docId, clientId: "sr-client", name: "Ada", store,
  }));
}

/* --------------------- App startup reclaim (component) ------------------- */

describe("App startup reclaim", () => {
  const prevSocket = globalThis.WebSocket;
  afterEach(() => { (globalThis as { WebSocket: unknown }).WebSocket = prevSocket; });

  it("reclaims an over-budget store on load and reports what was removed", async () => {
    // quota 100k → store budget 50k. 70k of redundant copies is seeded, so
    // reclaim must delete the superseded bank (30k) and then be under budget
    // — both versions survive.
    setEstimate({ estimate: async () => ({ quota: 100_000, usage: 70_000 }) });
    const store = new InMemoryBundleStore();
    await store.put(bundleAt(supersededKey("rdoc", "g0"), 500, 30_000));
    await store.put(bundleAt(versionKey("rdoc", 1000, "old-1"), 1000, 20_000));
    await store.put(bundleAt(versionKey("rdoc", 2000, "old-2"), 2000, 20_000));

    const host = mountApp(store, "rdoc");
    await until(() => byId(host, "reclaim-notice") !== null, "the reclaim must be reported on screen");

    // The STORE lost exactly the superseded bank.
    expect(await store.get(supersededKey("rdoc", "g0"))).toBeNull();
    expect(await store.get(versionKey("rdoc", 1000, "old-1")), "versions survive when superseded frees enough").toBeTruthy();
    expect(await store.get(versionKey("rdoc", 2000, "old-2"))).toBeTruthy();
    const notice = byId(host, "reclaim-notice")!;
    expect(notice.textContent).toContain("superseded copy");
    expect(notice.textContent).toContain("Freed");
    // Dismissible, per the version-retention pattern.
    click(byId(host, "reclaim-dismiss"));
    expect(byId(host, "reclaim-notice")).toBeNull();
  });

  it("says plainly when reclaiming everything safe is still not enough — and deletes no unique work", async () => {
    setEstimate({ estimate: async () => ({ quota: 100_000, usage: 90_000 }) });
    const store = new InMemoryBundleStore();
    await store.put(bundleAt(draftKey("rdoc", "g1"), 500, 40_000)); // unique work
    await store.put(bundleAt("local:doc-7", 600, 40_000)); // unique work
    await store.put(bundleAt(supersededKey("rdoc", "g0"), 700, 10_000)); // redundant

    const host = mountApp(store, "rdoc");
    await until(() => byId(host, "reclaim-notice") !== null, "over-budget-after-reclaim must surface");

    const notice = byId(host, "reclaim-notice")!;
    expect(notice.textContent).toContain("File > Saved");
    // The redundant copy went; the unique work provably did not.
    expect(await store.get(supersededKey("rdoc", "g0"))).toBeNull();
    expect(await store.get(draftKey("rdoc", "g1")), "a draft is the only copy of diverged work").toBeTruthy();
    expect(await store.get("local:doc-7"), "a local archive is the only copy of that document").toBeTruthy();
  });

  it("degrades to a clear message when the store cannot be listed", async () => {
    const inner = new InMemoryBundleStore();
    const store: BundleStore = {
      get: (k) => inner.get(k),
      put: (b) => inner.put(b),
      delete: (k) => inner.delete(k),
      list: async () => { throw new Error("storage blocked"); },
    };
    const host = mountApp(store, "rdoc");
    await until(() => byId(host, "reclaim-notice") !== null, "an unlistable store must produce words, not a crash");
    expect(byId(host, "reclaim-notice")!.textContent).toContain(RECLAIM_LIST_FAILED_MESSAGE);
    // The app is still up — the toolbar rendered.
    expect(byId(host, "toolbar")).toBeTruthy();
  });

  it("storage-full banner counts what is stored and routes to the saved-documents list", async () => {
    let rejectFirstPut!: (reason?: unknown) => void;
    const firstPutFailure = new Promise<void>((_, reject) => { rejectFirstPut = reject; });
    let reportFirstPutStarted!: () => void;
    const firstPutStarted = new Promise<void>((resolve) => { reportFirstPutStarted = resolve; });
    let reportPersistError!: () => void;
    const persistErrorReported = new Promise<void>((resolve) => { reportPersistError = resolve; });
    const consoleError = vi.spyOn(console, "error").mockImplementation((message) => {
      if (message === "[wordinweb] bundle-persist") reportPersistError();
    });
    // Live-bundle writes fail (the wedged-quota shape); versions seeded so
    // the banner has something to count. Small sizes: reclaim stays inert.
    const inner = new InMemoryBundleStore();
    for (let i = 1; i <= 3; i++) await inner.put(bundleAt(versionKey("rdoc", i * 1000, `v${i}`), i * 1000, 10_000));
    let putCalls = 0;
    const store: BundleStore = {
      get: (k) => inner.get(k),
      list: () => inner.list(),
      delete: (k) => inner.delete(k),
      put: () => {
        putCalls++;
        if (putCalls === 1) {
          reportFirstPutStarted();
          return firstPutFailure;
        }
        return Promise.reject(new DOMException("quota", "QuotaExceededError"));
      },
    };
    try {
      const host = mountApp(store, "rdoc");
      await until(sessionReady, "the session must become ready");
      // Page hide flushes the persister immediately. This tests the failure
      // path without waiting for its one-second production throttle.
      act(() => { window.dispatchEvent(new Event("pagehide")); });
      await firstPutStarted;
      await act(async () => {
        rejectFirstPut(new DOMException("quota", "QuotaExceededError"));
        await persistErrorReported;
      });
      await until(() => byId(host, "persist-banner") !== null, "a failed persist write must raise the banner");
      await until(() => byId(host, "persist-stored-summary") !== null, "the banner must say what is taking the space");
      const summary = byId(host, "persist-stored-summary")!;
      expect(summary.textContent).toContain("3 saved copies");
      expect(summary.textContent).toContain("29 KB");
      // The route to act: the button opens the File menu's saved list.
      expect(byId(host, "file-menu-panel")).toBeNull();
      click(byId(host, "persist-manage-saved"));
      await until(() => byId(host, "file-menu-panel") !== null, "Manage saved copies must open the File menu");
      await until(() => byId(host, "file-saved-entry") !== null, "the delete-capable saved list is on screen");
      unmountAll();
      await tick();
      expect(consoleError).toHaveBeenCalledWith("[wordinweb] bundle-persist", expect.any(DOMException));
    } finally {
      consoleError.mockRestore();
    }
  });
});

/* --------------------- FileMenu quota visibility ------------------------- */

describe("FileMenu storage usage line", () => {
  const entries = [
    { ...({ docId: "d", kind: "version" as const }), key: versionKey("d", 1000), savedAt: 1000, byteLength: 2 * 1024 * 1024 },
    { ...({ docId: "d", kind: "version" as const }), key: versionKey("d", 2000), savedAt: 2000, byteLength: 1024 * 1024 },
  ];

  it("shows usage against quota where estimate() reports one", async () => {
    setEstimate({ estimate: async () => ({ usage: 3 * 1024 * 1024, quota: 100 * 1024 * 1024 }) });
    const host = render(createElement(FileMenu, { listSaved: async () => entries }));
    click(byId(host, "file-menu"));
    await until(() => byId(host, "file-saved-quota") !== null, "the quota line must render");
    expect(byId(host, "file-saved-quota")!.textContent).toContain("3.0 MB used of 100.0 MB");
  });

  it("falls back to the listing's summed sizes where estimate() is absent", async () => {
    setEstimate(undefined);
    const host = render(createElement(FileMenu, { listSaved: async () => entries }));
    click(byId(host, "file-menu"));
    await until(() => byId(host, "file-saved-quota") !== null, "the fallback usage line must render");
    expect(byId(host, "file-saved-quota")!.textContent).toContain("Saved copies use 3.0 MB");
  });
});
