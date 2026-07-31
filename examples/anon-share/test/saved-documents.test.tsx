// @vitest-environment jsdom
/**
 * Saved documents in the demo (examples/anon-share): the version strip must
 * come back after a reload (it used to render from useState([]) — saved
 * versions became invisible on refresh even though they sat in IndexedDB),
 * and the File menu's saved listing must be visible. A live demo can open a
 * saved copy only when it has a local-editor handoff, which leaves the shared
 * session instead of silently replacing the live document.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { InMemoryBundleStore, versionKey, type DocBundle, type StoredDocSummary } from "@wordinweb/collab/client";
import { App } from "../src/app";
import { FileMenu } from "../src/file-menu";

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
afterEach(unmountAll);

const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);
const allById = (host: HTMLElement, id: string) => [...host.querySelectorAll<HTMLElement>(`[data-testid="${id}"]`)];

async function click(el: HTMLElement | null) {
  expect(el).toBeTruthy();
  await act(async () => {
    el!.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    await Promise.resolve();
  });
}

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

const provider: DocProvider = { load: () => blankDocxBytes() };

/** Loopback WebSocket onto a real CollabHub (same shape the fidelity suites use). */
function hubSocketClass(hub: CollabHub) {
  let n = 0;
  return class HubSocket {
    constructor(_url: string) {
      const ls: ((ev: { data: unknown }) => void)[] = [];
      const conn: Connection = {
        id: `sd-c${n++}`,
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

function bundleAt(key: string, savedAt: number): DocBundle {
  return {
    docId: key,
    genesisId: "g1",
    confirmedSeq: 1,
    confirmedBytes: blankDocxBytes(),
    confirmedSidecar: { next: 1, entries: [] },
    pending: [],
    clientSeq: 0,
    savedAt,
    lineage: [],
  };
}

function mountApp(store: InMemoryBundleStore, docId: string, clientId: string, onDisconnect?: (bytes: Uint8Array | null) => void) {
  return render(createElement(App, {
    url: "ws://loopback/collab",
    httpBase: "http://loopback",
    docId,
    clientId,
    name: "Ada",
    store,
    onDisconnect,
  }));
}

describe("the version strip", () => {
  it("renders saved versions after a reload, from a fresh store read", async () => {
    const store = new InMemoryBundleStore();
    // Two versions saved by a PREVIOUS visit (what a reload finds on disk).
    await store.put(bundleAt(versionKey("vdoc", 1000, "first draft"), 1000));
    await store.put(bundleAt(versionKey("vdoc", 2000), 2000));
    // A different document's version must NOT bleed into this strip.
    await store.put(bundleAt(versionKey("otherdoc", 1500, "not mine"), 1500));

    const hub = new CollabHub(provider);
    const prevSocket = globalThis.WebSocket;
    (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub);
    try {
      const host = mountApp(store, "vdoc", "v-client");
      for (let i = 0; i < 100 && allById(host, "versions").length === 0; i++) await tick();
      const chips = [...(byId(host, "versions")?.querySelectorAll(".version-chip") ?? [])];
      expect(chips).toHaveLength(2);
      // Oldest first, unnamed shown as "Unnamed", never the storage sentinel.
      expect(chips[0].textContent).toContain("first draft");
      expect(chips[1].textContent).toContain("Unnamed");
      expect(byId(host, "versions")!.textContent).not.toContain("not mine");
    } finally {
      (globalThis as { WebSocket: unknown }).WebSocket = prevSocket;
    }
  });
});

describe("the saved listing in a live session", () => {
  it("lists entries with dates, refuses open, offers download, and deletes in two steps", async () => {
    const store = new InMemoryBundleStore();
    const vKey = versionKey("sdoc", 3000, "keep");
    await store.put(bundleAt(vKey, 3000));

    const hub = new CollabHub(provider);
    const prevSocket = globalThis.WebSocket;
    (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub);
    try {
      const host = mountApp(store, "sdoc", "s-client");
      const trigger = () => byId(host, "file-menu") as HTMLButtonElement | null;
      for (let i = 0; i < 120 && trigger()!.disabled; i++) await tick();
      expect(trigger()!.disabled, "session never became ready").toBe(false);

      await click(trigger());
      for (let i = 0; i < 50 && allById(host, "file-saved-entry").length === 0; i++) await tick();
      const entries = allById(host, "file-saved-entry");
      expect(entries.length, "the stored version must be listed").toBeGreaterThan(0);
      const entry = entries.find((e) => e.textContent?.includes("keep"))!;
      expect(entry, "the version's label and date must render").toBeTruthy();

      // THE GATE: open is refused in a live session (it would fork the room),
      // and the docId never renders anywhere in the menu (it is a capability).
      const openBtn = entry.querySelector<HTMLButtonElement>('[data-testid="file-saved-open"]')!;
      expect(openBtn.disabled, "open must be refused in a live session").toBe(true);
      expect(byId(host, "file-menu-panel")!.textContent).not.toContain("sdoc");
      // Download stays live — the honest way at a stored copy mid-session.
      const dlBtn = entry.querySelector<HTMLButtonElement>('[data-testid="file-saved-download"]')!;
      expect(dlBtn.disabled).toBe(false);

      // Delete is two-step: the first click only arms the confirmation.
      await click(entry.querySelector('[data-testid="file-saved-delete"]'));
      expect(await store.get(vKey), "one click must not delete").toBeTruthy();
      const confirm = byId(host, "file-saved-delete-confirm");
      expect(confirm, "the second, explicit step must be offered").toBeTruthy();
      await click(confirm);
      let gone = await store.get(vKey);
      for (let i = 0; i < 50 && gone; i++) { await tick(); gone = await store.get(vKey); }
      expect(gone, "confirming must delete the entry").toBeNull();
    } finally {
      (globalThis as { WebSocket: unknown }).WebSocket = prevSocket;
    }
  });

  it("opens the saved-document browser from the visible button and hands a saved copy to the local editor", async () => {
    const store = new InMemoryBundleStore();
    const saved = bundleAt("local:archive", 4000);
    await store.put(saved);
    const onDisconnect = vi.fn();

    const hub = new CollabHub(provider);
    const prevSocket = globalThis.WebSocket;
    (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub);
    try {
      const host = mountApp(store, "open-saved-doc", "open-saved-client", onDisconnect);
      const browse = () => byId(host, "saved-documents") as HTMLButtonElement | null;
      for (let i = 0; i < 120 && browse()!.disabled; i++) await tick();
      expect(browse()!.disabled, "session never became ready").toBe(false);

      await click(browse());
      for (let i = 0; i < 50 && allById(host, "file-saved-entry").length === 0; i++) await tick();
      const entry = allById(host, "file-saved-entry").find((e) => e.getAttribute("data-kind") === "local")!;
      expect(entry, "the saved document must be visible").toBeTruthy();
      const open = entry.querySelector<HTMLButtonElement>('[data-testid="file-saved-open"]')!;
      expect(open.disabled).toBe(false);
      await click(open);
      for (let i = 0; i < 50 && onDisconnect.mock.calls.length === 0; i++) await tick();
      expect(onDisconnect).toHaveBeenCalledTimes(1);
      expect(Array.from(onDisconnect.mock.calls[0][0] as Uint8Array)).toEqual(Array.from(saved.confirmedBytes));
    } finally {
      (globalThis as { WebSocket: unknown }).WebSocket = prevSocket;
    }
  });
});

describe("FileMenu saved section (unit)", () => {
  it("opens an entry via onOpenSaved when no session gates it", async () => {
    const onOpenSaved = vi.fn();
    const summary: StoredDocSummary = {
      key: "local:doc-1", docId: "local:doc-1", kind: "local", savedAt: 1000, byteLength: 2048,
    };
    const host = render(createElement(FileMenu, {
      listSaved: async () => [summary],
      onOpenSaved,
    }));
    await click(byId(host, "file-menu"));
    for (let i = 0; i < 50 && allById(host, "file-saved-entry").length === 0; i++) await tick();
    await click(byId(host, "file-saved-open"));
    expect(onOpenSaved).toHaveBeenCalledWith(summary);
  });

  it("surfaces a failed listing instead of showing an empty menu", async () => {
    const host = render(createElement(FileMenu, {
      listSaved: async () => { throw new DOMException("blocked", "InvalidStateError"); },
    }));
    await click(byId(host, "file-menu"));
    for (let i = 0; i < 50 && !byId(host, "file-saved-error"); i++) await tick();
    expect(byId(host, "file-saved-error"), "a failed list read must be said, not blank").toBeTruthy();
  });
});
