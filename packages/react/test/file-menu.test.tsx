// @vitest-environment jsdom
/**
 * The demo's File menu (examples/anon-share/src/file-menu.tsx), and the one
 * rule about it that is not cosmetic: OPEN IS NEVER LIVE IN A SESSION.
 *
 * Loading a file replaces the whole document locally and emits no intent
 * describing the replacement, so a collaborative tab that did it would fork
 * from every peer while both sides still believed they shared a document —
 * the failure class fixed in 1087298, where three insert APIs mutated locally
 * under collab without emitting. The gate is a prop (`openBlockedReason`), so
 * these tests check both halves: the component honours it, and the
 * collaborative screen actually passes it.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { FileMenu } from "../../../examples/anon-share/src/file-menu";
import { App } from "../../../examples/anon-share/src/app";

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
});

const byId = (host: HTMLElement, id: string) =>
  host.querySelector<HTMLElement>(`[data-testid="${id}"]`);

function click(el: HTMLElement | null) {
  expect(el).toBeTruthy();
  act(() => { el!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}

/**
 * Hand the hidden input a file the way the OS picker would, then fire change.
 *
 * jsdom never populates `input.value` from a pick, so a plain
 * `expect(input.value).toBe("")` is VACUOUSLY true and proves nothing about
 * the reset. The accessor below stands in for the browser's populated value,
 * so the assertion has something real to observe.
 */
async function pickFile(host: HTMLElement, name: string, bytes: Uint8Array) {
  const input = host.querySelector<HTMLInputElement>('.filemenu input[type="file"]');
  expect(input, "the file picker must exist when Open is live").toBeTruthy();
  const file = new File([bytes as BlobPart], name, {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  Object.defineProperty(input!, "files", { configurable: true, value: [file] });
  let value = `C:\\fakepath\\${name}`;
  Object.defineProperty(input!, "value", {
    configurable: true,
    get: () => value,
    set: (v: string) => { value = v; },
  });
  await act(async () => {
    input!.dispatchEvent(new Event("change", { bubbles: true }));
    // The handler awaits file.arrayBuffer(); let those microtasks drain.
    await new Promise<void>((r) => setTimeout(r, 0));
  });
  return input!;
}

describe("FileMenu", () => {
  it("picking a file calls onOpen with that file's exact bytes and name", async () => {
    const onOpen = vi.fn();
    const host = render(createElement(FileMenu, { onOpen }));

    // Closed to start: the panel is not in the DOM at all.
    expect(byId(host, "file-menu-panel")).toBeNull();
    expect(byId(host, "file-menu")!.getAttribute("aria-expanded")).toBe("false");

    click(byId(host, "file-menu"));
    expect(byId(host, "file-menu-panel")).toBeTruthy();

    const open = byId(host, "file-open") as HTMLButtonElement;
    expect(open.disabled).toBe(false);
    // Open delegates to the hidden input; in jsdom nothing opens a picker, so
    // watch for the delegation and then deliver the file ourselves.
    const input = host.querySelector<HTMLInputElement>('.filemenu input[type="file"]')!;
    const opened = vi.fn();
    input.click = opened;
    click(open);
    expect(opened).toHaveBeenCalledTimes(1);
    // Activating an item closes the menu.
    expect(byId(host, "file-menu-panel")).toBeNull();

    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 7, 9, 11]);
    const after = await pickFile(host, "quarterly.docx", bytes);

    expect(onOpen).toHaveBeenCalledTimes(1);
    const [gotBytes, gotName] = onOpen.mock.calls[0];
    expect(gotName).toBe("quarterly.docx");
    expect(gotBytes).toBeInstanceOf(Uint8Array);
    expect(Array.from(gotBytes as Uint8Array)).toEqual(Array.from(bytes));
    // Reset, so re-picking the SAME file fires change again.
    expect(after.value, "the picker must be cleared after each pick").toBe("");
  });

  it("openDisabled greys Open out and refuses to open a picker", async () => {
    const onOpen = vi.fn();
    // BOTH props on purpose: a screen that contradicts itself must get the
    // disabled item, never a live Open.
    const host = render(createElement(FileMenu, { onOpen, openDisabled: true }));

    click(byId(host, "file-menu"));
    const open = byId(host, "file-open") as HTMLButtonElement;
    // Present, not hidden — the menu holds still between the two screens.
    expect(open).toBeTruthy();
    expect(open.disabled).toBe(true);
    // No picker is rendered at all while Open is blocked.
    expect(host.querySelector('input[type="file"]')).toBeNull();

    click(open);
    expect(onOpen).not.toHaveBeenCalled();
  });
});

/* ---- Enough of a browser for the collaborative screen to reach `ready`. ---- */

const provider: DocProvider = { load: () => blankDocxBytes() };

/** Loopback WebSocket onto a real CollabHub (same shape the fidelity suites use). */
function hubSocketClass(hub: CollabHub) {
  let n = 0;
  return class HubSocket {
    constructor(_url: string) {
      const ls: ((ev: { data: unknown }) => void)[] = [];
      const conn: Connection = {
        id: `fm-c${n++}`,
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

/**
 * The demo's App hardwires IndexedDbBundleStore, and useCollab AWAITS
 * `store.get(docId)` before it joins — so with no indexedDB the session never
 * becomes ready and the File trigger never enables. This is the smallest fake
 * that lets the read resolve "no bundle".
 */
function fakeIndexedDB() {
  const request = (result: unknown) => {
    const r: { result: unknown; onsuccess: null | (() => void); onerror: null | (() => void); onupgradeneeded: null | (() => void) } =
      { result, onsuccess: null, onerror: null, onupgradeneeded: null };
    setTimeout(() => r.onsuccess?.(), 0);
    return r;
  };
  const objectStore = {
    get: () => request(undefined),
    put: () => request(undefined),
    delete: () => request(undefined),
  };
  const db = {
    createObjectStore: () => objectStore,
    transaction: () => ({ objectStore: () => objectStore }),
  };
  return { open: () => request(db) };
}

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

describe("the collaborative screen", () => {
  it("never offers a live Open — File is enabled in a live session, and Open inside it is refused", async () => {
    const hub = new CollabHub(provider);
    const prevSocket = globalThis.WebSocket;
    const prevIdb = (globalThis as { indexedDB?: unknown }).indexedDB;
    (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub);
    (globalThis as { indexedDB: unknown }).indexedDB = fakeIndexedDB();
    try {
      const host = render(createElement(App, {
        url: "ws://loopback/collab",
        httpBase: "http://loopback",
        docId: "file-menu-doc",
        clientId: "file-menu-client",
        name: "Ada",
      }));
      const trigger = () => byId(host, "file-menu") as HTMLButtonElement | null;
      expect(trigger(), "the collaborative screen must have a File menu").toBeTruthy();

      // Wait for the live session: the trigger gates on session.ready, exactly
      // as the Download button it replaced did.
      for (let i = 0; i < 120 && trigger()!.disabled; i++) await tick();
      expect(trigger()!.disabled, "session never became ready").toBe(false);

      click(trigger());
      const item = byId(host, "file-open") as HTMLButtonElement;
      expect(item, "Open must be OFFERED and greyed out, not hidden").toBeTruthy();
      expect(item.disabled, "Open must never be live in a session").toBe(true);
      // No .docx picker on this screen — nothing to click through to. (Scoped
      // to the menu: the ribbon has its own picker for images, which DO ride
      // the wire as intents and are therefore fine.)
      expect(host.querySelector('.filemenu input[type="file"]')).toBeNull();
      // Download IS offered here (it replaced the standalone button).
      expect(byId(host, "file-download")).toBeTruthy();
    } finally {
      (globalThis as { WebSocket: unknown }).WebSocket = prevSocket;
      (globalThis as { indexedDB: unknown }).indexedDB = prevIdb;
    }
  });
});
