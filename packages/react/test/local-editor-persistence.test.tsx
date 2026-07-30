// @vitest-environment jsdom
/**
 * The local (landing) editor persists — the owner's complaint made a test:
 * "what's the point of offline edits if there is no way to save?". Editing
 * locally and refreshing must reopen the same document, a failed autosave
 * must be SAID (zero custody: this browser is the only durable copy), and
 * File > New must bank the current document instead of losing it.
 *
 * "Reload" here is a real simulation: unmount the whole tree, then mount a
 * FRESH LocalEditor sharing only the store — exactly what a browser refresh
 * preserves. The restored text is asserted from the re-rendered DOM, so the
 * test fails if either the write or the restore-read is removed.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { blankDocxBytes } from "@wordinweb/server";
import { InMemoryBundleStore, type DocBundle } from "@wordinweb/collab/client";
import { DocxDocument } from "@wordinweb/core";
import { LocalEditor, LOCAL_AUTOSAVE_KEY } from "../../../examples/anon-share/src/local-editor";

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

const prevFetch = globalThis.fetch;
beforeEach(() => {
  // The blank-template endpoint, so New (and a broken restore path) has
  // deterministic bytes instead of a network error.
  globalThis.fetch = vi.fn(async () =>
    new Response(blankDocxBytes() as unknown as BodyInit, { status: 200 }),
  ) as unknown as typeof fetch;
});
afterEach(() => { globalThis.fetch = prevFetch; });

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);

function click(el: HTMLElement | null) {
  expect(el).toBeTruthy();
  act(() => { el!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}

/** Click into the page and type, the way a browser delivers it. */
async function typeText(host: HTMLElement, text: string) {
  for (let i = 0; i < 100 && !host.querySelector(".dxw-page"); i++) await tick();
  const page = host.querySelector<HTMLElement>(".dxw-page")!;
  expect(page, "the editor must have painted a page").toBeTruthy();
  const surface = page.querySelector("span") ?? page;
  await act(async () => {
    const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    surface.dispatchEvent(new MouseEvent("mousedown", opts));
    surface.dispatchEvent(new MouseEvent("mouseup", opts));
    surface.dispatchEvent(new MouseEvent("click", opts));
  });
  await tick();
  const target = (document.activeElement as HTMLElement) ?? host;
  await act(async () => {
    for (const key of text) {
      target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
    }
  });
  await tick();
}

function docText(bytes: Uint8Array): string {
  const doc = DocxDocument.load(bytes);
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(doc.docRoot as never);
}

const editor = (store: InMemoryBundleStore, extra: Record<string, unknown> = {}) =>
  createElement(LocalEditor, {
    httpBase: "http://blankhost",
    onGoLive: async () => {},
    store,
    autosaveMs: 30,
    initialBytes: blankDocxBytes(),
    ...extra,
  });

describe("local editor persistence", () => {
  it("typed work autosaves, and survives a simulated reload", async () => {
    const store = new InMemoryBundleStore();
    const host = render(editor(store));
    await typeText(host, "PERSISTME");

    // The throttled autosave lands within its interval.
    let saved: DocBundle | null = null;
    for (let i = 0; i < 100 && !saved; i++) { await tick(); saved = await store.get(LOCAL_AUTOSAVE_KEY); }
    expect(saved, "the autosave slot must be written").toBeTruthy();
    expect(docText(saved!.confirmedBytes)).toContain("PERSISTME");

    // Reload: tear the whole tree down, mount fresh sharing only the store.
    unmountAll();
    const host2 = render(editor(store, { initialBytes: undefined }));
    for (let i = 0; i < 200 && !(host2.textContent ?? "").includes("PERSISTME"); i++) await tick();
    expect(host2.textContent, "the restored editor must show the typed text").toContain("PERSISTME");
  });

  it("a failed autosave raises the banner instead of being swallowed", async () => {
    const store = new InMemoryBundleStore();
    store.put = async () => { throw new DOMException("quota", "QuotaExceededError"); };
    const host = render(editor(store));
    await typeText(host, "x");
    for (let i = 0; i < 100 && !byId(host, "local-persist-banner"); i++) await tick();
    expect(byId(host, "local-persist-banner"), "the failed write must be surfaced").toBeTruthy();
    // The escape hatch rides the banner: the one true copy is a download.
    expect(byId(host, "local-persist-download")).toBeTruthy();
  });

  it("File > New banks the current document as a dated local entry", async () => {
    const store = new InMemoryBundleStore();
    const host = render(editor(store));
    await typeText(host, "KEEPTHIS");
    click(byId(host, "file-menu"));
    click(byId(host, "file-new"));
    let archived = (await store.list()).filter((s) => s.kind === "local" && s.key !== LOCAL_AUTOSAVE_KEY);
    for (let i = 0; i < 100 && archived.length === 0; i++) {
      await tick();
      archived = (await store.list()).filter((s) => s.kind === "local" && s.key !== LOCAL_AUTOSAVE_KEY);
    }
    expect(archived).toHaveLength(1);
    const banked = await store.get(archived[0].key);
    expect(docText(banked!.confirmedBytes)).toContain("KEEPTHIS");
  });

  it("refuses to blank the document when the New-archive write fails", async () => {
    const store = new InMemoryBundleStore();
    const host = render(editor(store));
    await typeText(host, "DONTLOSE");
    store.put = async () => { throw new DOMException("quota", "QuotaExceededError"); };
    click(byId(host, "file-menu"));
    click(byId(host, "file-new"));
    for (let i = 0; i < 100 && !byId(host, "local-persist-banner"); i++) await tick();
    expect(byId(host, "local-persist-banner"), "the failed archive must be surfaced").toBeTruthy();
    // The document was NOT replaced by a blank — the work is still on screen.
    expect(host.textContent).toContain("DONTLOSE");
  });
});
