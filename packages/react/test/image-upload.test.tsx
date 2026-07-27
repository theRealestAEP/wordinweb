// @vitest-environment jsdom
/**
 * Image insertion, both modes (user report: "uploading images doesn't work at
 * all").
 *
 * The LOCAL half is verification, not new work: it pins that the toolbar's
 * insert path really lands a media part in the package, and it pins the one
 * silent failure that path has — no caret, no insert, no feedback.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import type { DocxDocument } from "@wordinweb/core";

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);

const glob = globalThis as unknown as Record<string, unknown>;
// jsdom has no image decoder; insertImage measures the bitmap. And rendering
// an image binding builds an object URL. (Same two shims capability-matrix
// installs — jsdom gaps, not product behavior.)
glob.createImageBitmap ??= async () => ({ width: 64, height: 48, close() {} });
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:image-upload-test";
  URL.revokeObjectURL = () => {};
}

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

async function mountLocal() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(),
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const click = async () => {
    const page = container.querySelector<HTMLElement>(".dxw-page")!;
    const span = page.querySelector("span") ?? page;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      span.dispatchEvent(new MouseEvent("mousedown", opts));
      span.dispatchEvent(new MouseEvent("mouseup", opts));
    });
    await tick();
  };
  /** Media parts present in the package (what "the image landed" means). */
  const mediaParts = () => seen.doc!.pkg.names().filter((n) => n.startsWith("word/media/"));
  return {
    container, click, mediaParts,
    api: () => seen.api!,
    doc: () => seen.doc!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

describe("image insert — LOCAL mode (the pre-existing path)", () => {
  it("inserts a media part and a drawing after a caret is placed", async () => {
    const ed = await mountLocal();
    await ed.click();
    expect(ed.mediaParts()).toEqual([]);
    await act(async () => { await ed.api().insertImage(new Blob([PNG], { type: "image/png" })); });
    await tick();

    // The bytes landed in the package…
    const parts = ed.mediaParts();
    expect(parts.length).toBe(1);
    expect(ed.doc().pkg.binary(parts[0])).toEqual(PNG);
    // …and the document references them through a real w:drawing.
    const xml = ed.doc().pkg.text("word/document.xml") ?? "";
    void xml; // document.xml is only rewritten on save(); check the live tree
    const hasDrawing = (e: { name: string; children: unknown[] }): boolean =>
      e.name.endsWith(":drawing") || (e.children as typeof e[]).some(hasDrawing);
    expect(hasDrawing(ed.doc().docRoot as never)).toBe(true);
    await ed.unmount();
  });

  it("works with NO caret placed (reaching for the toolbar first)", async () => {
    // WAS a silent no-op, and a plausible half of "images don't work at all":
    // clicking Insert-image before clicking into the page did nothing, with no
    // feedback. Picking a file is an unambiguous request, so the insert now
    // falls back to the selection's end and then to the document's start.
    const ed = await mountLocal();
    await act(async () => { await ed.api().insertImage(new Blob([PNG], { type: "image/png" })); });
    await tick();
    expect(ed.mediaParts().length).toBe(1);
    await ed.unmount();
  });
});

// ---------------------------------------------------------------- collab bed
const provider: DocProvider = { load: () => blankDocxBytes() };

let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `f${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => (delayMs > 0 ? setTimeout(fn, delayMs) : fn());
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `${ns}c${n++}`, send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))) };
    let opened = false;
    return {
      send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") ls.push(cb as never);
        else if (!opened) { opened = true; (cb as () => void)(); }
      },
    } as unknown as WebSocket;
  };
}
async function settle(n = 30) { for (let i = 0; i < n; i++) await tick(); }

/**
 * Route global fetch at the REAL hub's relay methods, so uploads are
 * hash-verified by the same admission code the HTTP route calls. Returns
 * counters so a test can prove a byte actually moved.
 */
function installRelayFetch(hub: CollabHub) {
  const calls = { put: 0, get: 0, miss: 0 };
  const prior = globalThis.fetch;
  (globalThis as unknown as { fetch: unknown }).fetch = async (url: string, init?: { method?: string; body?: Uint8Array }) => {
    const m = /\/docs\/([^/]+)\/media\/([0-9a-f]{64})$/.exec(String(url));
    if (!m) return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) };
    const [, docId, sha] = m;
    if (init?.method === "PUT") {
      calls.put++;
      const status = await hub.mediaUpload(decodeURIComponent(docId), sha, init.body!);
      return { ok: status === 200 || status === 201, status, arrayBuffer: async () => new ArrayBuffer(0) };
    }
    calls.get++;
    const bytes = hub.mediaDownload(decodeURIComponent(docId), sha);
    if (!bytes) { calls.miss++; return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) }; }
    return { ok: true, status: 200, arrayBuffer: async () => bytes.slice().buffer as ArrayBuffer };
  };
  return { calls, restore: () => { (globalThis as unknown as { fetch: unknown }).fetch = prior; } };
}

async function mountCollab(hub: CollabHub, docId: string, clientId: string) {
  const factory = factoryFor(hub);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let api: DocxViewApi | null = null;
  await act(async () => {
    root.render(createElement(CollabEditor, {
      url: "ws://x", docId, clientId, createSocket: factory,
      httpBase: "http://relay", onReady: (a: DocxViewApi) => { api = a; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const liveDoc = (): DocxDocument => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 5000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { doc?: DocxDocument } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const d = f.memoizedProps?.collab?.doc;
      if (d) return d;
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    throw new Error("collab doc not found");
  };
  const click = async () => {
    const page = container.querySelector<HTMLElement>(".dxw-page")!;
    const span = page.querySelector("span") ?? page;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      span.dispatchEvent(new MouseEvent("mousedown", opts));
      span.dispatchEvent(new MouseEvent("mouseup", opts));
    });
    await tick();
  };
  /** The one media part, whatever it got named. */
  const part = () => [...liveDoc().mediaMeta.keys()][0];
  return {
    container, click, liveDoc, part,
    api: () => api!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

describe("image insert — COLLAB mode (doc 16 media relay)", () => {
  it("uploads the blob, reserves it on the wire, and a peer eager-fetches the pixels", async () => {
    const hub = new CollabHub(provider);
    const relay = installRelayFetch(hub);
    try {
      const a = await mountCollab(hub, "img-duo", "alice");
      const b = await mountCollab(hub, "img-duo", "bob");
      await settle();
      await a.click();
      await act(async () => { await a.api().insertImage(new Blob([PNG], { type: "image/png" })); });
      await settle();

      // The blob really went over HTTP (not the sequencer).
      expect(relay.calls.put).toBe(1);

      // PLACER: registered + already holding its own pixels.
      const aPart = a.part();
      expect(aPart).toBeTruthy();
      expect(a.liveDoc().mediaStatus(aPart)).toBe("ready");
      expect(Buffer.from(a.liveDoc().media(aPart)!)).toEqual(Buffer.from(PNG));

      // PEER: the reservation replicated AND the eager fetch filled it,
      // with no manual prodding — this is the whole loop.
      const bPart = b.part();
      expect(bPart).toBe(aPart);                        // same deterministic part name
      expect(relay.calls.get).toBeGreaterThanOrEqual(1);
      expect(b.liveDoc().mediaStatus(bPart)).toBe("ready");
      // MEDIA PARITY (asserted explicitly, not via the XML guard: pixel
      // arrival is asynchronous by design — doc 16 §6).
      expect(Buffer.from(b.liveDoc().media(bPart)!)).toEqual(Buffer.from(PNG));

      // And the DOCUMENTS agree, which is the convergence claim.
      expect(Buffer.from(b.liveDoc().save())).toEqual(Buffer.from(a.liveDoc().save()));

      // IT ACTUALLY PAINTS. Asserting mediaStatus alone is not enough: the
      // arriving bytes change the PACKAGE, not the model, so the repaint
      // differ happily kept the skeleton on screen forever while every
      // status assertion above stayed green. Caught in a real browser, pinned
      // here where it is cheap.
      expect(b.container.querySelector(".dxw-media-skeleton")).toBeNull();
      expect(b.container.querySelector(".dxw-page img")).toBeTruthy();
      await a.unmount();
      await b.unmount();
    } finally {
      relay.restore();
    }
  });

  it("a LATE joiner FETCHES the image it never saw inserted", async () => {
    // This was the arc's known gap and is now closed from both ends: a
    // joiner's snapshot carries the image's REGISTRATION but not its declared
    // sha (that address lives only in intents already folded into the
    // snapshot), so the joiner could reserve the box and never fill it. The
    // welcome now carries the address map — in plaintext rooms as
    // `welcome.media`, in encrypted rooms inside the SEALED checkpoint body,
    // so a blind server still learns nothing about part structure.
    const hub = new CollabHub(provider);
    const relay = installRelayFetch(hub);
    try {
      const a = await mountCollab(hub, "img-late", "alice");
      await a.click();
      await act(async () => { await a.api().insertImage(new Blob([PNG], { type: "image/png" })); });
      await settle();

      // Carol joins only now — she never saw the insert broadcast.
      const late = await mountCollab(hub, "img-late", "carol");
      await settle();
      const part = late.part();
      expect(part).toBeTruthy();
      expect(late.liveDoc().mediaStatus(part)).toBe("ready");
      expect(Buffer.from(late.liveDoc().media(part)!)).toEqual(Buffer.from(PNG));
      // …and nothing is left showing a placeholder.
      expect(late.container.querySelector(".dxw-media-skeleton")).toBeNull();
      await a.unmount();
      await late.unmount();
    } finally {
      relay.restore();
    }
  });

  it("a REFUSED upload reserves nothing — no skeleton anywhere", async () => {
    // The ordering rule (doc 16 §5.1) made observable: if the relay says no,
    // the room must look exactly as it did before the user picked a file.
    const hub = new CollabHub(provider);
    const prior = globalThis.fetch;
    (globalThis as unknown as { fetch: unknown }) = globalThis;
    (globalThis as unknown as { fetch: unknown }).fetch = async () => ({
      ok: false, status: 507, arrayBuffer: async () => new ArrayBuffer(0),
    });
    try {
      const a = await mountCollab(hub, "img-refused", "alice");
      const before = Buffer.from(a.liveDoc().save());
      await a.click();
      await act(async () => { await a.api().insertImage(new Blob([PNG], { type: "image/png" })); });
      await settle();
      expect(a.liveDoc().mediaMeta.size).toBe(0);
      expect(a.liveDoc().pendingMedia.size).toBe(0);
      expect(Buffer.from(a.liveDoc().save())).toEqual(before);
      await a.unmount();
    } finally {
      (globalThis as unknown as { fetch: unknown }).fetch = prior;
    }
  });

});
