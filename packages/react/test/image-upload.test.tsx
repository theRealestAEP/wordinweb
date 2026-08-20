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
import { CollabHub, blankDocxBytes, DEFAULT_LIMITS, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import type { DocxDocument } from "@wordinweb/core";

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);
const SVG = new TextEncoder().encode(
  `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><rect width="32" height="32"/></svg>`,
);

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

  it("still takes an SVG — the collab restriction must not leak into local documents", async () => {
    // The fix narrows what a SHARED document offers, because the wire
    // allowlist is raster-only. A local document has nothing to agree with,
    // so it keeps SVG; asserting this is what stops the narrowing from being
    // applied everywhere as the easy fix.
    const ed = await mountLocal();
    expect(ed.api().imageAccept()).toContain("svg");
    await ed.click();
    let outcome: string | undefined;
    await act(async () => { outcome = await ed.api().insertImage(new Blob([SVG], { type: "image/svg+xml" })); });
    await tick();
    expect(outcome).toBe("inserted");
    expect(ed.mediaParts().length).toBe(1);
    await ed.unmount();
  });

  it("stamps a picture watermark, filling the part it reserved", async () => {
    // The watermark path reserves its media part PENDING, because in a room
    // the bytes arrive out of band. A local document is its own source, so
    // "inserted" has to mean the bytes are really in the package — a part
    // left pending here would render as a placeholder forever.
    const ed = await mountLocal();
    let outcome: string | undefined;
    await act(async () => { outcome = await ed.api().insertPictureWatermark(new Blob([PNG], { type: "image/png" })); });
    await tick();

    expect(outcome).toBe("inserted");
    const parts = ed.mediaParts();
    expect(parts.length).toBe(1);
    expect(ed.doc().pkg.binary(parts[0])).toEqual(PNG);
    expect(ed.doc().pendingMedia.size).toBe(0);
    await ed.unmount();
  });

  it("refuses an SVG watermark in a local document too", async () => {
    // Unlike insertImage, which keeps SVG locally: a VML v:imagedata cannot
    // carry an SVG at all, so this restriction is about the markup rather
    // than about what the wire agrees on.
    const ed = await mountLocal();
    let outcome: string | undefined;
    await act(async () => { outcome = await ed.api().insertPictureWatermark(new Blob([SVG], { type: "image/svg+xml" })); });
    await tick();
    expect(outcome).toBe("unsupported-format");
    expect(ed.mediaParts()).toEqual([]);
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

/**
 * A DocxView mounted with a HAND-BUILT collab object, so the published-limit
 * field can be set to each of the three things a real session can hand it:
 * a number, an explicit null (a server that publishes no limit), and absent
 * entirely (a client built before the field existed).
 */
async function mountLimitBed(mediaMaxBlobBytes: number | null | undefined) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const uploads: number[] = [];
  let nextId = 900_000;
  const seen: { api: DocxViewApi | null } = { api: null };
  const collab: Record<string, unknown> = {
    submit: () => {},
    submitOp: () => {},
    allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
    uploadMedia: async (bytes: Uint8Array) => {
      uploads.push(bytes.length);
      return { blobSha: "a".repeat(64), bytesLen: bytes.length };
    },
  };
  // ABSENT vs null is the distinction under test — only assign when the case
  // says the field exists at all.
  if (mediaMaxBlobBytes !== undefined) collab.mediaMaxBlobBytes = mediaMaxBlobBytes;
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(),
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      collab,
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  return {
    uploads,
    api: () => seen.api!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

describe("the published size limit — the null contract", () => {
  it("NO published limit means SKIP the check, never substitute a default", async () => {
    // The branch that must stay a no-op. A client that invented a default
    // would refuse uploads the server would have accepted, and the user could
    // not tell a real limit from one the client made up. Skipping is safe
    // precisely because the server still enforces its own number.
    for (const absent of [null, undefined] as const) {
      const ed = await mountLimitBed(absent);
      expect(ed.api().imageMaxBytes()).toBeNull();
      const outcome = await act(async () => ed.api().insertImage(new Blob([PNG], { type: "image/png" })));
      expect(outcome, `limit ${String(absent)} must not block the upload`).toBe("inserted");
      expect(ed.uploads).toEqual([PNG.length]);
      await ed.unmount();
    }
  });

  it("a published limit is enforced as a THRESHOLD, and reported as the configured number", async () => {
    const LIMIT = 8;
    const ed = await mountLimitBed(LIMIT);
    // What the UI will name — read from the session, not from a constant.
    expect(ed.api().imageMaxBytes()).toBe(LIMIT);
    // Over: refused with nothing sent.
    const over = await act(async () => ed.api().insertImage(new Blob([PNG], { type: "image/png" })));
    expect(over).toBe("too-large");
    expect(ed.uploads, "an oversized file must not reach the relay at all").toEqual([]);
    // Exactly at the limit: allowed (the check is `>`, not `>=` — a file the
    // server would accept must not be refused by the client).
    const atLimit = await act(async () => ed.api().insertImage(new Blob([PNG.slice(0, LIMIT)], { type: "image/png" })));
    expect(atLimit).toBe("inserted");
    expect(ed.uploads).toEqual([LIMIT]);
    await ed.unmount();
  });
});


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

  it("an SVG is DECLINED OUT LOUD, and the picker no longer offers one", async () => {
    // USER-HIT BUG. The Pictures picker advertised `image/svg+xml` while the
    // collab insert refused that exact extension with a bare `return`, so
    // choosing an SVG in a shared document did nothing at all — no skeleton,
    // no message, no log. Reported as "it just didn't work and then
    // disappeared", which is precisely what a silent guard looks like from
    // the outside.
    //
    // Two halves, both asserted here because either alone can regress:
    // the picker must not OFFER what the insert refuses, and the insert must
    // still SAY something when it refuses (the API path remains reachable).
    const hub = new CollabHub(provider);
    const relay = installRelayFetch(hub);
    try {
      const a = await mountCollab(hub, "img-svg", "alice");
      await settle();

      // THE OFFER. Read off the live attribute the file dialog actually uses,
      // not the constant behind it — the drift was between exactly those two.
      const picker = a.container.querySelector<HTMLInputElement>('input[type="file"][accept*="image/png"]');
      expect(picker, "the Pictures file input should be mounted").toBeTruthy();
      expect(picker!.accept).not.toContain("svg");
      expect(picker!.accept).toContain("image/png");
      expect(a.api().imageAccept()).not.toContain("svg");

      // THE REFUSAL, still reachable through the API — and now audible.
      await a.click();
      const before = Buffer.from(a.liveDoc().save());
      let outcome: string | undefined;
      await act(async () => { outcome = await a.api().insertImage(new Blob([SVG], { type: "image/svg+xml" })); });
      await settle();
      expect(outcome).toBe("unsupported-format");
      // Nothing was uploaded, reserved, or forked.
      expect(relay.calls.put).toBe(0);
      expect(a.liveDoc().mediaMeta.size).toBe(0);
      expect(Buffer.from(a.liveDoc().save())).toEqual(before);
      await a.unmount();
    } finally {
      relay.restore();
    }
  });

  it("picking an unsupported file through the TOOLBAR produces a visible notice, never a no-op", async () => {
    // The user's actual path: the file dialog, not the API. Whatever the
    // guard decides, the person who picked the file has to be told — a
    // control that accepts a click and answers with nothing is the failure
    // this whole class keeps reproducing.
    const hub = new CollabHub(provider);
    const relay = installRelayFetch(hub);
    try {
      const a = await mountCollab(hub, "img-svg-ui", "alice");
      await settle();
      await a.click();
      // The user's route to the picker: the Insert tab, where the Pictures
      // button lives — and where its answer has to appear.
      const insertTab = a.container.querySelector<HTMLButtonElement>('button[data-tab="insert"]');
      expect(insertTab, "the Insert tab should be present").toBeTruthy();
      await act(async () => { insertTab!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
      await tick();
      const picker = a.container.querySelector<HTMLInputElement>('input[type="file"][accept*="image/png"]')!;
      // A file dialog can still hand back an SVG (drag-drop, "All files",
      // a stale dialog) — `accept` is a filter, never a guarantee.
      const file = new File([SVG], "logo.svg", { type: "image/svg+xml" });
      Object.defineProperty(picker, "files", { value: [file], configurable: true });
      await act(async () => { picker.dispatchEvent(new Event("change", { bubbles: true })); });
      await settle();

      const notice = a.container.querySelector("[data-dxw-image-status]");
      expect(notice, "an unsupported pick must produce a visible notice").toBeTruthy();
      expect(notice!.textContent).toMatch(/PNG/i);
      expect(relay.calls.put).toBe(0);
      await a.unmount();
    } finally {
      relay.restore();
    }
  });

  it("a PENDING image is selectable and deletable — the skeleton is not a dead box", async () => {
    // OWNER-REPORTED: "users should actually be able to delete an image before
    // it even finishes downloading from the server."
    //
    // A participant looking at "Image unavailable" has a reservation whose
    // bytes may never arrive (nobody holding them is online). Before this, that
    // box could not be selected at all, so it could not be removed either — the
    // document had a permanent hole only a re-share could clear.
    //
    // NOT auto-invalidation: the bytes are fine for everyone else and come back
    // when a holder reconnects (doc 16 §7). This is the USER choosing to remove
    // it, which is an ordinary removeDrawing and converges like any other edit.
    const hub = new CollabHub(provider);
    const relay = installRelayFetch(hub);
    let a: Awaited<ReturnType<typeof mountCollab>> | null = null;
    let b: Awaited<ReturnType<typeof mountCollab>> | null = null;
    try {
      a = await mountCollab(hub, "img-del", "alice");
      await settle();
      await a.click();
      await act(async () => { await a!.api().insertImage(new Blob([PNG], { type: "image/png" })); });
      await settle();
      expect(a.container.querySelector(".dxw-page img")).toBeTruthy();

      // B joins into a relay that will not serve the bytes, so B holds the
      // reservation and paints the skeleton — the state the user reported.
      relay.restore();
      const prior = globalThis.fetch;
      (globalThis as unknown as { fetch: unknown }).fetch = async () => ({
        ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0),
      });
      try {
        b = await mountCollab(hub, "img-del", "bob");
        await settle();
        const skeleton = b.container.querySelector<HTMLElement>("[data-dxw-media-state]");
        expect(skeleton, "B should paint a skeleton for the image it cannot fetch").toBeTruthy();
        // There are THREE skeleton states, not two: "fetching", "waiting" and
        // "unavailable" (collab/src/media.ts). This scenario produces
        // "waiting" — the bytes are gone from the relay and B is waiting for a
        // holder to volunteer. The hit test below matches on the ATTRIBUTE'S
        // PRESENCE rather than its value, so it is state-independent by
        // construction and a fourth state would be covered the day it exists.
        expect(["fetching", "waiting", "unavailable"]).toContain(skeleton!.dataset.dxwMediaState);
        expect(b.container.querySelector(".dxw-page img")).toBeNull();

        // CLICK IT. The skeleton is a div, not an <img> — the whole bug.
        await act(async () => {
          const opts = { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0 };
          skeleton!.dispatchEvent(new MouseEvent("mousedown", opts));
          skeleton!.dispatchEvent(new MouseEvent("mouseup", opts));
        });
        await tick();
        expect(
          b.container.querySelector("[data-dxw-object-selection]"),
          "clicking a pending image should select it like any other drawing",
        ).toBeTruthy();

        // DELETE IT, and the removal must reach the other participant — a
        // local-only delete would be a silent fork.
        const target = (b.container.contains(document.activeElement)
          ? (document.activeElement as HTMLElement)
          : b.container.querySelector<HTMLElement>("textarea")) ?? b.container;
        await act(async () => {
          target.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true }));
        });
        await settle();
        expect(b.container.querySelector("[data-dxw-media-state]"), "the skeleton should be gone on B").toBeNull();
        await expect
          .poll(() => a!.container.querySelectorAll(".dxw-page img").length, { timeout: 4000 })
          .toBe(0);
        // THE MODEL is what has to agree — the drawing is gone on both sides.
        //
        // Whole-package byte-identity is deliberately NOT the oracle here, and
        // this is the one scenario where using it would be wrong rather than
        // merely strict: A holds `word/media/image1.png` and B never received
        // those bytes at all, so the packages differ for a reason that has
        // nothing to do with convergence (measured: A ['word/media/image1.png'],
        // B []). Media parts are outside the byte-identity oracle for exactly
        // this reason — pixel arrival is asynchronous and per-replica.
        const hasDrawing = (e: { name: string; children: unknown[] }): boolean =>
          e.name.endsWith(":drawing") || (e.children as typeof e[]).some(hasDrawing);
        expect(hasDrawing(a.liveDoc().docRoot as never), "the drawing should be gone on A").toBe(false);
        expect(hasDrawing(b.liveDoc().docRoot as never), "the drawing should be gone on B").toBe(false);

        // UNDO, and the ordering worry it raises. `removeDrawing` has NO
        // inverse — invert.ts falls through to `default: return null` — so
        // undoing this delete is a documented cannot-undo, not a restore.
        // That is what makes "the bytes might arrive mid-undo" moot rather
        // than merely unlikely: there is no window in which the reservation
        // comes back to race them. Pinned because if an inverse is ever added,
        // this assertion is where the race has to be thought about again.
        await act(async () => {
          target.dispatchEvent(new KeyboardEvent("keydown", { key: "z", metaKey: true, bubbles: true, cancelable: true }));
        });
        await settle();
        expect(hasDrawing(b.liveDoc().docRoot as never), "undo must not resurrect a drawing that has no inverse").toBe(false);
        expect(hasDrawing(a.liveDoc().docRoot as never), "and must not fork the room trying").toBe(false);
      } finally {
        (globalThis as unknown as { fetch: unknown }).fetch = prior;
      }
    } finally {
      await a?.unmount();
      await b?.unmount();
      relay.restore();
    }
  });

  it("an OVERSIZED image is refused locally, before a single byte is uploaded", async () => {
    // The limit travels FORWARD (published in the welcome) rather than being
    // read backward out of a 413, so the refusal can happen before the file is
    // read, decoded, sealed, hashed or sent — the user finds out immediately
    // instead of after waiting out a whole upload that was always going to
    // fail.
    //
    // THE LIMIT HERE IS INJECTED, never a literal. A test written against a
    // hardcoded 5MB passes just as happily when the client ignores the
    // published number and uses its own default — which is the bug this
    // design exists to prevent, and the dev stack (50MB) would expose it in
    // production only.
    const TINY = 4; // bytes — smaller than PNG, so PNG is "oversized" here
    const limits = { ...DEFAULT_LIMITS, media: { ...DEFAULT_LIMITS.media, maxBlobBytes: TINY } };
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, undefined, limits);
    const relay = installRelayFetch(hub);
    try {
      const a = await mountCollab(hub, "img-big", "alice");
      await settle();
      // The client learned the server's number — not a default of its own.
      expect(a.api().imageMaxBytes()).toBe(TINY);
      await a.click();
      const before = Buffer.from(a.liveDoc().save());
      let outcome: string | undefined;
      await act(async () => { outcome = await a.api().insertImage(new Blob([PNG], { type: "image/png" })); });
      await settle();
      expect(outcome).toBe("too-large");
      // NOTHING went out: no PUT attempted, no reservation, document untouched.
      expect(relay.calls.put).toBe(0);
      expect(a.liveDoc().mediaMeta.size).toBe(0);
      expect(Buffer.from(a.liveDoc().save())).toEqual(before);
      // …and a file INSIDE the injected limit still goes through the same path,
      // so the check is a threshold and not a blanket refusal.
      await act(async () => { outcome = await a.api().insertImage(new Blob([PNG.slice(0, TINY)], { type: "image/png" })); });
      await settle();
      expect(outcome).toBe("inserted");
      expect(relay.calls.put).toBe(1);
      await a.unmount();
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
