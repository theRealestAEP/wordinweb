// @vitest-environment jsdom
/**
 * "When you click on a collaborator's name we should jump to that person."
 *
 * The demo's roster chips (examples/anon-share app) become jump buttons: a
 * click scrolls the view to that participant's presence caret via the new
 * DocxViewApi.revealPresence. These tests drive the REAL demo App over a
 * loopback hub with a real second participant, and assert:
 *
 *  1. the ACTUAL scroll target — the rendered text of the paragraph the
 *     peer's caret sits in, not merely that a handler fired;
 *  2. a participant with no caret yet produces the stated notice, no throw;
 *  3. an unresolvable (stale/hostile) position degrades to a notice;
 *  4. the chip is a real, focusable button (keyboard reachable);
 *  5. THE ONE THAT MATTERS MOST: the jump is a pure view operation — zero
 *     layout work, asserted causally via canvas measureText counts and the
 *     __dxwPerf.incr counter, the remote-repaint-scoped pattern.
 */
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport, InMemoryBundleStore } from "@wordinweb/collab/client";
import type { DocxDocument, Paragraph, Run } from "@wordinweb/core";
import { App } from "../src/app";

// ---- causal layout-work counter (remote-repaint-scoped's) ------------------
// Canvas measureText calls are the real cost of layout; a scroll must add 0.
let measureCalls = 0;
function wrapGetContext(proto: { getContext: (...a: unknown[]) => unknown }): void {
  const orig = proto.getContext;
  proto.getContext = function (...args: unknown[]) {
    const ctx = orig.apply(this, args) as { measureText?: (t: string) => unknown } | null;
    if (ctx && typeof ctx.measureText === "function") {
      const m = ctx.measureText.bind(ctx);
      ctx.measureText = (text: string) => {
        measureCalls++;
        return m(text);
      };
    }
    return ctx;
  };
}
wrapGetContext(HTMLCanvasElement.prototype as never);
wrapGetContext((globalThis as { OffscreenCanvas: { prototype: never } }).OffscreenCanvas.prototype);

// ---- scroll recorder -------------------------------------------------------
// jsdom's scrollIntoView is a no-op stub (test setup); record the receiver so
// the assertion is on the actual scroll TARGET.
let scrolled: Element[] = [];
beforeEach(() => {
  scrolled = [];
  Element.prototype.scrollIntoView = function () {
    scrolled.push(this);
  };
});

function docBytes(paras: number): Uint8Array {
  const para = (i: number) =>
    `<w:p><w:r><w:t xml:space="preserve">Zebra${i} the quick brown fox jumps over the lazy dog while the committee deliberates. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(documentXml),
  });
}

/** One loopback socket class per hub — used by BOTH the App (via the global
 * WebSocket) and the peer connection, so conn ids can never collide.
 * `presenceLog` records every server→client presence fan-out, the arrival
 * signal the tests poll (the hub excludes the sender, so a logged entry IS
 * delivery to the App). */
let sockSeq = 0;
function hubSocketClass(hub: CollabHub, presenceLog: { participant: string }[]) {
  const ns = `rj${sockSeq++}-`;
  let n = 0;
  return class HubSocket {
    constructor(_url: string) {
      const ls: ((ev: { data: unknown }) => void)[] = [];
      const conn: Connection = {
        id: `${ns}c${n++}`,
        send: (m: ServerMessage) => {
          if ((m as { t?: string }).t === "presence") {
            presenceLog.push({ participant: (m as { participant: string }).participant });
          }
          setTimeout(() => ls.forEach((l) => l({ data: JSON.stringify(m) })), 1);
        },
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

let mounted: { root: Root; host: HTMLElement }[] = [];
const prevWebSocket = globalThis.WebSocket;
afterEach(() => {
  for (const { root, host } of mounted) {
    act(() => { root.unmount(); });
    host.remove();
  }
  mounted = [];
  (globalThis as { WebSocket: unknown }).WebSocket = prevWebSocket;
});

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function until(cond: () => boolean, label: string, tries = 300) {
  for (let i = 0; i < tries; i++) {
    if (cond()) return;
    await tick();
  }
  if (!cond()) throw new Error(`timeout: ${label}`);
}

const byId = (host: HTMLElement, id: string) => host.querySelector<HTMLElement>(`[data-testid="${id}"]`);

/** The live collab prop DocxView received (fresh presence + doc), via the
 * fiber tree — the remote-repaint-scoped helper. */
function collabOf(host: HTMLElement): { doc: DocxDocument; presence: Record<string, unknown> } {
  const key = Object.keys(host).find((k) => k.startsWith("__reactContainer$"))!;
  const stack: unknown[] = [(host as unknown as Record<string, unknown>)[key]];
  let guard = 0;
  while (stack.length && guard++ < 20000) {
    const f = stack.pop() as {
      memoizedProps?: { collab?: { doc?: DocxDocument | null; presence?: unknown } };
      child?: unknown;
      sibling?: unknown;
    } | null;
    if (!f) continue;
    const c = f.memoizedProps?.collab;
    if (c && c.presence && typeof c.presence === "object") return c as never;
    if (f.child) stack.push(f.child);
    if (f.sibling) stack.push(f.sibling);
  }
  throw new Error("collab session not found in fiber tree");
}

function addrOf(doc: DocxDocument, i: number): { blockId: number; runId: number } {
  const para = doc.sections[0].blocks[i] as Paragraph;
  const run = para.children[0] as Run;
  const ids = doc.stableIds!;
  return { blockId: ids.idOf(para.src!)!, runId: ids.idOf(run.src!)! };
}

/** Mount the demo App on `hub` and connect a named peer; returns both plus
 * a waiter for the peer's presence to have been fanned out AND rendered. */
async function mountWithPeer(hub: CollabHub, docId: string) {
  const presenceLog: { participant: string }[] = [];
  (globalThis as { WebSocket: unknown }).WebSocket = hubSocketClass(hub, presenceLog);
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act(async () => {
    root.render(createElement(App, {
      url: "ws://loopback/collab",
      httpBase: "http://loopback",
      docId,
      clientId: "ada-client",
      name: "Ada",
      store: new InMemoryBundleStore(),
    }));
  });
  mounted.push({ root, host });
  await until(() => !!host.querySelector(".dxw-page"), "app paints");

  const WS = globalThis.WebSocket as new (u: string) => WebSocket;
  const peer = new CollabConnection(createWebSocketTransport(new WS("ws://peer") as never), "bea-client");
  peer.join(docId, undefined, { profile: { name: "Bea", color: "" } });
  await until(() => peer.ready, "peer welcome");
  // The peer's chip renders once the roster lands AND the api is ready.
  await until(() => !!byId(host, "roster-jump"), "peer jump chip renders");
  // Presence has ARRIVED once the hub fanned it to the App's socket; a few
  // more ticks let React fold it into presenceRef before the click.
  const presenceDelivered = async () => {
    await until(() => presenceLog.some((p) => p.participant === "bea-client"), "presence reaches the app");
    await tick(20);
    await tick(20);
  };
  return { host, peer, presenceDelivered };
}

function clickJump(host: HTMLElement) {
  const btn = byId(host, "roster-jump")!;
  act(() => { btn.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
}

describe("roster chip → jump to a collaborator's caret", () => {
  it("clicking the chip scrolls to the paragraph the peer's caret is in", async () => {
    const hub = new CollabHub({ load: () => docBytes(40) } as DocProvider);
    const { host, peer, presenceDelivered } = await mountWithPeer(hub, "jumpdoc");

    const addr = addrOf(peer.doc!, 30);
    await act(async () => { peer.setPresence({ anchor: { blockId: addr.blockId, runId: addr.runId, offset: 3 } }); });
    await presenceDelivered();

    scrolled = [];
    clickJump(host);
    // The actual scroll target: the rendered text span holding paragraph
    // 30's (unique) first word — an anchored match, so scrolling an ancestor
    // (whose text contains every paragraph) or a neighbour cannot pass.
    expect(scrolled.length).toBeGreaterThan(0);
    expect(scrolled[scrolled.length - 1].textContent ?? "").toMatch(/^Zebra30\b/);
    // A successful jump raises no notice.
    expect(byId(host, "jump-notice")).toBeNull();
  }, 60_000);

  it("a participant with no caret yet gets a legible notice, not silence or a throw", async () => {
    const hub = new CollabHub({ load: () => docBytes(6) } as DocProvider);
    const { host } = await mountWithPeer(hub, "nocursor");

    scrolled = [];
    clickJump(host);
    expect(scrolled.length).toBe(0);
    expect(byId(host, "jump-notice")?.textContent).toContain("Bea hasn’t placed a cursor yet");
  }, 60_000);

  it("an unresolvable (stale/hostile) position degrades to a notice, not a throw", async () => {
    const hub = new CollabHub({ load: () => docBytes(6) } as DocProvider);
    const { host, peer, presenceDelivered } = await mountWithPeer(hub, "staledoc");

    // Presence crosses a trust boundary: a payload can reference ids that
    // never existed (or no longer do). The jump must degrade, never throw.
    await act(async () => { peer.setPresence({ anchor: { blockId: 999_999, runId: 999_998, offset: 0 } }); });
    await presenceDelivered();

    scrolled = [];
    clickJump(host);
    expect(scrolled.length).toBe(0);
    expect(byId(host, "jump-notice")?.textContent).toContain("Bea moved their cursor");
  }, 60_000);

  it("the chip is a real focusable button and activation jumps (keyboard reachable)", async () => {
    const hub = new CollabHub({ load: () => docBytes(12) } as DocProvider);
    const { host, peer, presenceDelivered } = await mountWithPeer(hub, "keydoc");

    const addr = addrOf(peer.doc!, 8);
    await act(async () => { peer.setPresence({ anchor: { blockId: addr.blockId, runId: addr.runId, offset: 1 } }); });
    await presenceDelivered();

    const btn = byId(host, "roster-jump")!;
    // Native button semantics carry Enter/Space activation in a real browser
    // (jsdom does not synthesize them, so assert the semantics + activate).
    expect(btn.tagName).toBe("BUTTON");
    expect(btn.hasAttribute("disabled")).toBe(false);
    expect(btn.getAttribute("aria-label")).toBe("Jump to Bea’s cursor");
    act(() => { btn.focus(); });
    expect(document.activeElement).toBe(btn);
    scrolled = [];
    act(() => { btn.click(); });
    expect(scrolled.length).toBeGreaterThan(0);
    expect(scrolled[scrolled.length - 1].textContent ?? "").toMatch(/^Zebra8\b/);
  }, 60_000);

  it("the jump does ZERO layout work — no measure, no incremental pass, no busy gate", async () => {
    type Perf = { incr?: unknown; last?: unknown };
    (globalThis as { __dxwPerf?: Perf }).__dxwPerf = {};
    const hub = new CollabHub({ load: () => docBytes(200) } as DocProvider);
    const { host, peer, presenceDelivered } = await mountWithPeer(hub, "bigdoc");

    // A far-away caret: paragraph 190 lives on a page the virtualized view
    // has NOT mounted, so the jump exercises the materialize/scroll/restore
    // path — the expensive-looking case that must still cost zero layout.
    const addr = addrOf(peer.doc!, 190);
    await act(async () => { peer.setPresence({ anchor: { blockId: addr.blockId, runId: addr.runId, offset: 2 } }); });
    await presenceDelivered();
    await tick(80); // let any presence-triggered work settle before measuring

    // Any relayout would set the input-blocking busy attribute — record it.
    let busySeen = 0;
    const mo = new MutationObserver((records) => {
      for (const r of records) {
        if ((r.target as HTMLElement).hasAttribute?.("data-dxw-layout-busy")) busySeen++;
      }
    });
    mo.observe(host, { subtree: true, attributes: true, attributeFilter: ["data-dxw-layout-busy", "inert"] });

    const perf = (globalThis as { __dxwPerf: Perf }).__dxwPerf;
    perf.incr = undefined;
    // THE LOAD-BEARING COUNTER: paintLayout replaces perf.last on EVERY
    // layout+paint, incremental or global. measureText counts alone are
    // vacuous here — the document is unchanged, so the measurer's width
    // cache satisfies even a full relayout without one canvas call (proved
    // by mutation: a forced global rerender left measureCalls at 0).
    const perfLast0 = perf.last;
    expect(perfLast0).toBeTruthy(); // the mount paint recorded — the sentinel is armed
    const doc = collabOf(host).doc;
    const modelVersion0 = doc.modelVersion;
    const m0 = measureCalls;
    scrolled = [];
    clickJump(host);
    await tick(120); // past DocxView's 60ms repaint fallback — catch queued layout

    // The jump landed…
    expect(scrolled.length).toBeGreaterThan(0);
    expect(scrolled[scrolled.length - 1].textContent ?? "").toMatch(/^Zebra190\b/);
    // …and the view did ZERO layout work for it: no layout pass painted
    // (perf.last identity), not one text measured, no incremental pass ran,
    // nothing marked dirty, input never gated.
    expect(perf.last).toBe(perfLast0);
    expect(measureCalls - m0).toBe(0);
    expect(perf.incr).toBeUndefined();
    expect(doc.modelVersion).toBe(modelVersion0);
    expect(busySeen).toBe(0);
    const container = host.querySelector<HTMLElement>("[data-dxw-layout-busy]");
    expect(container).toBeNull();
    mo.disconnect();
  }, 120_000);
});
