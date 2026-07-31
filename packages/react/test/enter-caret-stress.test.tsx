// @vitest-environment jsdom
/**
 * ENTER + CARET-MOVE STRESS: two live CollabEditors over a delayed transport,
 * driven with LONG (200+ op) seeded-random interleavings that deliberately mix
 *   (a) Enter (new paragraphs / splitParagraph intents),
 *   (b) caret JUMPS via clickAt(x,y) to a different text node — the pattern that
 *       trips commit()'s `caretMoved` full-refresh() branch (editor.ts ~6039),
 *   (c) fast multi-pending bursts with NO drain between keystrokes,
 * then asserts BYTE-IDENTICAL convergence between both editors and the server.
 *
 * Motivation: the existing demo-fuzz is SHORT (60-80 ops, ~50 log entries — it
 * barely crosses the checkpoint boundary) and its click() always hits the same
 * middle span, so it rarely exercises repeated caret jumps between DIFFERENT
 * text nodes interleaved with structural edits. The `caretMoved` change forces
 * an EXTRA doc.refresh() (= stableIds.assignFromRoots + prune) on the optimistic
 * replica for the first edit after every caret move; the server mirror only
 * refreshes on STRUCTURAL intents (session.ts ~171). This test stresses whether
 * that asymmetry lets the optimistic doc's id table (or a subsequent intent's
 * encoded block/run ids) drift out of step with the mirror → divergence.
 *
 * Deterministic (seeded mulberry32 PRNG, no Date.now/Math.random in the driver)
 * so any failure replays exactly. A failing seed is added as a permanent case.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { serializeXml } from "@wordinweb/core";

const provider: DocProvider = { load: () => blankDocxBytes() };
let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs: number) {
  const ns = `s${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => (delayMs > 0 ? setTimeout(fn, delayMs) : fn());
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `${ns}c${n++}`, send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))) };
    let opened = false;
    return { send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 40) { for (let i = 0; i < n; i++) await tick(); }

/** Deterministic PRNG (mulberry32) so failures replay exactly by seed. */
function rng(seed: number) {
  let a = seed >>> 0;
  return () => {
    a |= 0; a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

interface Client {
  name: string;
  container: HTMLElement;
  key: (k: string, mods?: { ctrl?: boolean; shift?: boolean }) => Promise<void>;
  clickAt: (x: number, y: number) => Promise<void>;
  clickSpan: (idx: number) => Promise<void>;
  spanCount: () => number;
  xml: () => string;
  unmount: () => Promise<void>;
}

async function mountClient(hub: CollabHub, docId: string, name: string, delayMs: number): Promise<Client> {
  const factory = factoryFor(hub, delayMs);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(CollabEditor, { url: "ws://x", docId, clientId: name, toolbar: false, createSocket: factory }));
  });
  for (let i = 0; i < 40 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
  return {
    name, container,
    key: async (k, mods = {}) => {
      await act(async () => {
        target().dispatchEvent(new KeyboardEvent("keydown", { key: k, ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, bubbles: true, cancelable: true }));
        await new Promise((r) => setTimeout(r, 1));
      });
    },
    // Positional click — lands the caret in whatever line sits at (x,y); used to
    // JUMP the caret to a different text node (the caretMoved trigger).
    clickAt: async (x, y) => {
      const surface = container.querySelector<HTMLElement>(".dxw-page")!.firstElementChild as HTMLElement;
      await act(async () => {
        const o = { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0 };
        surface.dispatchEvent(new MouseEvent("mousedown", o));
        surface.dispatchEvent(new MouseEvent("mouseup", o));
      });
    },
    // Click a SPECIFIC span by index — a deterministic caret jump to a known run.
    clickSpan: async (idx) => {
      const page = container.querySelector<HTMLElement>(".dxw-page")!;
      const spans = [...page.querySelectorAll("span")];
      const span = spans[Math.min(idx, spans.length - 1)] ?? page;
      await act(async () => {
        const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
        span.dispatchEvent(new MouseEvent("mousedown", o));
        span.dispatchEvent(new MouseEvent("mouseup", o));
      });
    },
    spanCount: () => container.querySelector<HTMLElement>(".dxw-page")!.querySelectorAll("span").length,
    xml: () => {
      const kk = Object.keys(container).find((x) => x.startsWith("__reactContainer$"))!;
      const stack: unknown[] = [(container as unknown as Record<string, unknown>)[kk]];
      let g = 0;
      while (stack.length && g++ < 6000) {
        const f = stack.pop() as { memoizedProps?: { collab?: { doc?: { docRoot: unknown } } }; child?: unknown; sibling?: unknown } | null;
        if (!f) continue;
        const d = f.memoizedProps?.collab?.doc;
        if (d) return serializeXml(d.docRoot as never);
        if (f.child) stack.push(f.child);
        if (f.sibling) stack.push(f.sibling);
      }
      throw new Error("no doc");
    },
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

function serverXml(hub: CollabHub, docId: string): string {
  const room = (hub as unknown as { rooms: Map<string, { session: { doc: { docRoot: unknown } } }> }).rooms.get(docId)!;
  return serializeXml(room.session.doc.docRoot as never);
}
function rejectionReport(hub: CollabHub, docId: string): { total: number; rejected: string[] } {
  const room = (hub as unknown as { rooms: Map<string, { session: { entriesSince: (n: number) => { kind: string; reason?: string }[] } }> }).rooms.get(docId)!;
  const entries = room.session.entriesSince(0);
  return { total: entries.length, rejected: entries.filter((e) => e.kind === "rejected").map((e) => e.reason ?? "?") };
}

const ALPHABET = "abcdefghij ";

/**
 * Enter-heavy, caret-jump-heavy interleaving. Each turn a client takes a burst
 * (drained afterwards — the SUPPORTED discipline, so a correct engine MUST
 * converge). Within a burst we bias toward Enter and, crucially, prepend a
 * caret jump to a different text node before typing, to force the caretMoved
 * refresh right before a structural or text intent is encoded.
 */
async function enterCaretRound(seed: number, ops: number, delayMs: number) {
  const rand = rng(seed);
  const hub = new CollabHub(provider);
  const docId = `ec-${seed}`;
  const a = await mountClient(hub, docId, "alice", delayMs);
  const b = await mountClient(hub, docId, "bob", delayMs);
  await settle(10);
  await a.clickAt(12, 12);
  await b.clickAt(12, 12);
  const clients = [a, b];
  let done = 0;
  while (done < ops) {
    const c = clients[rand() < 0.5 ? 0 : 1];
    // JUMP the caret before this burst: either a positional click at a random
    // line band (y grows with paragraph count) or a specific span index. This is
    // the caretMoved trigger — a different text node than the last commit's.
    if (rand() < 0.85) {
      if (rand() < 0.5) {
        const y = 8 + Math.floor(rand() * 320);
        const x = 6 + Math.floor(rand() * 120);
        await c.clickAt(x, y);
      } else {
        const n = c.spanCount();
        if (n > 0) await c.clickSpan(Math.floor(rand() * n));
      }
    }
    const burst = 1 + Math.floor(rand() * 6);
    for (let j = 0; j < burst && done < ops; j++, done++) {
      const roll = rand();
      // Enter-heavy: ~28% Enter, plus mid-burst caret jumps to keep tripping
      // caretMoved between intents of the SAME burst (multi-pending window).
      if (roll < 0.42) await c.key(ALPHABET[Math.floor(rand() * ALPHABET.length)]);
      else if (roll < 0.70) await c.key("Enter");
      else if (roll < 0.80) await c.key("Backspace");
      else if (roll < 0.88) {
        const n = c.spanCount();
        if (n > 0) await c.clickSpan(Math.floor(rand() * n));
      }
      else if (roll < 0.95) await c.key(rand() < 0.5 ? "ArrowLeft" : "ArrowRight");
      else await c.key("Delete");
    }
    await settle(10); // drain: echoes + cross-client broadcast land
  }
  await settle(80);
  const rep = rejectionReport(hub, docId);
  console.error(`seed ${seed}: ${rep.total} entries, ${rep.rejected.length} rejected`, rep.rejected.slice(0, 6));
  const sx = serverXml(hub, docId);
  const ax = a.xml();
  const bx = b.xml();
  await a.unmount();
  await b.unmount();
  expect(ax, `seed ${seed}: A vs SERVER (rejected: ${rep.rejected.length})`).toBe(sx);
  expect(bx, `seed ${seed}: B vs SERVER (rejected: ${rep.rejected.length})`).toBe(sx);
}

/**
 * FAST BURSTS: a single client fires a LONG multi-pending burst (many keys +
 * Enter, NO drain between keystrokes), then only drains at the very end. This
 * puts many intents in flight simultaneously (the "fast typing breaks" report),
 * repeatedly crossing the fast-path / reload boundary in replica.ts, with Enter
 * splits interleaved. Single client, so it is a supported (drained-at-end) run
 * and MUST converge.
 */
async function fastBurstRound(seed: number, chunks: number, delayMs: number) {
  const rand = rng(seed);
  const hub = new CollabHub(provider);
  const docId = `fb-${seed}`;
  const a = await mountClient(hub, docId, "alice", delayMs);
  const b = await mountClient(hub, docId, "bob", delayMs);
  await settle(10);
  await a.clickAt(12, 12);
  await b.clickAt(12, 12);
  for (let ch = 0; ch < chunks; ch++) {
    const c = ch % 2 === 0 ? a : b;
    // Occasionally jump the caret before the burst (caretMoved), then FIRE a
    // long burst with no per-key drain.
    if (rand() < 0.7) {
      const n = c.spanCount();
      if (n > 0) await c.clickSpan(Math.floor(rand() * n));
    }
    const len = 6 + Math.floor(rand() * 12);
    for (let i = 0; i < len; i++) {
      const roll = rand();
      if (roll < 0.6) await c.key(ALPHABET[Math.floor(rand() * ALPHABET.length)]);
      else if (roll < 0.82) await c.key("Enter");
      else await c.key("Backspace");
      // NO settle here — keep them pending.
    }
    await settle(12); // drain this client's burst before the other acts
  }
  await settle(80);
  const rep = rejectionReport(hub, docId);
  console.error(`fastburst seed ${seed}: ${rep.total} entries, ${rep.rejected.length} rejected`, rep.rejected.slice(0, 6));
  const sx = serverXml(hub, docId);
  const ax = a.xml();
  const bx = b.xml();
  await a.unmount();
  await b.unmount();
  expect(ax, `fastburst seed ${seed}: A vs SERVER`).toBe(sx);
  expect(bx, `fastburst seed ${seed}: B vs SERVER`).toBe(sx);
}

describe("enter + caret-move stress: long runs, byte-identical convergence", () => {
  it("seed 101 (220 ops) converges", async () => { await enterCaretRound(101, 220, 3); }, 120000);
  it("seed 102 (220 ops, fast wire) converges", async () => { await enterCaretRound(102, 220, 1); }, 120000);
  it("seed 103 (260 ops) converges", async () => { await enterCaretRound(103, 260, 4); }, 120000);
  it("seed 104 (260 ops, slow wire) converges", async () => { await enterCaretRound(104, 260, 7); }, 120000);
  it("seed 105 (300 ops) converges", async () => { await enterCaretRound(105, 300, 2); }, 150000);
});

describe("fast bursts: multi-pending keystrokes with Enter, no per-key drain", () => {
  it("seed 201 converges", async () => { await fastBurstRound(201, 14, 3); }, 120000);
  it("seed 202 (fast wire) converges", async () => { await fastBurstRound(202, 16, 1); }, 120000);
  it("seed 203 (slow wire) converges", async () => { await fastBurstRound(203, 14, 6); }, 120000);
});
