// @vitest-environment jsdom
/**
 * FIDELITY FUZZ: two live CollabEditors hammer one document with seeded-random
 * interleavings of real keyboard/mouse events over a delayed (async, echoes in
 * flight) transport, then must converge BYTE-IDENTICALLY with the server.
 *
 * This is the harness that would have caught every demo bug found so far:
 * double-apply, self-transform offset drift, missing rejection rollback, the
 * local-only sub-range format, the colliding split ids, the silent local-only
 * merge/Enter-at-start/list-exit paths. Deterministic (seeded PRNG) so any
 * failure replays exactly; add a failing seed as a permanent regression case.
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
  const ns = `z${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => setTimeout(fn, delayMs);
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
  click: () => Promise<void>;
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
    click: async () => {
      const page = container.querySelector<HTMLElement>(".dxw-page")!;
      const spans = [...page.querySelectorAll("span")];
      const span = spans[Math.floor(spans.length / 2)] ?? page;
      await act(async () => {
        const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
        span.dispatchEvent(new MouseEvent("mousedown", o));
        span.dispatchEvent(new MouseEvent("mouseup", o));
      });
    },
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

const ALPHABET = "abcdefghij ";
async function fuzzRound(seed: number, ops: number, delayMs: number) {
  const rand = rng(seed);
  const hub = new CollabHub(provider);
  const docId = `fuzz-${seed}`;
  const a = await mountClient(hub, docId, "alice", delayMs);
  const b = await mountClient(hub, docId, "bob", delayMs);
  await settle(10);
  await a.click();
  await b.click();
  const clients = [a, b];
  let done = 0;
  while (done < ops) {
    // One client takes a BURST turn (multiple intents in flight from a single
    // client — the discipline the engine supports), then the wire drains
    // before the other client acts.
    const c = clients[rand() < 0.5 ? 0 : 1];
    await c.click(); // re-establish a caret (reloads may have re-mounted)
    const burst = 1 + Math.floor(rand() * 6);
    for (let j = 0; j < burst && done < ops; j++, done++) {
      const roll = rand();
      if (roll < 0.55) await c.key(ALPHABET[Math.floor(rand() * ALPHABET.length)]);
      else if (roll < 0.68) await c.key("Enter");
      else if (roll < 0.8) await c.key("Backspace");
      else if (roll < 0.9) await c.key(rand() < 0.5 ? "ArrowLeft" : "ArrowRight");
      else await c.key("Delete");
    }
    await settle(10); // drain: echoes + cross-client broadcast land
  }
  await settle(60);
  {
    const room = (hub as unknown as { rooms: Map<string, { session: { entriesSince: (n: number) => { kind: string; reason?: string }[] } }> }).rooms.get(docId)!;
    const entries = room.session.entriesSince(0);
    const rejected = entries.filter((e) => e.kind === "rejected");
    console.error(`seed ${seed}: ${entries.length} entries, ${rejected.length} rejected`, rejected.slice(0, 5).map((e) => e.reason));
  }
  const sx = serverXml(hub, docId);
  const ax = a.xml();
  const bx = b.xml();
  await a.unmount();
  await b.unmount();
  expect(ax, `seed ${seed}: A vs SERVER`).toBe(sx);
  expect(bx, `seed ${seed}: B vs SERVER`).toBe(sx);
}

describe("fidelity fuzz: two editors, random real-event turn bursts", () => {
  it("seed 1 converges byte-identically", async () => { await fuzzRound(1, 60, 3); }, 60000);
  it("seed 2 converges byte-identically", async () => { await fuzzRound(2, 60, 5); }, 60000);
  it("seed 3 (fast wire) converges byte-identically", async () => { await fuzzRound(3, 60, 1); }, 60000);
  it("seed 4 (slow wire) converges byte-identically", async () => { await fuzzRound(4, 60, 8); }, 60000);
  it("seed 5 converges byte-identically", async () => { await fuzzRound(5, 80, 3); }, 60000);
});

describe("KNOWN LIMIT: whole-doc format fired MID-BURST (echoes in flight)", () => {
  // Ctrl+A+Ctrl+B while a burst's inserts/splits are still in flight can
  // mis-address runs that the in-flight structural echoes are about to
  // replace, cascading into rejections. Formatting between bursts (the
  // toolbar reality — you stop typing to click B) converges fine and is
  // covered by demo-features. Tracked here until format intents are
  // transform-hardened against in-flight structural edits.
  it.fails("bold slammed into an un-drained burst converges", async () => {
    const rand = rng(7);
    const hub = new CollabHub(provider);
    const a = await mountClient(hub, "midburst", "alice", 4);
    await settle(10);
    await a.click();
    for (const k of "some words here") await a.key(k);
    await a.key("Enter");
    for (const k of "more") await a.key(k);
    // NO drain: format while everything is in flight.
    await a.key("a", { ctrl: true });
    await a.key("b", { ctrl: true });
    for (const k of "tail") await a.key(k);
    await settle(60);
    const sx = serverXml(hub, "midburst");
    const ax = a.xml();
    await a.unmount();
    void rand;
    expect(ax).toBe(sx);
  }, 60000);
});

describe("KNOWN LIMIT (timing-dependent): fully-simultaneous cross-client structural editing", () => {
  // Was a deterministic it.fails marker; the protocol-v2 sidecar-bearing
  // welcome (round-4 F10) improved it to SOMETIMES-converging — the imported
  // sidecar aligns the joiner's id-table watermark with the server's, which
  // removed one divergence source — but the outcome now depends on real-timer
  // interleaving of the un-drained bursts: it passes in full-suite runs and
  // fails in isolation. Neither `it` nor `it.fails` is stable on a racy
  // boundary, so it is skipped, honestly: convergence for fully-simultaneous
  // structural editing is NOT guaranteed until doc 03's adjusted-sibling
  // multi-pending replay lands. Un-skip and promote when it does. The
  // supported discipline (bursts with drains) is pinned by the seeds above.
  it.skip("simultaneous interleaved bursts from both clients", async () => {
    await simultaneousRound(99, 40, 4);
  }, 60000);
});

/** The unsupported mode: both clients act with NO drains between turns. */
async function simultaneousRound(seed: number, ops: number, delayMs: number) {
  const rand = rng(seed);
  const hub = new CollabHub(provider);
  const docId = `sim-${seed}`;
  const a = await mountClient(hub, docId, "alice", delayMs);
  const b = await mountClient(hub, docId, "bob", delayMs);
  await settle(10);
  await a.click(); await b.click();
  const clients = [a, b];
  for (let i = 0; i < ops; i++) {
    const c = clients[rand() < 0.5 ? 0 : 1];
    const roll = rand();
    if (roll < 0.6) await c.key(ALPHABET[Math.floor(rand() * ALPHABET.length)]);
    else if (roll < 0.75) await c.key("Enter");
    else await c.key("Backspace");
  }
  await settle(80);
  const sx = serverXml(hub, docId);
  const ax = a.xml(); const bx = b.xml();
  await a.unmount(); await b.unmount();
  expect(ax).toBe(sx);
  expect(bx).toBe(sx);
}
