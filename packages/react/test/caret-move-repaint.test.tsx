// @vitest-environment jsdom
/**
 * Regression for the "click-then-type paints nothing" bug (BUGS.md). The fast
 * in-place text path (syncTextModel) mutates the retained layout item WITHOUT
 * bumping doc.modelVersion, so the incremental renderer's page adoption keeps
 * the STALE DOM for the first edit after the caret JUMPS to a different text
 * node (click / arrow) — the glyph only appears on the next keystroke, and in
 * collab that stale view drives cascading divergence as the user re-clicks.
 *
 * The DOM staleness itself is a real-browser layout-reuse effect that jsdom's
 * degenerate layout does not reproduce (every jsdom click converges), so this
 * pins the FIX'S MECHANISM instead: commit() must take the full refresh() path
 * (which bumps modelVersion, disabling page adoption) on the first edit after
 * the caret moves, while continuous typing in one spot keeps the fast path.
 * Verified end-to-end in a real browser; see BUGS.md.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";

const provider: DocProvider = { load: () => blankDocxBytes() };
let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `m${factorySeq++}-`;
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
async function settle(n = 30) { for (let i = 0; i < n; i++) await tick(); }

async function mount(hub: CollabHub, docId: string, clientId: string) {
  const factory = factoryFor(hub);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(CollabEditor, { url: "ws://x", docId, clientId, createSocket: factory }));
  });
  for (let i = 0; i < 40 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
  const keys = async (seq: string[]) => {
    await act(async () => {
      for (const key of seq) {
        target().dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
        await new Promise((r) => setTimeout(r, 2));
      }
    });
  };
  const clickAt = async (x: number, y: number) => {
    const surface = container.querySelector<HTMLElement>(".dxw-page")!.firstElementChild as HTMLElement;
    await act(async () => {
      const o = { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0 };
      surface.dispatchEvent(new MouseEvent("mousedown", o));
      surface.dispatchEvent(new MouseEvent("mouseup", o));
    });
    await tick();
  };
  // Reach the live collab doc (has modelVersion) through the React tree.
  const modelVersion = (): number => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 5000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { doc?: { modelVersion: number } } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const d = f.memoizedProps?.collab?.doc;
      if (d && typeof d.modelVersion === "number") return d.modelVersion;
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    throw new Error("collab doc not found");
  };
  const unmount = async () => { await act(async () => { root.unmount(); }); };
  return { container, keys, clickAt, modelVersion, unmount };
}

describe("caret-move repaint: first edit after a caret jump takes the full-refresh path", () => {
  it("typing after clicking a different line bumps modelVersion; same-spot typing does not", async () => {
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "caret-move", "alice");
    // Build three lines. (Line y-bands in this harness: line1<=84, line2~120, line3>=200.)
    await ed.clickAt(12, 12);
    await ed.keys([..."aaaa"]);
    await ed.keys(["Enter"]); await ed.keys([..."bbbb"]);
    await ed.keys(["Enter"]); await ed.keys([..."cccc"]);
    await settle();

    // Caret is in line 3. Click BACK into line 1 (a different text node), type one
    // char: the fix forces the full refresh() path, which bumps modelVersion.
    await ed.clickAt(12, 12);
    const beforeJumpEdit = ed.modelVersion();
    await ed.keys([..."X"]);
    await settle();
    const afterJumpEdit = ed.modelVersion();
    expect(afterJumpEdit, "first edit after a caret jump must bump modelVersion (refresh path)").toBeGreaterThan(beforeJumpEdit);

    // Control: keep typing in the SAME spot — the fast in-place path is retained,
    // so at least one subsequent char must NOT bump modelVersion (no needless
    // full relayout on every keystroke).
    const versions: number[] = [];
    for (const c of [..."YZW"]) {
      await ed.keys([c]);
      await settle(6);
      versions.push(ed.modelVersion());
    }
    const bumpedEveryChar = versions.every((v, i) => v > (i === 0 ? afterJumpEdit : versions[i - 1]));
    expect(bumpedEveryChar, "continuous same-spot typing must keep the fast path (not refresh every char)").toBe(false);

    await ed.unmount();
  });
});
