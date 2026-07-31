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
 * pins the FIX'S MECHANISM instead. The fix is a SCOPED reparse of the caret's
 * paragraph: fresh model objects make pageEq fail for exactly the touched
 * page, so its DOM is rebuilt — page adoption is defeated without rebuilding
 * the whole model.
 *
 * The first revision of the fix forced a full doc.refresh() here instead.
 * That was the 500-page editing regression (scripts/bench-local-typing.mjs):
 * refresh() rebuilt every model object (whole-document re-measure, ~96k
 * measureText calls on a 3000-paragraph doc vs ~34 for one paragraph) and its
 * modelVersion bump sent rerender down the >50-page background-relayout path
 * with the container inert — a multi-second input-eating stall after every
 * click-then-type. So this test pins BOTH halves:
 *   1. the caret paragraph's retained model object is REPLACED by the first
 *      edit after a jump (the adoption-defeat that repaints the glyph), and
 *   2. modelVersion does NOT change (no O(document) refresh/relayout).
 * Verified end-to-end in a real browser; see BUGS.md.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import type { DocxDocument, Paragraph, XmlElement } from "@wordinweb/core";

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
  // Reach the live collab doc through the React tree.
  const doc = (): DocxDocument => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 5000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { doc?: DocxDocument } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const d = f.memoizedProps?.collab?.doc;
      if (d && typeof d.modelVersion === "number") return d;
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    throw new Error("collab doc not found");
  };
  const unmount = async () => { await act(async () => { root.unmount(); }); };
  return { container, keys, clickAt, doc, unmount };
}

describe("caret-move repaint: first edit after a caret jump reparses the caret paragraph", () => {
  it("replaces the paragraph model object (adoption defeat) WITHOUT a whole-document refresh", async () => {
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "caret-move", "alice");
    // Build three lines. (Line y-bands in this harness: line1<=84, line2~120, line3>=200.)
    await ed.clickAt(12, 12);
    await ed.keys([..."aaaa"]);
    await ed.keys(["Enter"]); await ed.keys([..."bbbb"]);
    await ed.keys(["Enter"]); await ed.keys([..."cccc"]);
    await settle();

    // Caret is in line 3. Click BACK into line 1 (a different text node) and
    // type one char. The paragraph model holding the typed text must be a
    // FRESH object afterwards — identical retained objects are exactly what
    // page adoption compares equal, leaving the glyph unpainted in a real
    // browser (the click-then-type bug).
    await ed.clickAt(12, 12);
    const d = ed.doc();
    const holds = (p: Paragraph, text: string): boolean =>
      p.children.some((r) => "content" in r && (r as { content: { kind: string; text?: string }[] }).content.some(
        (c) => c.kind === "text" && (c.text ?? "").includes(text)));
    const paraWith = (text: string): Paragraph => {
      for (const s of d.sections) {
        for (const b of s.blocks) if (b.type === "paragraph" && holds(b, text)) return b;
      }
      throw new Error(`no paragraph containing "${text}"`);
    };
    const beforeObj = paraWith("aaaa");
    const beforeSrc = beforeObj.src as XmlElement;
    const beforeVersion = d.modelVersion;
    await ed.keys([..."X"]);
    await settle();

    // 1. Adoption defeat: the retained model object was replaced (same XML
    //    element underneath — the reparse is scoped, not a rebuild).
    const afterObj = paraWith("aaa"); // the X landed somewhere inside the aaaa line
    expect(afterObj, "first edit after a caret jump must produce a FRESH paragraph model object").not.toBe(beforeObj);
    expect(afterObj.src, "the reparse must retain the same XML source element").toBe(beforeSrc);
    expect(holds(afterObj, "X"), "the typed char must be in the reparsed model").toBe(true);

    // 2. No whole-document work: modelVersion must NOT move. A bump here means
    //    doc.refresh() ran — the O(document) model rebuild whose >50-page
    //    rerender answer is an async global relayout behind an inert container
    //    (the 500-page click-then-type stall this pins against).
    expect(d.modelVersion, "a caret-jump edit must not take the O(document) refresh path").toBe(beforeVersion);

    // Control: continuous typing in the SAME spot keeps the fast in-place path
    // (no refresh, and no reparse churn is observable beyond the model staying
    // paint-correct — the glyphs land).
    await ed.keys([..."YZ"]);
    await settle(6);
    expect(d.modelVersion, "same-spot typing must keep the fast path").toBe(beforeVersion);
    expect(ed.container.textContent).toContain("Y");

    await ed.unmount();
  });
});
