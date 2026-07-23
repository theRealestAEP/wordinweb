// @vitest-environment jsdom
/**
 * Browser-fidelity tests for the DEMO FLOW: a user opens a blank collab doc in
 * CollabEditor, clicks the page, and types. This is exactly what the
 * examples/anon-share demo does — and what the intent/convergence suites do
 * NOT cover (they submit intents programmatically, never through the editor's
 * real mouse/keyboard path). These tests drive the same DOM events a browser
 * delivers and assert the two invariants the demo broke:
 *
 *  1. PARITY: the client replica must never mutate its doc outside the intent
 *     loop (the client's paragraph count must match the server's).
 *  2. TYPING WORKS: click → caret → keydown inserts text locally AND the
 *     intent reaches the hub (other clients converge).
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";

// The REAL blank doc the demo serves (packages/server/src/blank.ts).
const provider: DocProvider = { load: () => blankDocxBytes() };

function factoryFor(hub: CollabHub, opts: { delayMs?: number } = {}) {
  let n = 0;
  const delay = opts.delayMs ?? 0;
  const defer = (fn: () => void) => (delay > 0 ? setTimeout(fn, delay) : fn());
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    // With delayMs > 0 this behaves like a real network: submits and echoes
    // each take a hop, so a burst of keystrokes has MANY intents in flight
    // before the first echo returns. The synchronous loopback hid the
    // multiple-in-flight transform bug that broke burst typing in the browser.
    const conn: Connection = { id: `c${n++}`, send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))) };
    let opened = false;
    return { send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

/** Paragraph count + concatenated text of a hub-side doc, via a spy client. */
async function spyState(hub: CollabHub, factory: (u: string) => WebSocket, doc: string) {
  const spy = new CollabConnection(createWebSocketTransport(factory("ws://spy") as never), "spy");
  spy.join(doc);
  // The hub handles the hello asynchronously; wait for the welcome.
  for (let i = 0; i < 40 && !spy.doc; i++) await new Promise((r) => setTimeout(r, 5));
  expect(spy.doc).toBeTruthy();
  const body = spy.doc!.docRoot.children.find((c) => c.name.endsWith("body"))!;
  const textOf = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(textOf).join("");
  const paras = body.children.filter((c) => c.name.endsWith(":p"));
  return { paragraphs: paras.length, text: paras.map(textOf).join("|") };
}

/** Mount the demo editor on a blank doc; returns DOM + accessors. */
async function mountDemo(docId: string, opts: { delayMs?: number } = {}) {
  const hub = new CollabHub(provider);
  const factory = factoryFor(hub, opts);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(CollabEditor, { url: "ws://x", docId, clientId: "typist", createSocket: factory }));
  });
  for (let i = 0; i < 20 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  return { hub, factory, container, root };
}

/** The client-side doc as rendered — reach it via the debug hook DocxView sets
 * (data flows: replica.doc === collab.doc rendered live). We read paragraph
 * spans painted in the page instead of internals. */
function paintedLineCount(container: HTMLElement): number {
  const surface = container.querySelector(".dxw-page")?.firstElementChild;
  return surface ? surface.querySelectorAll(":scope > span[data-dxw-item-kind=\"text\"], :scope > span").length : 0;
}

describe("demo fidelity: blank doc parity (the divergence bug)", () => {
  it("mounting the editor must not grow the client doc beyond the server's", async () => {
    const { hub, factory, container, root } = await mountDemo("parity-doc");
    // Let any post-mount effects (renderSignal repaint, editor attach) settle.
    for (let i = 0; i < 10; i++) await tick();

    const server = await spyState(hub, factory, "parity-doc");
    expect(server.paragraphs).toBe(1); // blank doc: exactly one paragraph

    // THE BUG: the editor's attach() locally seeded paragraphs the server
    // never saw (client 4 paras vs server 1 in the live demo). The painted
    // page must show exactly the server's single (empty) paragraph line.
    expect(paintedLineCount(container)).toBeLessThanOrEqual(1 + 0); // one line for one empty paragraph
    await act(async () => { root.unmount(); });
  });
});

describe("demo fidelity: click + type on a blank doc (the dead-typing bug)", () => {
  it("mousedown on the page then keydown types text that reaches the server", async () => {
    const { hub, factory, container, root } = await mountDemo("typing-doc");
    for (let i = 0; i < 10; i++) await tick();

    // 1. Click the page like a mouse would (mousedown/mouseup/click bubble
    //    from the painted surface to the editor's container listeners).
    const page = container.querySelector<HTMLElement>(".dxw-page")!;
    const surfaceSpan = page.querySelector("span") ?? page;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      surfaceSpan.dispatchEvent(new MouseEvent("mousedown", opts));
      surfaceSpan.dispatchEvent(new MouseEvent("mouseup", opts));
      surfaceSpan.dispatchEvent(new MouseEvent("click", opts));
    });
    await tick();

    // 2. Type like a keyboard would: keydown events on the focused element
    //    (the editor's hidden textarea) bubbling through the container.
    const target = (document.activeElement as HTMLElement) ?? container;
    await act(async () => {
      for (const key of "Hi!") {
        target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
      }
    });
    for (let i = 0; i < 10; i++) await tick();

    // 3. The text must exist on the SERVER (intent loop ran) — not merely
    //    locally. This is the invariant the broken demo failed: keystrokes
    //    were swallowed with no caret and nothing ever reached the hub.
    const server = await spyState(hub, factory, "typing-doc");
    expect(server.text).toBe("Hi!");

    // 4. And the CLIENT must equal the server EXACTLY. A "contains" check
    //    passed while the live demo showed "Hello collab!" + its own reversal:
    //    the connection re-applied optimistically what the editor had already
    //    applied to the live doc (double-apply). Painted text is the client
    //    truth the user sees.
    const painted = container.querySelector(".dxw-page")!.textContent ?? "";
    expect(painted.replace(/​/g, "")).toBe("Hi!");
    await act(async () => { root.unmount(); });
  });

  it("BURST typing over a delayed (real-network-like) transport converges exactly", async () => {
    // delayMs makes submits/echoes each take a hop, so every keystroke after
    // the first is in flight before any echo returns — the browser reality.
    // Before the fix, the server transformed each keystroke against the SAME
    // client's own in-flight inserts (double-counting them): the first char
    // applied, the rest were rejected, and the client silently kept its
    // rejected optimistic text ("Hello collab" + its reversal on screen while
    // the server held just "H").
    const { hub, factory, container, root } = await mountDemo("burst-doc", { delayMs: 2 });
    for (let i = 0; i < 10; i++) await tick();

    const page = container.querySelector<HTMLElement>(".dxw-page")!;
    const surfaceSpan = page.querySelector("span") ?? page;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      surfaceSpan.dispatchEvent(new MouseEvent("mousedown", opts));
      surfaceSpan.dispatchEvent(new MouseEvent("mouseup", opts));
    });
    await tick();

    const target = (document.activeElement as HTMLElement) ?? container;
    const typed = "The demo works";
    await act(async () => {
      for (const key of typed) {
        target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
      }
    });
    // Drain all delayed hops (submits, echoes, broadcasts).
    for (let i = 0; i < 30; i++) await tick();

    const server = await spyState(hub, factory, "burst-doc");
    expect(server.text).toBe(typed); // every keystroke accepted, in order
    const painted = container.querySelector(".dxw-page")!.textContent ?? "";
    expect(painted.replace(/​/g, "")).toBe(typed); // client shows exactly the canon
    await act(async () => { root.unmount(); });
  });
});
