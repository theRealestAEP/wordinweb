// @vitest-environment jsdom
/**
 * Browser-fidelity tests for the demo's feature surface BEYOND plain typing:
 * Enter (paragraph split), Backspace (delete), select-all + Ctrl+B (whole-run
 * bold), shift-arrow sub-range + Ctrl+B (formatRange with carried ids),
 * TWO live editors converging bidirectionally, and presence carets. Every
 * interaction is driven the way a browser delivers it — real mousedown /
 * keydown events through the editor's own handlers — over a delayed (async)
 * transport so echoes are in flight like a real network. Companion to
 * demo-fidelity.test.tsx (mount parity + typing).
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";
import { serializeXml } from "@wordinweb/core";

const provider: DocProvider = { load: () => blankDocxBytes() };

let factorySeq = 0; // unique across factories: two editors in one test must
// NOT collide on transport connection ids (the hub's presence fan-out filters
// by conn.id — a collision silently ate every presence update).
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `f${factorySeq++}-`;
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

/** Server-side state via a spy client: text per paragraph + document XML. */
async function spyState(factory: (u: string) => WebSocket, doc: string) {
  const spy = new CollabConnection(createWebSocketTransport(factory("ws://spy") as never), "spy");
  spy.join(doc);
  for (let i = 0; i < 40 && !spy.doc; i++) await new Promise((r) => setTimeout(r, 5));
  expect(spy.doc).toBeTruthy();
  const body = spy.doc!.docRoot.children.find((c) => c.name.endsWith("body"))!;
  const textOf = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(textOf).join("");
  const paras = body.children.filter((c) => c.name.endsWith(":p"));
  return {
    paragraphs: paras.length,
    text: paras.map(textOf).join("|"),
    xml: serializeXml(spy.doc!.docRoot),
  };
}

/** Mount one CollabEditor; returns its container + input helpers. */
async function mountEditor(hub: CollabHub, docId: string, clientId: string) {
  const factory = factoryFor(hub);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(CollabEditor, { url: "ws://x", docId, clientId, createSocket: factory }));
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
  const keys = async (seq: (string | { key: string; ctrl?: boolean; shift?: boolean })[]) => {
    // Target the container-scoped focus target so two mounted editors don't
    // cross wires: the editor's textarea lives inside its own container.
    const target = () =>
      (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    await act(async () => {
      for (const s of seq) {
        const k = typeof s === "string" ? { key: s } : s;
        target().dispatchEvent(new KeyboardEvent("keydown", {
          key: k.key, ctrlKey: !!k.ctrl, shiftKey: !!k.shift, bubbles: true, cancelable: true,
        }));
        await new Promise((r) => setTimeout(r, 2));
      }
    });
  };
  const typed = async (text: string) => keys([...text]);
  const painted = () => (container.querySelector(".dxw-page")!.textContent ?? "").replace(/​/g, "");
  /** The client replica's live document.xml, reached through the React tree —
   * the strongest convergence check: byte-equal to the server's. */
  const clientXml = (): string => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 5000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { doc?: { docRoot: unknown } } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const d = f.memoizedProps?.collab?.doc;
      if (d) return serializeXml(d.docRoot as never);
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    throw new Error("collab doc not found in React tree");
  };
  const unmount = async () => { await act(async () => { root.unmount(); }); };
  return { container, factory, click, keys, typed, painted, clientXml, unmount };
}

describe("demo features: paragraph split and delete through real keys", () => {
  it("Enter splits a paragraph; both sides + server agree", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "split-doc", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys(["Enter"]);
    await ed.typed("c");
    await settle();

    const server = await spyState(ed.factory, "split-doc");
    expect(server.paragraphs).toBe(2);
    expect(server.text).toBe("ab|c");
    expect(ed.painted()).toBe("abc"); // two painted lines concatenate to abc
    await ed.unmount();
  });

  it("Backspace deletes and converges", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "del-doc", "alice");
    await ed.click();
    await ed.typed("abcd");
    await ed.keys(["Backspace", "Backspace"]);
    await settle();

    const server = await spyState(ed.factory, "del-doc");
    expect(server.text).toBe("ab");
    expect(ed.painted()).toBe("ab");
    await ed.unmount();
  });
});

describe("demo features: formatting through real keyboard shortcuts", () => {
  it("Ctrl+A then Ctrl+B bolds the whole run on the server", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "bold-doc", "alice");
    await ed.click();
    await ed.typed("boldme");
    await ed.keys([{ key: "a", ctrl: true }, { key: "b", ctrl: true }]);
    await settle();

    const server = await spyState(ed.factory, "bold-doc");
    expect(server.text).toBe("boldme"); // formatting never mutates text
    expect(server.xml).toContain("<w:b"); // the bold reached the server
    expect(ed.painted()).toBe("boldme");
    await ed.unmount();
  });

  it("shift+ArrowLeft selection + Ctrl+B formats; client XML byte-equals the server's", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "range-doc", "alice");
    await ed.click();
    await ed.typed("plainBOLD");
    // Select trailing chars, then bold them. (jsdom's zero-geometry hit
    // testing can widen the selection vs a real browser, so the invariant
    // asserted is the one that matters everywhere: whatever the editor DID —
    // formatRun or formatRange — the client and server converge to the SAME
    // bytes. The double-apply and self-transform bugs both broke exactly this.)
    await ed.keys([
      { key: "ArrowLeft", shift: true }, { key: "ArrowLeft", shift: true },
      { key: "ArrowLeft", shift: true }, { key: "ArrowLeft", shift: true },
      { key: "b", ctrl: true },
    ]);
    await settle();

    const server = await spyState(ed.factory, "range-doc");
    expect(server.text).toBe("plainBOLD");
    expect(server.xml).toContain("<w:b"); // the bold reached the server
    expect(ed.painted()).toBe("plainBOLD");
    expect(ed.clientXml()).toBe(server.xml); // byte-identical convergence
    await ed.unmount();
  });
});

describe("demo features: two live editors, bidirectional", () => {
  it("edits from both editors interleave and converge everywhere", async () => {
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "duo-doc", "alice");
    const b = await mountEditor(hub, "duo-doc", "bob");
    await settle();

    // Alice types; Bob receives it.
    await a.click();
    await a.typed("hi ");
    await settle();
    expect(b.painted()).toBe("hi ");

    // Bob clicks (his caret lands via the same hit-test) and appends.
    await b.click();
    await b.keys(["End"]);
    await b.typed("bob");
    await settle();

    const server = await spyState(a.factory, "duo-doc");
    expect(server.text).toBe("hi bob");
    expect(a.painted()).toBe("hi bob"); // Alice sees Bob's edit
    expect(b.painted()).toBe("hi bob");
    await a.unmount();
    await b.unmount();
  });
});

describe("demo features: presence carets flow editor-to-editor", () => {
  it("Alice's caret placement draws a remote caret in Bob's view", async () => {
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "pres-doc", "alice");
    const b = await mountEditor(hub, "pres-doc", "bob");
    await settle();

    // Give the doc real text first (presence geometry needs a text position),
    // then click: Alice's editor reports the caret via onCaretMove →
    // setPresence → hub → Bob draws a colored caret bar over his page.
    await a.click();
    await a.typed("shared text");
    await settle();
    await a.click(); // a fresh caret placement after the text exists
    await settle();

    const remoteCaret = b.container.querySelector<HTMLElement>(".dxw-presence-caret");
    expect(remoteCaret).toBeTruthy();
    expect(remoteCaret!.dataset.participant).toBeTruthy();
    // And it is not drawn in Alice's own view (presence excludes the sender).
    expect(a.container.querySelector(".dxw-presence-caret")).toBeNull();
    await a.unmount();
    await b.unmount();
  });
});

describe("demo features: toolbar API ops route through the canonical apply", () => {
  /** Mount CollabEditor WITH its api observable (toolbar disabled to keep the
   * DOM minimal — the api methods are exactly what the toolbar buttons call). */
  async function mountWithApi(hub: CollabHub, docId: string, clientId: string) {
    const factory = factoryFor(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    let api: import("../src/index.js").DocxViewApi | null = null;
    await act(async () => {
      root.render(createElement(CollabEditor, {
        url: "ws://x", docId, clientId, createSocket: factory,
        toolbar: false, onReady: (a) => { api = a; },
      }));
    });
    for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
    expect(api).toBeTruthy();
    const clientXml = (): string => {
      const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
      const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
      let guard = 0;
      while (stack.length && guard++ < 5000) {
        const f = stack.pop() as { memoizedProps?: { collab?: { doc?: { docRoot: unknown } } }; child?: unknown; sibling?: unknown } | null;
        if (!f) continue;
        const d = f.memoizedProps?.collab?.doc;
        if (d) return serializeXml(d.docRoot as never);
        if (f.child) stack.push(f.child);
        if (f.sibling) stack.push(f.sibling);
      }
      throw new Error("no collab doc");
    };
    const click = async () => {
      const page = container.querySelector<HTMLElement>(".dxw-page")!;
      const span = page.querySelector("span") ?? page;
      await act(async () => {
        const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
        span.dispatchEvent(new MouseEvent("mousedown", o));
        span.dispatchEvent(new MouseEvent("mouseup", o));
      });
      await tick();
    };
    return { api: api as unknown as import("../src/index.js").DocxViewApi, factory, container, root, clientXml, click,
      unmount: async () => { await act(async () => { root.unmount(); }); } };
  }

  it("insertTable / insertEquation / setParagraphSpacing converge byte-identically", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountWithApi(hub, "api-doc", "alice");
    await ed.click(); // place a caret (the ops anchor at it)
    await act(async () => {
      expect(ed.api.insertEquation("a+b")).toBe(true);
      ed.api.insertTable(2, 2);
      ed.api.setParagraphSpacing({ beforePt: 12 });
    });
    await settle();

    const server = await spyState(ed.factory, "api-doc");
    expect(server.xml).toContain("<w:tbl");   // table landed on the server
    expect(server.xml).toContain("oMath");    // equation landed
    expect(server.xml).toContain("w:spacing"); // spacing landed
    expect(ed.clientXml()).toBe(server.xml);  // and the client is byte-identical
    await ed.unmount();
  });

  it("a second editor receives toolbar-inserted content", async () => {
    const hub = new CollabHub(provider);
    const a = await mountWithApi(hub, "api-duo", "alice");
    const b = await mountWithApi(hub, "api-duo", "bob");
    await settle();
    await a.click();
    await act(async () => { a.api.insertTable(2, 2); });
    await settle();
    // Bob's replica applied the broadcast table; both docs byte-identical.
    expect(b.clientXml()).toContain("<w:tbl");
    expect(b.clientXml()).toBe(a.clientXml());
    await a.unmount();
    await b.unmount();
  });
});
