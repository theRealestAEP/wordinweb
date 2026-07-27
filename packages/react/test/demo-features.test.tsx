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

let spySeq = 0; // fresh identity per spy: a reused clientId is refused `already-open` (doc 12 §7)
/** Server-side state via a spy client: text per paragraph + document XML. */
async function spyState(factory: (u: string) => WebSocket, doc: string) {
  const spy = new CollabConnection(createWebSocketTransport(factory("ws://spy") as never), `spy-${spySeq++}`);
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
    // EVERY editable part, not just document.xml: header/footer content lives
    // in its own parts, so a document.xml-only oracle is blind to an hf edit
    // that never rode the wire.
    allXml: allRootsXml(spy.doc!),
  };
}

/** Serialize every editable root (document, settings, and each header/footer
 * part) as one string — the parity oracle for edits outside document.xml. */
function allRootsXml(doc: { editableRoots(): { name: string }[] }): string {
  return doc.editableRoots().map((r) => serializeXml(r as never)).join("\n----\n");
}

/** Mount one CollabEditor; returns its container + input helpers.
 *
 * EMISSION GUARD (always-on): `unmount()` asserts the client's document is
 * BYTE-IDENTICAL to the server's before tearing down. A local mutation that
 * never emitted an intent produces no wire traffic — every wire-level assert
 * stays green while the room silently forks — so parity at teardown is the
 * one check that catches the whole class automatically in every test that
 * uses this harness. Tests that intentionally leave divergence (none today)
 * can pass { skipParity: true }. */
async function mountEditor(hub: CollabHub, docId: string, clientId: string) {
  const factory = factoryFor(hub);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let api: import("../src/index.js").DocxViewApi | null = null;
  await act(async () => {
    root.render(createElement(CollabEditor, {
      url: "ws://x", docId, clientId, createSocket: factory, onReady: (a) => { api = a; },
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
  const clientDoc = (): { docRoot: unknown; editableRoots(): { name: string }[] } => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 5000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { doc?: never } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const d = f.memoizedProps?.collab?.doc;
      if (d) return d;
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    throw new Error("collab doc not found in React tree");
  };
  const clientXml = (): string => serializeXml(clientDoc().docRoot as never);
  /** Every editable part of the client's replica (see allRootsXml). */
  const clientAllXml = (): string => allRootsXml(clientDoc());
  const unmount = async (opts?: { skipParity?: boolean }) => {
    if (!opts?.skipParity) {
      await settle();
      const server = await spyState(factory, docId);
      expect(clientXml(), "EMISSION GUARD: client diverged from server — some local mutation never emitted an intent").toBe(server.xml);
      // Same guard, widened past document.xml: a header/footer edit that never
      // emitted leaves document.xml identical and the hf part forked.
      expect(clientAllXml(), "EMISSION GUARD: client diverged from server outside document.xml (header/footer part)").toBe(server.allXml);
    }
    await act(async () => { root.unmount(); });
  };
  return {
    container, factory, click, keys, typed, painted, clientDoc, clientXml, clientAllXml, unmount,
    api: () => api as unknown as import("../src/index.js").DocxViewApi,
  };
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

  it("Ctrl+A + Ctrl+B over a MULTI-w:t run (inline tab) bolds the WHOLE run and converges", async () => {
    // Offset-basis pin (checkpoint B1 audit): a run holding several w:t
    // yields one selection segment PER w:t, all sharing the runId. The
    // format emitter must recognize the segments' COMBINED cumulative span
    // as whole-run and emit formatRun — treating the first segment alone as
    // the selection emits a partial formatRange that silently drops the
    // later w:t's bold on the wire (local bolds everything, server doesn't:
    // byte-divergence, exactly what the clientXml === server.xml pin catches).
    const tabProvider: DocProvider = {
      load: () => {
        const { zipSync, strToU8 } = require("fflate") as typeof import("fflate");
        const body = `<w:p><w:r><w:t xml:space="preserve">HELLO</w:t><w:tab/><w:t xml:space="preserve">WORLD</w:t></w:r></w:p>`;
        return zipSync({
          "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
          "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
          "word/document.xml": strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`),
        });
      },
    };
    const hub = new CollabHub(tabProvider);
    const ed = await mountEditor(hub, "tab-bold-doc", "alice");
    await ed.click();
    await ed.keys([{ key: "a", ctrl: true }, { key: "b", ctrl: true }]);
    await settle();

    const server = await spyState(ed.factory, "tab-bold-doc");
    expect(server.text).toBe("HELLOWORLD"); // formatting never mutates text
    expect(server.xml).toContain("<w:b"); // the bold reached the server
    expect(ed.clientXml()).toBe(server.xml); // byte-identical convergence
    await ed.unmount();
  });
});

describe("demo features: selection deletes ride the wire (review finding)", () => {
  it("shift-select + Backspace reaches the server (was a silent local-only mutation)", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "selbksp-doc", "alice");
    await ed.click();
    await ed.typed("abcdef");
    await ed.keys([
      { key: "ArrowLeft", shift: true }, { key: "ArrowLeft", shift: true },
      { key: "ArrowLeft", shift: true },
      "Backspace",
    ]);
    await settle();

    const server = await spyState(ed.factory, "selbksp-doc");
    expect(server.text).toBe("abc"); // the selected "def" is gone SERVER-side
    expect(ed.painted()).toBe("abc");
    expect(ed.clientXml()).toBe(server.xml); // byte-identical convergence
    await ed.unmount();
  });

  it("typing over a selection emits the delete AND the insert", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "seltype-doc", "alice");
    await ed.click();
    await ed.typed("abcdef");
    await ed.keys([
      { key: "ArrowLeft", shift: true }, { key: "ArrowLeft", shift: true },
      { key: "ArrowLeft", shift: true },
    ]);
    await ed.typed("XY");
    await settle();

    const server = await spyState(ed.factory, "seltype-doc");
    expect(server.text).toBe("abcXY"); // "def" replaced by "XY" server-side
    expect(ed.painted()).toBe("abcXY");
    expect(ed.clientXml()).toBe(server.xml);
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

describe("demo features: paragraph merges + Enter fast paths (the desync bugs)", () => {
  // Each of these was a SILENT LOCAL-ONLY mutation before the fix: the editor
  // took a fast path that never emitted an intent, so the client's structure
  // diverged from every other participant ("hitting Enter and typing more
  // don't work"). Byte-identical client==server is the assertion throughout.
  async function editorOn(hub: CollabHub, doc: string) {
    const ed = await mountEditor(hub, doc, "alice");
    await ed.click();
    return ed;
  }
  const xmlEq = async (ed: Awaited<ReturnType<typeof mountEditor>>, doc: string, hub: CollabHub) => {
    const server = await spyState(ed.factory, doc);
    expect(ed.clientXml()).toBe(server.xml);
    return server;
  };

  it("Backspace at paragraph start merges and converges", async () => {
    const hub = new CollabHub(provider);
    const ed = await editorOn(hub, "merge-back");
    await ed.typed("ab");
    await ed.keys(["Enter"]);
    await ed.typed("cd");
    await settle();
    // Caret is at end of "cd": Home to reach paragraph start, then Backspace merges.
    await ed.keys(["Home", "Backspace"]);
    await settle();
    const server = await xmlEq(ed, "merge-back", hub);
    expect(server.paragraphs).toBe(1);
    expect(server.text).toBe("abcd");
    await ed.unmount();
  });

  it("Delete at paragraph end merges the next paragraph and converges", async () => {
    const hub = new CollabHub(provider);
    const ed = await editorOn(hub, "merge-fwd");
    await ed.typed("ab");
    await ed.keys(["Enter"]);
    await ed.typed("cd");
    await settle();
    // Move to end of first paragraph structurally: Home (start of "cd") then
    // ArrowLeft steps across the paragraph boundary to the end of "ab"
    // (vertical arrows need real geometry, which jsdom lacks). Delete merges.
    await ed.keys(["Home", "ArrowLeft", "Delete"]);
    await settle();
    const server = await xmlEq(ed, "merge-fwd", hub);
    expect(server.paragraphs).toBe(1);
    expect(server.text).toBe("abcd");
    await ed.unmount();
  });

  it("Enter at paragraph START converges (the blank-before fast path bug)", async () => {
    const hub = new CollabHub(provider);
    const ed = await editorOn(hub, "enter-start");
    await ed.typed("xy");
    await settle();
    await ed.keys(["Home", "Enter"]);
    await settle();
    const server = await xmlEq(ed, "enter-start", hub);
    expect(server.paragraphs).toBe(2);
    expect(server.text).toBe("|xy");
    // And typing afterwards still works + reaches the server (the user's bug:
    // everything after this Enter went silently local-only).
    await ed.typed("z");
    await settle();
    const after = await spyState(ed.factory, "enter-start");
    expect(after.text.includes("z")).toBe(true);
    expect(ed.clientXml()).toBe(after.xml);
    await ed.unmount();
  });

  it("Enter on an EMPTY list item exits the list and converges", async () => {
    const hub = new CollabHub(provider);
    const ed = await editorOn(hub, "list-exit");
    await ed.typed("item");
    await settle();
    // Make it a bullet via Ctrl+Shift+L (the editor's list shortcut), add an
    // empty item, then Enter on the empty item exits the list.
    await ed.keys([{ key: "l", ctrl: true, shift: true } as never]);
    await settle();
    await ed.keys(["End", "Enter"]); // new empty list item
    await settle();
    await ed.keys(["Enter"]); // exit list (setListType null — was local-only)
    await settle();
    const server = await xmlEq(ed, "list-exit", hub);
    // The exited paragraph carries no numbering; the first still does.
    expect(server.xml).toContain("numPr");
    await ed.unmount();
  });

  it("CONCURRENT Enter from two editors converges (split-id collision regression)", async () => {
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "dual-enter", "alice");
    const b = await mountEditor(hub, "dual-enter", "bob");
    await settle();
    await a.click();
    await a.typed("aa");
    await settle(20);
    await b.click();
    // Both split in the same tick — before either echo lands. Pre-fix, both
    // editors minted the SAME carried ids from their local counters and the
    // second split was rejected on every replica (dead typing after Enter).
    await act(async () => {
      (a.container.querySelector("textarea") as HTMLElement).dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
      (b.container.querySelector("textarea") as HTMLElement).dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    });
    await settle(40);
    await a.typed("z");
    await settle(30);
    expect(a.clientXml()).toBe(b.clientXml());
    const server = await spyState(a.factory, "dual-enter");
    expect(a.clientXml()).toBe(server.xml);
    await a.unmount();
    await b.unmount();
  });
});

describe("caret survives a true-conflict reload (docEpoch remount)", () => {
  /** Mount with the api AND session observable plus real-key input — the
   * remount under test re-fires onReady with a fresh api, so `seen` always
   * holds the CURRENT view's handles. */
  async function mountLive(hub: CollabHub, docId: string, clientId: string) {
    const factory = factoryFor(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    const seen: {
      api: import("../src/index.js").DocxViewApi | null;
      session: import("../src/collab.js").CollabSession | null;
    } = { api: null, session: null };
    await act(async () => {
      root.render(createElement(CollabEditor, {
        url: "ws://x", docId, clientId, createSocket: factory, toolbar: false,
        onReady: (a) => { seen.api = a; },
        onSession: (s) => { seen.session = s; },
      }));
    });
    for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
    expect(container.querySelector(".dxw-page")).toBeTruthy();
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
    // Raw dispatch (no act wrapper) so two editors' bursts can interleave
    // inside ONE act() — overlapping act() calls are unsupported.
    const rawKey = (key: string) => {
      const target =
        (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
      target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
    };
    const keys = async (seq: string[]) => {
      await act(async () => {
        for (const key of seq) {
          rawKey(key);
          await new Promise((r) => setTimeout(r, 2));
        }
      });
    };
    return { container, factory, seen, click, keys, rawKey,
      unmount: async () => { await act(async () => { root.unmount(); }); } };
  }

  it("a remote edit landing mid-burst reloads the doc but keeps the local caret", async () => {
    const hub = new CollabHub(provider);
    const a = await mountLive(hub, "caret-doc", "alice");
    const b = await mountLive(hub, "caret-doc", "bob");
    await settle();

    // Seed text and settle so both carets sit in real content.
    await a.click();
    await a.keys([..."base "]);
    await settle();
    await b.click();
    await b.keys(["End"]);

    // Interleave: both type at once over the delayed wire, so remote
    // broadcasts land while local echoes are still in flight — the replica's
    // slow path (restore + replay), which reloads the doc object.
    await act(async () => {
      for (let i = 0; i < 6; i++) {
        a.rawKey("a");
        await new Promise((r) => setTimeout(r, 2));
        b.rawKey("b");
        await new Promise((r) => setTimeout(r, 2));
      }
    });
    await settle(40);

    // The reload really happened (the path under test): docEpoch advanced,
    // which re-keyed and REMOUNTED Alice's DocxView.
    expect(a.seen.session!.docEpoch).toBeGreaterThanOrEqual(1);
    // Her caret survived the remount: still encodable to a stable-id address
    // and still painted in the replacement view. Pre-fix, the remount lost it
    // (getEncodedCaret() === null, no caret bar).
    expect(a.seen.api!.getEncodedCaret()).not.toBeNull();
    const bar = a.container.querySelector<HTMLElement>("[data-dxw-caret]");
    expect(bar).toBeTruthy();
    expect(bar!.style.display).not.toBe("none");
    // And the reconciliation converged as always (byte-identical to server).
    const server = await spyState(a.factory, "caret-doc");
    expect(serializeXml(a.seen.session!.doc!.docRoot)).toBe(server.xml);
    await a.unmount();
    await b.unmount();
  });
});

/**
 * The inline equation editor, driven the way a user drives it: click the
 * rendered equation, retype its linear form, press Apply.
 *
 * This popover is why the class of bug it belongs to survived a green suite —
 * NOTHING in the suite ever opened it. Its three commit paths (Apply, Delete,
 * and the empty-input-deletes shortcut) all called the core mutation directly
 * with no emission, so an equation edit was a permanent silent fork: the
 * author saw the new equation, everyone else kept the old one, and no wire
 * traffic existed for any assertion to catch. The tests below drive the real
 * DOM the popover builds, so the emission guard in mountEditor covers them.
 */
describe("demo features: the inline math popover rides the wire", () => {
  /** The live popover's controls, located the way a user sees them. */
  function popover() {
    const input = Array.from(document.querySelectorAll<HTMLInputElement>("input"))
      .find((i) => (i.title ?? "").startsWith("Linear math"));
    if (!input) return null;
    const box = input.closest("div")?.parentElement as HTMLElement | undefined;
    if (!box) return null;
    const button = (label: string) =>
      Array.from(box.querySelectorAll("button")).find((b) => b.textContent === label) ?? null;
    return { input, box, apply: button("Apply"), remove: button("Delete") };
  }

  /** Seed an equation through the toolbar API, then click it open. */
  async function withEquation(hub: CollabHub, docId: string, clientId = "alice") {
    const ed = await mountEditor(hub, docId, clientId);
    await ed.click();
    await act(async () => { expect(ed.api().insertEquation("a+b")).toBe(true); });
    await settle();
    return ed;
  }
  async function openPopover(ed: Awaited<ReturnType<typeof mountEditor>>) {
    const mathEl = ed.container.querySelector<HTMLElement>("[data-dxw-math]");
    expect(mathEl, "the equation did not render a clickable math element").toBeTruthy();
    await act(async () => {
      mathEl!.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, clientX: 20, clientY: 20, button: 0 }));
    });
    await tick();
    const p = popover();
    expect(p, "clicking the equation did not open the math editor").toBeTruthy();
    return p!;
  }

  it("editing an equation's linear form and pressing Apply converges byte-identically", async () => {
    const hub = new CollabHub(provider);
    const ed = await withEquation(hub, "math-apply");
    const p = await openPopover(ed);

    // The popover opened on the real equation, showing its linear form.
    expect(p.input.value).toBe("a+b");
    expect(p.input.readOnly).toBe(false);
    // Retype it as a fraction — the exact shape the user reported losing.
    p.input.value = "{a+b}/{2}";
    await act(async () => { p.apply!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
    await settle();

    const server = await spyState(ed.factory, "math-apply");
    // The rewritten equation reached the SERVER, not just the local screen.
    expect(server.xml).toContain("m:f");     // the fraction wrapper
    expect(server.xml).toContain("m:num");
    expect(ed.clientXml()).toBe(server.xml); // byte-identical
    await ed.unmount();
  });

  it("a second editor receives an equation edit made through the popover", async () => {
    const hub = new CollabHub(provider);
    const a = await withEquation(hub, "math-duo");
    const b = await mountEditor(hub, "math-duo", "bob");
    await settle();
    expect(b.clientXml()).toContain("oMath");

    const p = await openPopover(a);
    p.input.value = "x^2";
    await act(async () => { p.apply!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
    await settle();

    // Bob's replica applied the broadcast: superscript present on BOTH sides.
    expect(b.clientXml()).toContain("m:sSup");
    expect(b.clientXml()).toBe(a.clientXml());
    await a.unmount();
    await b.unmount();
  });

  it("the popover's Delete removes the equation everywhere", async () => {
    const hub = new CollabHub(provider);
    const a = await withEquation(hub, "math-del");
    const b = await mountEditor(hub, "math-del", "bob");
    await settle();
    expect(a.clientXml()).toContain("oMath");

    const p = await openPopover(a);
    await act(async () => { p.remove!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
    await settle();

    const server = await spyState(a.factory, "math-del");
    expect(server.xml).not.toContain("oMath"); // the delete rode the wire
    expect(a.clientXml()).toBe(server.xml);
    expect(b.clientXml()).toBe(server.xml);
    await a.unmount();
    await b.unmount();
  });

  it("an unchanged Apply is a no-op (never rewrites the OMML)", async () => {
    const hub = new CollabHub(provider);
    const ed = await withEquation(hub, "math-noop");
    const before = ed.clientXml();
    const p = await openPopover(ed);
    await act(async () => { p.apply!.dispatchEvent(new MouseEvent("click", { bubbles: true })); });
    await settle();
    expect(ed.clientXml()).toBe(before);
    await ed.unmount();
  });
});

/**
 * Header/footer editing in a room.
 *
 * Two things had to be true for this to be safe, and neither was: header
 * paragraphs must be ADDRESSABLE by the ordinary text intents (they carry
 * stable ids, but the apply's run index walked only the body story, so every
 * hf intent was a clean reject on every peer while the author kept the edit),
 * and CREATING the part — a new part, relationship, content-type override and
 * sectPr reference — must ride the wire instead of happening locally. The
 * assertions below are on the whole package, not document.xml, because an hf
 * fork is invisible in document.xml.
 */
describe("demo features: header/footer editing rides the wire", () => {
  it("typing in the header reaches the server, byte-identically", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "hf-solo", "alice");
    await ed.click();
    // The real gesture the ribbon's Header button performs.
    await act(async () => { expect(ed.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await ed.typed("Acme");
    await settle();

    const server = await spyState(ed.factory, "hf-solo");
    // The header part exists on the SERVER and holds the typed text.
    expect(server.allXml).toContain("<w:hdr");
    expect(server.allXml).toContain("Acme");
    // …and document.xml alone would not have shown it — which is the point.
    expect(server.xml).not.toContain("Acme");
    expect(ed.clientAllXml()).toBe(server.allXml);
    await ed.unmount();
  });

  it("a second editor sees header text typed by the first", async () => {
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "hf-duo", "alice");
    const b = await mountEditor(hub, "hf-duo", "bob");
    await settle();
    await a.click();
    await act(async () => { expect(a.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await a.typed("Draft");
    await settle();

    expect(b.clientAllXml()).toContain("Draft");
    expect(b.clientAllXml()).toBe(a.clientAllXml());
    await a.unmount();
    await b.unmount();
  });

  it("a LATE joiner rehydrates the header (checkpoint/sidecar covers hf roots)", async () => {
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "hf-late", "alice");
    await a.click();
    await act(async () => { expect(a.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await a.typed("Late");
    await settle();

    // Joining only NOW: this replica never saw the ensureHeaderFooter intent
    // as a broadcast — it must arrive in the seed/checkpoint, with the hf
    // paragraph's stable ids reproduced from the sidecar.
    const b = await mountEditor(hub, "hf-late", "bob");
    await settle();
    expect(b.clientAllXml()).toContain("Late");
    expect(b.clientAllXml()).toBe(a.clientAllXml());

    // And the late joiner can go on editing the SAME header paragraph — proof
    // its ids resolve, not just that its bytes arrived.
    await act(async () => { expect(b.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await b.typed("r");
    await settle();
    expect(a.clientAllXml()).toContain("Later");
    expect(a.clientAllXml()).toBe(b.clientAllXml());
    await a.unmount();
    await b.unmount();
  });

  it("the hf hotbar's page-number field lands in the header on both sides", async () => {
    // The contextual hotbar's main action. It routes through the ordinary
    // insertPageField intent, so it only works because the apply's run index
    // now reaches header runs — the same fix typing depends on.
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "hf-pagenum", "alice");
    const b = await mountEditor(hub, "hf-pagenum", "bob");
    await settle();
    await a.click();
    await act(async () => { expect(a.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await act(async () => { expect(a.api().insertPageNumber("page")).toBe(true); });
    await settle();

    const server = await spyState(a.factory, "hf-pagenum");
    expect(server.allXml).toContain("PAGE");
    expect(a.clientAllXml()).toBe(server.allXml);
    expect(b.clientAllXml()).toBe(server.allXml);
    await a.unmount();
    await b.unmount();
  });

  it("a remote caret MOVES into the header (presence follows the peer there)", async () => {
    // Header runs are now resolvable ids, so a peer's caret in a header
    // reaches the presence receiver for the first time. The receiver is
    // purely binding-driven (runId → w:t → its rendered items → that item's
    // own surface) with no body-ancestry assumption anywhere, so this needed
    // no hf-specific work — but "it doesn't throw" would be a weak claim, so
    // this pins that the caret actually lands in the header BAND: its top
    // must rise above the body line it was on a moment earlier.
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "hf-presence", "alice");
    const b = await mountEditor(hub, "hf-presence", "bob");
    await settle();
    await a.click();
    await a.typed("body line");
    await settle();
    await a.click();
    await settle();

    const caretTop = () => {
      const el = b.container.querySelector<HTMLElement>(".dxw-presence-caret");
      return el ? parseFloat(el.style.top) : null;
    };
    const inBody = caretTop();
    expect(inBody, "no remote caret drawn for the body position").not.toBeNull();

    await act(async () => { expect(a.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await a.typed("Header");
    await settle();

    const inHeader = caretTop();
    expect(inHeader, "the remote caret vanished when the peer entered the header").not.toBeNull();
    expect(inHeader!).toBeLessThan(inBody!); // the header band sits above the body
    expect(b.container.querySelector(".dxw-presence-caret")!.closest(".dxw-page")).toBeTruthy();
    // Known cosmetic limitation, recorded rather than fixed: a header repeats
    // on every page, so its w:t has one binding per page and the receiver
    // picks the first — a remote header caret shows on page 1 only.
    await a.unmount();
    await b.unmount();
  });

  it("THE GUARD ITSELF: an unemitted header edit is caught (docRoot-only would miss it)", async () => {
    // The emission guard exists to catch FUTURE unemitted mutations, so
    // widening its oracle past document.xml is only worth anything if
    // something fails when the widening is removed. Nothing did — every hf
    // test above asserts allXml explicitly — so this is that pin.
    //
    // MUTATION CHECK: delete the `clientAllXml()` line from mountEditor's
    // unmount guard and this test is the ONLY one that goes red.
    const hub = new CollabHub(provider);
    const ed = await mountEditor(hub, "hf-guard", "alice");
    await ed.click();
    await act(async () => { expect(ed.api().openHeaderFooter("header")).toBe(true); });
    await settle();
    await ed.typed("Real");
    await settle();

    // Forge exactly the bug class the widened oracle is for: mutate the
    // header part locally, emitting nothing. This is what every hf edit did
    // before this arc.
    const hdr = ed.clientDoc().editableRoots().find((r) => r.name.endsWith("hdr")) as unknown as
      { name: string; text: string; children: { name: string; text: string; children: unknown[] }[] };
    expect(hdr, "no header part to fork").toBeTruthy();
    const firstT = (el: { name: string; text: string; children: unknown[] }): typeof el | null =>
      el.name.endsWith(":t") ? el : (el.children as (typeof el)[]).reduce<typeof el | null>((a, c) => a ?? firstT(c), null);
    firstT(hdr as never)!.text = "Forged";

    const server = await spyState(ed.factory, "hf-guard");
    // The OLD oracle is blind: document.xml is byte-identical on both sides…
    expect(ed.clientXml()).toBe(server.xml);
    // …while the widened one sees the fork.
    expect(ed.clientAllXml()).not.toBe(server.allXml);
    await expect(ed.unmount()).rejects.toThrow(/outside document\.xml/);
    await ed.unmount({ skipParity: true });
  });

  it("two participants opening the header at once create ONE part", async () => {
    const hub = new CollabHub(provider);
    const a = await mountEditor(hub, "hf-race", "alice");
    const b = await mountEditor(hub, "hf-race", "bob");
    await settle();
    await a.click();
    await b.click();
    // Both fire before either broadcast lands: the second ensureHeaderFooter
    // must apply as a clean no-op rather than minting a second header part.
    await act(async () => {
      a.api().openHeaderFooter("header");
      b.api().openHeaderFooter("header");
    });
    await settle();

    const server = await spyState(a.factory, "hf-race");
    expect(server.allXml.match(/<w:hdr/g)?.length ?? 0).toBe(1);
    expect(a.clientAllXml()).toBe(server.allXml);
    expect(b.clientAllXml()).toBe(server.allXml);
    await a.unmount();
    await b.unmount();
  });
});
