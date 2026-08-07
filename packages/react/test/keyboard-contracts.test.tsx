// @vitest-environment jsdom
/**
 * Word keyboard contracts (conformance wave) — collab byte-convergence.
 *
 * The editing-conformance matrix pins the LOCAL behavior of each fixed
 * gesture; these tests replay the same gestures through a live hub
 * (CollabEditor + CollabHub) and assert byte-identical convergence at
 * unmount, because each fix routes through canonical mutations + emitted
 * intents:
 *
 *  - G15: typing over a zero-segment boundary selection (Shift+ArrowRight
 *    into an empty paragraph) emits mergeParagraph + insertText.
 *  - G14: Enter over a selection emits deleteText intents, then the split.
 *
 * Harness pattern copied from list-delete.test.tsx (the 734b6a8 shape).
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { serializeXml, type DocxDocument } from "@wordinweb/core";
import type { DocxViewApi } from "../src/index.js";

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 30) { for (let i = 0; i < n; i++) await tick(); }

type Key = string | { key: string; ctrl?: boolean; shift?: boolean };

function keySender(container: HTMLElement) {
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
  return async (seq: Key[]) => {
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
}

async function clickFirstSpan(container: HTMLElement) {
  const page = container.querySelector<HTMLElement>(".dxw-page")!;
  const span = page.querySelector("span") ?? page;
  await act(async () => {
    const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    span.dispatchEvent(new MouseEvent("mousedown", opts));
    span.dispatchEvent(new MouseEvent("mouseup", opts));
  });
}

/** Body paragraph texts, one string per w:p. */
function paragraphTexts(doc: DocxDocument): string[] {
  const root = doc.editableRoots()[0] as unknown as { name: string; children: { name: string }[] };
  const body = (root.children as { name: string; children: never[] }[]).find((c) => c.name.endsWith("body"))!;
  const textOf = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(textOf).join("");
  return (body.children as { name: string; text: string; children: unknown[] }[])
    .filter((c) => c.name.endsWith(":p"))
    .map((p) => textOf(p as never));
}

const provider: DocProvider = { load: () => blankDocxBytes() };
let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `kc${factorySeq++}-`;
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

function serverDoc(hub: CollabHub, docId: string): DocxDocument {
  const room = (hub as unknown as { rooms: Map<string, { session: { doc: DocxDocument } }> }).rooms.get(docId)!;
  return room.session.doc;
}

async function mountCollab(hub: CollabHub, docId: string, clientId: string) {
  const factory = factoryFor(hub);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let api: DocxViewApi | null = null;
  await act(async () => {
    root.render(createElement(CollabEditor, {
      url: "ws://x", docId, clientId, createSocket: factory, onReady: (a: DocxViewApi) => { api = a; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const keys = keySender(container);
  const clientDoc = (): DocxDocument => {
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
    throw new Error("collab doc not found in React tree");
  };
  return {
    container, keys,
    typed: async (t: string) => keys([...t]),
    click: async () => { await clickFirstSpan(container); await tick(); },
    api: () => api!,
    clientDoc,
    texts: () => paragraphTexts(clientDoc()),
    /** EMISSION GUARD: the client's replica must byte-equal the server's —
     * a local-only mutation (nothing emitted) forks silently otherwise. */
    unmount: async () => {
      await settle();
      expect(serializeXml(clientDoc().docRoot), "client diverged from server — a mutation never emitted an intent")
        .toBe(serializeXml(serverDoc(hub, docId).docRoot));
      await act(async () => { root.unmount(); });
    },
  };
}

describe("keyboard contracts replicate (collab mount)", () => {
  it("G15: typing over a boundary selection into an empty paragraph merges everywhere", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g15", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys(["Enter"]); // "ab" + empty paragraph
    expect(ed.texts()).toEqual(["ab", ""]);
    await ed.keys(["ArrowLeft"]); // back to the end of "ab"
    await ed.keys([{ key: "ArrowRight", shift: true }]); // the zero-segment boundary selection
    await ed.typed("k"); // mergeParagraph + insertText intents
    expect(ed.texts()).toEqual(["abk"]);
    await ed.unmount();
  });

  it("G15: arrow keys collapse a boundary selection instead of wedging", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g15b", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys(["Enter", "ArrowLeft"]);
    await ed.keys([{ key: "ArrowRight", shift: true }, "ArrowLeft"]); // collapse back
    await ed.typed("z"); // lands at the collapse point, no merge
    expect(ed.texts()).toEqual(["abz", ""]);
    await ed.unmount();
  });

  it("G4: pending Cmd+B at a caret makes the next typed text bold, on every replica", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g4", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys([{ key: "b", ctrl: true }]); // pending bold (no selection)
    await ed.typed("cd"); // first char consumes the toggle; second continues bold
    expect(ed.texts()).toEqual(["abcd"]);
    const xml = serializeXml(ed.clientDoc().docRoot);
    expect(xml).toMatch(/<w:b\/>/);
    expect(xml).toContain("cd"); // the bold tail stayed one run's text
    await ed.unmount(); // formatRange/formatRun rode the wire byte-exactly
  });

  it("G4: the same shortcut twice cancels the pending toggle", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g4b", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys([{ key: "b", ctrl: true }, { key: "b", ctrl: true }]);
    await ed.typed("c");
    expect(ed.texts()).toEqual(["abc"]);
    expect(serializeXml(ed.clientDoc().docRoot)).not.toMatch(/<w:b\/>/);
    await ed.unmount();
  });

  it("G4: caret movement clears the pending toggle", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g4c", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys([{ key: "b", ctrl: true }, "ArrowLeft", "ArrowRight"]);
    await ed.typed("c");
    expect(ed.texts()).toEqual(["abc"]);
    expect(serializeXml(ed.clientDoc().docRoot)).not.toMatch(/<w:b\/>/);
    await ed.unmount();
  });

  it("G2: Shift+Enter inserts a w:br soft break in-paragraph, on every replica", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g2", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys([{ key: "a", ctrl: true }, "ArrowLeft", "ArrowRight"]); // caret after "a"
    await ed.keys([{ key: "Enter", shift: true }]); // insertSeparator intent
    await ed.typed("z"); // caret landed after the break
    expect(ed.texts()).toEqual(["azb"]); // ONE paragraph
    expect(serializeXml(ed.clientDoc().docRoot)).toMatch(/<w:br\/>/);
    await ed.unmount();
  });

  it("G2: suggesting Shift+Enter records the break as a w:ins and converges", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g2s", "alice");
    await ed.click();
    await ed.typed("ab");
    await ed.keys([{ key: "a", ctrl: true }, "ArrowLeft", "ArrowRight"]);
    ed.api().setSuggesting(true, "Alice");
    await tick();
    await ed.keys([{ key: "Enter", shift: true }]);
    const xml = serializeXml(ed.clientDoc().docRoot);
    expect(xml).toMatch(/<w:ins [^>]*>.*<w:br\/>/);
    expect(ed.texts()).toEqual(["ab"]); // nothing removed, one paragraph
    await ed.unmount();
  });

  it("G10: Enter at the end of a heading starts a Normal paragraph, on every replica", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g10", "alice");
    await ed.click();
    await ed.typed("Head");
    ed.api().setParagraphStyle("Heading1");
    await settle();
    await ed.keys(["End", "Enter"]); // splitParagraph — next-style applies in the shared mutation
    await ed.typed("x");
    expect(ed.texts()).toEqual(["Head", "x"]);
    const xml = serializeXml(ed.clientDoc().docRoot);
    expect((xml.match(/Heading1/g) ?? []).length).toBe(1); // body: the heading only
    await ed.unmount(); // server derived the same next-style from the same styles part
  });

  it("G14: Enter over a selection deletes it then splits, on every replica", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "kc-g14", "alice");
    await ed.click();
    await ed.typed("abcd");
    await ed.keys([{ key: "a", ctrl: true }, "ArrowLeft", "ArrowRight"]); // caret after "a"
    await ed.keys([{ key: "ArrowRight", shift: true }, { key: "ArrowRight", shift: true }]); // select "bc"
    await ed.keys(["Enter"]); // deleteText + splitParagraph intents
    expect(ed.texts()).toEqual(["a", "d"]);
    await ed.unmount();
  });
});
