// @vitest-environment jsdom
/**
 * Deleting list items the way Word does (user report: a page of empty
 * bulleted items whose markers could not be deleted — the caret sat on an
 * empty bullet at the start of the document and Backspace was a dead key).
 *
 * Word's contract, asserted here against BOTH mounts:
 *   1. Backspace at the START of a list item first removes the numbering
 *      (the item becomes a plain paragraph, marker gone); only the NEXT
 *      press merges into the previous paragraph. This is also the only
 *      escape for a list item at the very start of a document.
 *   2. Enter on an EMPTY list item exits the list (no new bullet).
 *   3. A selection spanning list items deletes the paragraphs entirely,
 *      markers included — an EMPTY item paints no text bindings, so it
 *      used to survive every selection delete as a marker-only paragraph.
 *   4. The toolbar bullet button toggles OFF on an already-bulleted item.
 *
 * Local (DocxView) tests drive the app's path; CollabEditor tests replay
 * the same gestures through a live hub and assert byte-identical
 * convergence (the emission guard at unmount), since the fixes emit
 * setListType / mergeParagraph / suggestRevision-mark intents.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

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

/** Body paragraphs of a document as compact strings: "bullet:text" / "plain:text". */
function paragraphShapes(doc: { editableRoots(): { name: string; children: { name: string }[] }[] }): string[] {
  const root = doc.editableRoots()[0] as unknown as { name: string; children: never[] };
  const body = (root.children as { name: string; children: never[] }[]).find((c) => c.name.endsWith("body"))!;
  const textOf = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(textOf).join("");
  // LIVE numbering only: a tracked unbullet keeps the old numPr inside the
  // w:pPrChange snapshot, which must not count as a marker.
  const hasNumPr = (el: { name: string; children: unknown[] }): boolean =>
    el.name.endsWith(":numPr") ||
    (el.children as { name: string; children: unknown[] }[])
      .filter((c) => !c.name.endsWith(":pPrChange"))
      .some(hasNumPr);
  return (body.children as { name: string; text: string; children: unknown[] }[])
    .filter((c) => c.name.endsWith(":p"))
    .map((p) => `${hasNumPr(p as never) ? "bullet" : "plain"}:${textOf(p as never)}`);
}

// ---------------------------------------------------------------- local mount

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
  const keys = keySender(container);
  return {
    container, keys,
    typed: async (t: string) => keys([...t]),
    click: async () => { await clickFirstSpan(container); await tick(); },
    shapes: () => paragraphShapes(seen.doc!),
    api: () => seen.api!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

/** Type a bulleted list the way a user does: bullet the first paragraph via
 * the toolbar, then Enter + type extends the list (the split carries the
 * numbering). One toggleList call on ONE paragraph keeps the fixture clear of
 * the pre-existing multi-paragraph toggle divergence in collab (each
 * per-block setListType intent mints a fresh numId on the server). */
async function typeBulletedList(
  ed: { click(): Promise<void>; typed(t: string): Promise<void>; keys(seq: Key[]): Promise<void>; api(): DocxViewApi; shapes(): string[] },
  items: string[],
) {
  await ed.click();
  await ed.typed(items[0]);
  await ed.keys([{ key: "a", ctrl: true }]);
  ed.api().toggleList("bullet");
  await tick();
  await ed.keys(["End"]); // collapse the select-all to the item's end
  for (const item of items.slice(1)) {
    await ed.keys(["Enter"]);
    await ed.typed(item);
  }
  expect(ed.shapes()).toEqual(items.map((t) => `bullet:${t}`));
}

/** Collapse to document start (Ctrl+A then ArrowLeft is geometry-free), then
 * walk right by n steps — reliable caret placement under jsdom's zero-rect
 * hit testing, where positional clicks all land on the first binding. */
async function caretFromStart(keys: (seq: Key[]) => Promise<void>, n: number) {
  await keys([{ key: "a", ctrl: true }, "ArrowLeft"]);
  await keys(Array.from({ length: n }, () => "ArrowRight"));
}

describe("Backspace at the start of a list item (local mount)", () => {
  it("EMPTY first item: first press removes the marker instead of dead-keying", async () => {
    const ed = await mountLocal();
    await ed.click();
    await ed.typed("a");
    await ed.keys([{ key: "a", ctrl: true }]);
    ed.api().toggleList("bullet");
    await tick();
    expect(ed.shapes()).toEqual(["bullet:a"]);
    await ed.keys(["End", "Backspace"]); // empty the item
    expect(ed.shapes()).toEqual(["bullet:"]);
    // THE BUG: this press did nothing (no previous paragraph to merge into),
    // leaving an undeletable marker. Word unbullets on the first press.
    await ed.keys(["Backspace"]);
    expect(ed.shapes()).toEqual(["plain:"]);
    await ed.unmount();
  });

  it("EMPTY middle item: first press unbullets, second merges it away", async () => {
    const ed = await mountLocal();
    await typeBulletedList(ed, ["aa", "b", "cc"]);
    await caretFromStart(ed.keys, 3); // start of "b"
    await ed.keys(["Delete"]); // empty the middle item
    expect(ed.shapes()).toEqual(["bullet:aa", "bullet:", "bullet:cc"]);
    await ed.keys(["Backspace"]);
    expect(ed.shapes()).toEqual(["bullet:aa", "plain:", "bullet:cc"]);
    await ed.keys(["Backspace"]);
    expect(ed.shapes()).toEqual(["bullet:aa", "bullet:cc"]);
    await ed.unmount();
  });

  it("NON-empty item: first press unbullets keeping the paragraph, second merges", async () => {
    const ed = await mountLocal();
    await typeBulletedList(ed, ["aa", "bb"]);
    await caretFromStart(ed.keys, 3); // start of "bb"
    await ed.keys(["Backspace"]);
    expect(ed.shapes()).toEqual(["bullet:aa", "plain:bb"]);
    await ed.keys(["Backspace"]);
    expect(ed.shapes()).toEqual(["bullet:aabb"]);
    await ed.unmount();
  });
});

describe("Delete key on empty list items (local mount)", () => {
  it("Delete at the end of an item removes the following empty item", async () => {
    const ed = await mountLocal();
    await typeBulletedList(ed, ["aa", "b", "cc"]);
    await caretFromStart(ed.keys, 3);
    await ed.keys(["Delete"]); // empty the middle item
    await ed.keys(["ArrowLeft"]); // end of "aa"
    await ed.keys(["Delete"]);
    expect(ed.shapes()).toEqual(["bullet:aa", "bullet:cc"]);
    await ed.unmount();
  });
});

describe("Enter on an empty list item (local mount)", () => {
  it("exits the list instead of adding another bullet", async () => {
    const ed = await mountLocal();
    await ed.click();
    await ed.typed("aa");
    await ed.keys([{ key: "a", ctrl: true }]);
    ed.api().toggleList("bullet");
    await tick();
    await ed.keys(["End", "Enter"]); // new empty bullet
    expect(ed.shapes()).toEqual(["bullet:aa", "bullet:"]);
    await ed.keys(["Enter"]); // exit
    expect(ed.shapes()).toEqual(["bullet:aa", "plain:"]);
    await ed.unmount();
  });
});

describe("selection spanning list items (local mount)", () => {
  it("Backspace removes the spanned paragraphs, markers included", async () => {
    const ed = await mountLocal();
    await typeBulletedList(ed, ["aa", "b", "cc"]);
    await caretFromStart(ed.keys, 3);
    await ed.keys(["Delete"]); // aa | <empty bullet> | cc
    expect(ed.shapes()).toEqual(["bullet:aa", "bullet:", "bullet:cc"]);
    // Select from the start of "aa" across the EMPTY item into "cc" (the
    // empty item paints no bindings — only the span's ends carry segments).
    await ed.keys([{ key: "a", ctrl: true }, "ArrowLeft"]);
    await ed.keys([
      { key: "ArrowRight", shift: true }, { key: "ArrowRight", shift: true }, // aa
      { key: "ArrowRight", shift: true }, // step into the empty item
      { key: "ArrowRight", shift: true }, // step into cc
      { key: "ArrowRight", shift: true }, // first char of cc
    ]);
    await ed.keys(["Backspace"]);
    // Word: one paragraph remains (the survivor keeps its own list format);
    // the fully-selected items and the empty marker-only item are GONE.
    expect(ed.shapes()).toEqual(["bullet:c"]);
    await ed.unmount();
  });
});

describe("bullet toolbar button as a toggle (local mount)", () => {
  it("toggles OFF on an already-bulleted selection", async () => {
    const ed = await mountLocal();
    await ed.click();
    await ed.typed("aa");
    await ed.keys([{ key: "a", ctrl: true }]);
    ed.api().toggleList("bullet");
    await tick();
    expect(ed.api().getListType()).toBe("bullet");
    expect(ed.shapes()).toEqual(["bullet:aa"]);
    ed.api().toggleList("bullet");
    await tick();
    expect(ed.api().getListType()).toBe(null);
    expect(ed.shapes()).toEqual(["plain:aa"]);
    await ed.unmount();
  });
});

// ---------------------------------------------------------------- collab mount

const provider: DocProvider = { load: () => blankDocxBytes() };
let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `ld${factorySeq++}-`;
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
    shapes: () => paragraphShapes(clientDoc()),
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

describe("list deletion replicates (collab mount)", () => {
  it("Backspace unbullets an empty first item on every replica", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "ld-unbullet", "alice");
    await ed.click();
    await ed.typed("a");
    await ed.keys([{ key: "a", ctrl: true }]);
    ed.api().toggleList("bullet");
    await settle();
    await ed.keys(["End", "Backspace"]); // empty the item
    await settle();
    expect(ed.shapes()).toEqual(["bullet:"]);
    await ed.keys(["Backspace"]); // unbullet (setListType null intent)
    await settle();
    expect(ed.shapes()).toEqual(["plain:"]);
    expect(paragraphShapes(serverDoc(hub, "ld-unbullet"))).toEqual(["plain:"]);
    await ed.unmount(); // byte-parity guard
  });

  it("Backspace at the start of a NON-empty item unbullets then merges, converging", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "ld-two-step", "alice");
    await typeBulletedList(ed, ["aa", "bb"]);
    await settle();
    await caretFromStart(ed.keys, 3); // start of "bb"
    await ed.keys(["Backspace"]);
    await settle();
    expect(ed.shapes()).toEqual(["bullet:aa", "plain:bb"]);
    await ed.keys(["Backspace"]);
    await settle();
    expect(ed.shapes()).toEqual(["bullet:aabb"]);
    await ed.unmount();
  });

  it("selection spanning an empty list item deletes it everywhere (mergeParagraph intents)", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "ld-span", "alice");
    await typeBulletedList(ed, ["aa", "b", "cc"]);
    await settle();
    await caretFromStart(ed.keys, 3);
    await ed.keys(["Delete"]); // aa | <empty bullet> | cc
    await settle();
    expect(ed.shapes()).toEqual(["bullet:aa", "bullet:", "bullet:cc"]);
    await ed.keys([{ key: "a", ctrl: true }, "ArrowLeft"]);
    await ed.keys([
      { key: "ArrowRight", shift: true }, { key: "ArrowRight", shift: true },
      { key: "ArrowRight", shift: true }, { key: "ArrowRight", shift: true },
      { key: "ArrowRight", shift: true },
    ]);
    await ed.keys(["Backspace"]);
    await settle();
    expect(ed.shapes()).toEqual(["bullet:c"]);
    expect(paragraphShapes(serverDoc(hub, "ld-span"))).toEqual(["bullet:c"]);
    await ed.unmount();
  });

  it("suggesting: Backspace on a list item records a tracked unbullet; a spanning delete strikes the marks — and converges", async () => {
    const hub = new CollabHub(provider);
    const ed = await mountCollab(hub, "ld-suggest", "alice");
    await typeBulletedList(ed, ["aa", "bb"]);
    await settle();
    ed.api().setSuggesting(true, "Reviewer");
    await tick();
    await caretFromStart(ed.keys, 3); // start of "bb"
    await ed.keys(["Backspace"]); // tracked unbullet: w:pPrChange, not a merge
    await settle();
    const afterUnbullet = serializeXml(ed.clientDoc().docRoot);
    expect(ed.shapes()).toEqual(["bullet:aa", "plain:bb"]);
    expect(afterUnbullet).toContain("pPrChange");
    // Spanning selection delete: text struck AND the interior paragraph mark
    // marked "del" (Word keeps both paragraphs until accept).
    await ed.keys([{ key: "a", ctrl: true }, "ArrowLeft"]);
    await ed.keys([
      { key: "ArrowRight", shift: true }, { key: "ArrowRight", shift: true },
      { key: "ArrowRight", shift: true }, { key: "ArrowRight", shift: true },
    ]);
    await ed.keys(["Backspace"]);
    await settle();
    const struck = serializeXml(ed.clientDoc().docRoot);
    expect(struck).toContain("w:del"); // struck text and/or paragraph mark
    expect(ed.shapes().length).toBe(2); // nothing actually merged yet
    await ed.unmount(); // byte-parity: suggestRevision marks replicated
  });
});
