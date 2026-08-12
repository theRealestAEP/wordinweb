// @vitest-environment jsdom
/**
 * The keyboard shortcut surface (shortcuts.ts), driven end to end.
 *
 * Every case dispatches a real KeyboardEvent at the live editor and asserts
 * the DOCUMENT changed — a handler that ran and did nothing is a failure, not
 * a pass. Each editor is mounted against a live CollabHub, so unmount also
 * asserts byte-convergence with the server: a shortcut that mutates locally
 * without emitting its intent forks the room, and that is caught here rather
 * than in production.
 *
 * jsdom reports a non-Apple platform, so `ctrl` stands in for the mod key.
 * Harness pattern copied from keyboard-contracts.test.tsx.
 */
import { afterEach, describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { serializeXml, type DocxDocument } from "@wordinweb/core";
import { DocxView, type DocxViewApi } from "../src/index.js";

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 30) { for (let i = 0; i < n; i++) await tick(); }

interface Mods { ctrl?: boolean; shift?: boolean; alt?: boolean }

function keySender(container: HTMLElement) {
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
  return async (key: string, mods: Mods = {}) => {
    await act(async () => {
      target().dispatchEvent(new KeyboardEvent("keydown", {
        key, ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, altKey: !!mods.alt, bubbles: true, cancelable: true,
      }));
      await new Promise((r) => setTimeout(r, 2));
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
  const ns = `ks${factorySeq++}-`;
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

async function mount(docId: string) {
  const hub = new CollabHub(provider);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let api: DocxViewApi | null = null;
  await act(async () => {
    root.render(createElement(CollabEditor, {
      url: "ws://x", docId, clientId: "alice", createSocket: factoryFor(hub),
      onReady: (a: DocxViewApi) => { api = a; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const press = keySender(container);
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
  const ed = {
    container,
    press,
    api: () => api!,
    typed: async (t: string) => { for (const ch of t) await press(ch); },
    /** jsdom has no layout, so text is selected through find(), not a drag. */
    select: async (q: string) => {
      let n = 0;
      await act(async () => { n = api!.find(q); });
      expect(n, `landmark ${JSON.stringify(q)} not found`).toBeGreaterThan(0);
      await tick();
    },
    xml: () => serializeXml(clientDoc().docRoot),
    texts: () => paragraphTexts(clientDoc()),
    /** EMISSION GUARD: the replica must byte-equal the server's document. */
    unmount: async () => {
      await settle();
      expect(serializeXml(clientDoc().docRoot), "client diverged from server — a shortcut never emitted its intent")
        .toBe(serializeXml(serverDoc(hub, docId).docRoot));
      await act(async () => { root.unmount(); });
    },
  };
  await clickFirstSpan(container);
  await tick();
  return ed;
}

/**
 * A single-process editor, for the shortcuts that emit no intent (navigation)
 * or that the collaboration path does not carry yet (api.addComment declines
 * in a room). Byte-convergence has nothing to say about either.
 */
async function mountLocal() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let api: DocxViewApi | null = null;
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(), editable: true, onReady: (a: DocxViewApi) => { api = a; },
    }));
  });
  for (let i = 0; i < 60 && (!api || !container.querySelector(".dxw-page")); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const press = keySender(container);
  const ed = {
    press,
    api: () => api!,
    typed: async (t: string) => { for (const ch of t) await press(ch); },
    select: async (q: string) => {
      let n = 0;
      await act(async () => { n = api!.find(q); });
      expect(n, `landmark ${JSON.stringify(q)} not found`).toBeGreaterThan(0);
      await tick();
    },
    xml: () => serializeXml(api!.document.docRoot),
    texts: () => paragraphTexts(api!.document),
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
  await clickFirstSpan(container);
  await tick();
  return ed;
}

afterEach(() => {
  // The dialogs mount on document.body and outlive their editor.
  for (const el of document.querySelectorAll('[role="dialog"]')) el.parentElement?.remove();
});

describe("character formatting shortcuts", () => {
  it("⌘⇧X strikes the selection", async () => {
    const ed = await mount("ks-strike");
    await ed.typed("alpha");
    await ed.select("alpha");
    await ed.press("x", { ctrl: true, shift: true });
    expect(ed.xml()).toMatch(/<w:strike\/>/);
    await ed.unmount();
  });

  it("⌘⇧= raises the selection and pressing it again lowers it", async () => {
    const ed = await mount("ks-super");
    await ed.typed("alpha");
    await ed.select("alpha");
    await ed.press("=", { ctrl: true, shift: true });
    expect(ed.xml()).toMatch(/<w:vertAlign w:val="superscript"\/>/);
    await ed.press("=", { ctrl: true, shift: true });
    expect(ed.xml()).not.toMatch(/vertAlign/);
    await ed.unmount();
  });

  it("⌘= lowers the selection", async () => {
    const ed = await mount("ks-sub");
    await ed.typed("alpha");
    await ed.select("alpha");
    await ed.press("=", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:vertAlign w:val="subscript"\/>/);
    await ed.unmount();
  });

  it("⌘⇧> grows and ⌘⇧< shrinks along Word's size ladder", async () => {
    const ed = await mount("ks-size");
    await ed.typed("alpha");
    await ed.select("alpha");
    await act(async () => { ed.api().applyFormat({ fontSizePt: 12 }); });
    await settle();
    await ed.press(">", { ctrl: true, shift: true });
    expect(ed.xml(), "12pt grows to 14pt (w:sz is half-points)").toMatch(/<w:sz w:val="28"\/>/);
    await ed.press("<", { ctrl: true, shift: true });
    expect(ed.xml()).toMatch(/<w:sz w:val="24"\/>/);
    await ed.unmount();
  });

  it("⌘\\ clears direct character formatting", async () => {
    const ed = await mount("ks-clear");
    await ed.typed("alpha");
    await ed.select("alpha");
    await ed.press("b", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:b\/>/);
    await ed.press("\\", { ctrl: true });
    expect(ed.xml()).not.toMatch(/<w:b\/>/);
    await ed.unmount();
  });

  it("⌥⌘C then ⌥⌘V paints formatting from one selection onto another", async () => {
    const ed = await mount("ks-painter");
    await ed.typed("alpha bravo");
    await ed.select("alpha");
    await ed.press("b", { ctrl: true });
    await ed.press("c", { ctrl: true, alt: true });
    await ed.select("bravo");
    expect(ed.xml(), "bravo is still plain before the paste").toMatch(/<w:b\/>[\s\S]*alpha/);
    await ed.press("v", { ctrl: true, alt: true });
    expect((ed.xml().match(/<w:b\/>/g) ?? []).length, "both runs are bold now").toBe(2);
    await ed.unmount();
  });
});

describe("paragraph shortcuts", () => {
  it("⌘] indents and ⌘[ removes the indent", async () => {
    const ed = await mount("ks-indent");
    await ed.typed("alpha");
    await ed.press("]", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:ind /);
    await ed.press("[", { ctrl: true });
    expect(ed.xml()).not.toMatch(/<w:ind w:left="720"/);
    await ed.unmount();
  });

  it("⌘2 sets double line spacing and ⌘1 puts it back to single", async () => {
    const ed = await mount("ks-spacing");
    await ed.typed("alpha");
    await ed.press("2", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:spacing[^>]*w:line="480"/);
    await ed.press("1", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:spacing[^>]*w:line="240"/);
    await ed.unmount();
  });

  it("⌘⌥1 still applies Heading 1 and ⌘⌥0 still returns to Normal", async () => {
    const ed = await mount("ks-heading");
    await ed.typed("alpha");
    await ed.press("1", { ctrl: true, alt: true });
    expect(ed.xml()).toContain("Heading1");
    await ed.press("0", { ctrl: true, alt: true });
    expect(ed.xml()).not.toContain("Heading1");
    await ed.unmount();
  });

  it("⌘E still centers the paragraph", async () => {
    const ed = await mount("ks-align");
    await ed.typed("alpha");
    await ed.press("e", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:jc w:val="center"\/>/);
    await ed.unmount();
  });

  it("⌘⇧L still starts a bulleted list", async () => {
    const ed = await mount("ks-bullet");
    await ed.typed("alpha");
    await ed.press("l", { ctrl: true, shift: true });
    expect(ed.xml()).toContain("<w:numPr>");
    await ed.unmount();
  });
});

describe("structure shortcuts", () => {
  it("⌘Enter still inserts a page break and ⌘⇧Enter a column break", async () => {
    const ed = await mount("ks-breaks");
    await ed.typed("alpha");
    await ed.press("Enter", { ctrl: true });
    expect(ed.xml()).toMatch(/<w:br w:type="page"\/>/);
    await ed.press("Enter", { ctrl: true, shift: true });
    expect(ed.xml()).toMatch(/<w:br w:type="column"\/>/);
    await ed.unmount();
  });
});

describe("review shortcuts", () => {
  it("⌘⇧E turns track changes on, so the next typing records as a suggestion", async () => {
    const ed = await mount("ks-track");
    await ed.typed("alpha");
    expect(ed.api().isSuggesting()).toBe(false);
    await ed.press("e", { ctrl: true, shift: true });
    expect(ed.api().isSuggesting()).toBe(true);
    await ed.typed("X");
    expect(ed.xml()).toMatch(/<w:ins [^>]*>/);
    await ed.press("e", { ctrl: true, shift: true });
    expect(ed.api().isSuggesting()).toBe(false);
    await ed.unmount();
  });

  it("⌥⌘M opens the comment dialog and anchors the comment to the selection", async () => {
    const ed = await mountLocal();
    await ed.typed("alpha");
    await ed.select("alpha");
    await ed.press("m", { ctrl: true, alt: true });
    await tick();
    const dialog = document.querySelector<HTMLFormElement>('[role="dialog"]')!;
    expect(dialog.textContent).toContain("New comment");
    const input = dialog.querySelector<HTMLTextAreaElement>("textarea")!;
    input.value = "look here";
    await act(async () => {
      input.dispatchEvent(new Event("input", { bubbles: true }));
      dialog.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
    });
    await tick();
    await tick();
    expect(ed.xml()).toMatch(/commentRangeStart/);
    await ed.unmount();
  });

  it("⌥⌘N selects the next comment's anchor", async () => {
    const ed = await mountLocal();
    await ed.typed("alpha bravo");
    await ed.select("alpha");
    await act(async () => { expect(ed.api().addComment("first")).toBe(true); });
    await tick();
    // Move the selection away, then step back to the comment and overtype it.
    await ed.select("bravo");
    await ed.press("n", { ctrl: true, alt: true });
    await tick();
    await ed.typed("Z");
    expect(ed.texts()[0], "the commented word was the selection").toBe("Z bravo");
    await ed.unmount();
  });
});

describe("navigation shortcuts", () => {
  it("⌘Home and ⌘End move the caret to the ends of the story", async () => {
    const ed = await mount("ks-docedge");
    await ed.typed("alpha");
    await ed.press("Enter");
    await ed.typed("bravo");
    await ed.press("Home", { ctrl: true });
    await ed.typed("1");
    expect(ed.texts()).toEqual(["1alpha", "bravo"]);
    await ed.press("End", { ctrl: true });
    await ed.typed("2");
    expect(ed.texts()).toEqual(["1alpha", "bravo2"]);
    await ed.unmount();
  });

  it("⌘Up and ⌘Down step between paragraphs", async () => {
    const ed = await mount("ks-para");
    await ed.typed("alpha");
    await ed.press("Enter");
    await ed.typed("bravo");
    // Word: the first press reaches the current paragraph's start…
    await ed.press("ArrowUp", { ctrl: true });
    await ed.typed("1");
    expect(ed.texts()).toEqual(["alpha", "1bravo"]);
    // …the next one reaches the paragraph above.
    await ed.press("ArrowUp", { ctrl: true });
    await ed.press("ArrowUp", { ctrl: true });
    await ed.typed("2");
    expect(ed.texts()).toEqual(["2alpha", "1bravo"]);
    await ed.press("ArrowDown", { ctrl: true });
    await ed.typed("3");
    expect(ed.texts()).toEqual(["2alpha", "31bravo"]);
    await ed.unmount();
  });

  it("Page Down moves to the next page and Page Up back", async () => {
    const ed = await mount("ks-page");
    await ed.typed("alpha");
    await ed.press("Enter", { ctrl: true }); // page break
    await ed.typed("bravo");
    expect(ed.api().pageCount(), "the break made a second page").toBe(2);
    await ed.press("PageUp");
    await ed.typed("1");
    expect(ed.texts()[0], "the caret reached the first page").toBe("1alphabravo");
    await ed.press("PageDown");
    await ed.typed("2");
    expect(ed.texts()[0], "and back to the second page's first text").toBe("1alpha2bravo");
    await ed.unmount();
  });

  it("⌥⌘G opens the go-to-page prompt", async () => {
    const ed = await mount("ks-goto");
    await ed.typed("alpha");
    await ed.press("g", { ctrl: true, alt: true });
    await tick();
    expect(document.querySelector('[role="dialog"]')?.textContent).toContain("Go to page");
    document.querySelector<HTMLElement>('[role="dialog"]')?.closest("div")?.remove();
    await ed.unmount();
  });
});

describe("Tab and Shift+Tab in lists and tables", () => {
  it("Tab in the last cell appends a row on every replica", async () => {
    const ed = await mount("ks-tabrow");
    await ed.typed("alpha");
    await act(async () => { ed.api().insertTable(2, 2); });
    await settle();
    await ed.press("ArrowRight"); // into the first cell
    await ed.press("Tab");
    await ed.press("Tab");
    await ed.press("Tab"); // last cell of the last row
    await ed.press("Tab"); // Word: appends a row
    expect((ed.xml().match(/<w:tr[ >]/g) ?? []).length, "a third row exists").toBe(3);
    await ed.unmount(); // and it rode the wire
  });
});
