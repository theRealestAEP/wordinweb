// @vitest-environment jsdom
/**
 * Cmd+Z in a shared session.
 *
 * THE BUG THIS PINS: the editor's local history stack was never gated in
 * collab mode. The toolbar's undo button is hidden in a room
 * (COLLAB_TOOLBAR_DEFAULTS `history: false`), which made it look handled —
 * but the KEYBOARD path ran straight into `applyHistory`, replaying local
 * history against the document with no intent on the wire. Measured before
 * the fix: typing "hello" emitted 5 intents, then Cmd+Z changed the document
 * and emitted 0. Cmd+Z is a reflex, so every room was one keystroke from a
 * silent permanent fork.
 *
 * The rule (checkpoint A18): past the collab gate a command either rides the
 * wire or does nothing. Undo now routes to the collaborative path; with no
 * host hook wired it declines instead of mutating.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { blankDocxBytes } from "@wordinweb/server";
import { serializeXml, type DocxDocument, type EditorIntent } from "@wordinweb/core";

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

async function mount(opts: { collab: boolean }) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const intents: EditorIntent[] = [];
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(),
      editable: true,
      onReady: (a: DocxViewApi) => { seen.api = a; },
      onLoad: (i: { document: DocxDocument }) => { seen.doc = i.document; },
      collab: opts.collab
        ? { submit: (i: EditorIntent) => { intents.push(i); }, allocIds: (n: number) => Array.from({ length: n }, (_, k) => 500_000 + k) }
        : undefined,
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
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
  const type = async (s: string) => {
    for (const ch of s) {
      await act(async () => { target().dispatchEvent(new KeyboardEvent("keydown", { key: ch, bubbles: true, cancelable: true })); });
      await tick(2);
    }
  };
  const undoKey = async () => {
    await act(async () => { target().dispatchEvent(new KeyboardEvent("keydown", { key: "z", metaKey: true, bubbles: true, cancelable: true })); });
    await tick(20);
  };
  return {
    intents, click, type, undoKey,
    xml: () => serializeXml(seen.doc!.docRoot),
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

describe("Cmd+Z in a shared session never mutates locally", () => {
  it("does not change the document without emitting (was a silent fork)", async () => {
    const ed = await mount({ collab: true });
    await ed.click();
    await ed.type("hello");
    const typedIntents = ed.intents.length;
    expect(typedIntents).toBeGreaterThan(0); // typing rides the wire…
    const afterTyping = ed.xml();

    await ed.undoKey();

    // …and undo must not quietly rewrite this replica. Either it emitted an
    // intent (the wired collaborative path) or it did nothing at all.
    const changed = ed.xml() !== afterTyping;
    const emitted = ed.intents.length > typedIntents;
    expect(
      !changed || emitted,
      "Cmd+Z mutated the document in collab without emitting an intent — the room forks silently",
    ).toBe(true);
    await ed.unmount();
  });

  it("still performs a real LOCAL undo when not in a room", async () => {
    // The gate must not cost solo users their undo.
    const ed = await mount({ collab: false });
    await ed.click();
    await ed.type("hello");
    const afterTyping = ed.xml();
    expect(afterTyping).toContain("hello");

    await ed.undoKey();
    expect(ed.xml()).not.toBe(afterTyping); // local history really ran
    await ed.unmount();
  });
});
