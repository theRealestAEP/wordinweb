// @vitest-environment jsdom
/**
 * THE SAME GESTURE MUST LEAVE THE SAME DOCUMENT, LOCAL OR IN A ROOM.
 *
 * Enter at a paragraph start takes a different route depending on whether a
 * collab connection is attached. With one, `onIntent` is set and Enter goes
 * through the general split. Without one it takes a local fast path,
 * `insertBlankParagraphBeforeAtStart`, which splices a blank paragraph in
 * before the caret's paragraph and keeps the original parsed model — cheaper,
 * and correct as far as the XML goes.
 *
 * Two routes to one gesture is a standing invitation to drift, and it drifted:
 * the fast path used to hand the caret to the blank it had just inserted
 * ABOVE, so the caret kept its old y while the content slid down under it
 * (fixed in bb8bd58). Nothing caught it, because every existing convergence
 * test drives the COLLAB route, which was always right, and the local route
 * had no oracle at all.
 *
 * This is that oracle, and it needs no view on which answer is correct: run
 * the identical keystrokes through both routes and require the same document.
 * The caret is what differs when these paths drift, and a caret is not in the
 * XML — so every gesture below ENDS BY TYPING a marker character. Where the
 * marker lands is where the caret was, which turns caret position into
 * document text and makes it directly comparable.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { DocxView } from "../src/index.js";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";
import type { DocxDocument } from "@wordinweb/core";

const provider: DocProvider = { load: () => blankDocxBytes() };

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 30) { for (let i = 0; i < n; i++) await tick(); }

let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `p${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => (delayMs > 0 ? setTimeout(fn, delayMs) : fn());
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `${ns}c${n++}`, send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))) };
    let opened = false;
    return {
      send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); }
      },
    } as unknown as WebSocket;
  };
}

/** Paragraph texts of a body, joined — the comparable shape of a document. */
function bodyText(docRoot: { name: string; text: string; children: unknown[] }): string {
  const body = (docRoot.children as { name: string }[]).find((c) => c.name.endsWith("body"))!;
  const textOf = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(textOf).join("");
  return (body as unknown as { children: never[] }).children
    .filter((c: { name: string }) => c.name.endsWith(":p"))
    .map(textOf)
    .join("|");
}

/** Drive a mounted editor's own key handlers, scoped to its container. */
function driver(container: HTMLElement) {
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
  return async (seq: string[]) => {
    await act(async () => {
      for (const key of seq) {
        target().dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
        await new Promise((r) => setTimeout(r, 2));
      }
    });
  };
}

async function clickIn(container: HTMLElement) {
  const page = container.querySelector<HTMLElement>(".dxw-page")!;
  const span = page.querySelector("span") ?? page;
  await act(async () => {
    const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    span.dispatchEvent(new MouseEvent("mousedown", opts));
    span.dispatchEvent(new MouseEvent("mouseup", opts));
  });
  await tick();
}

/** Run `keys` through a plain editable DocxView — the LOCAL route. */
async function runLocal(keys: string[]): Promise<string> {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { doc: DocxDocument | null } = { doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(),
      editable: true,
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  await clickIn(container);
  await driver(container)(keys);
  await tick();
  const text = bodyText(seen.doc!.docRoot as never);
  await act(async () => { root.unmount(); });
  container.remove();
  return text;
}

/** Run `keys` through a CollabEditor in a room — the GENERAL route. Reads the
 * SERVER's copy through a spy client, so a local mutation that never emitted
 * cannot pass as agreement. */
async function runCollab(keys: string[], docId: string): Promise<string> {
  const hub = new CollabHub(provider);
  const factory = factoryFor(hub);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(CollabEditor, { url: "ws://x", docId, clientId: "alice", createSocket: factory }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  await clickIn(container);
  await driver(container)(keys);
  await settle();

  const spy = new CollabConnection(createWebSocketTransport(factory("ws://spy") as never), `spy-${docId}`);
  spy.join(docId);
  for (let i = 0; i < 40 && !spy.doc; i++) await tick();
  expect(spy.doc, "spy never received the document").toBeTruthy();
  const text = bodyText(spy.doc!.docRoot as never);
  await act(async () => { root.unmount(); });
  container.remove();
  return text;
}

/** Every gesture ends by typing "Z": where it lands is where the caret was. */
const GESTURES: { name: string; keys: string[] }[] = [
  // The reported bug. Enter at a paragraph start puts the blank above and the
  // caret stays with the text, so the marker joins the text on line 2.
  { name: "Enter at a paragraph start", keys: [..."ab", "Home", "Enter", "Z"] },
  // Every Enter after the first is an Enter on an EMPTY paragraph, which is
  // the same fast path (offset 0) and the shape the reporter actually hit.
  { name: "Enter twice at a paragraph end", keys: [..."ab", "End", "Enter", "Enter", "Z"] },
  { name: "Enter three times at a paragraph end", keys: [..."ab", "End", "Enter", "Enter", "Enter", "Z"] },
  // Mid-paragraph split does NOT take the fast path (offset > 0); included so
  // the comparison covers the general route through the same harness.
  { name: "Enter mid-paragraph", keys: [..."abcd", "Home", "ArrowRight", "ArrowRight", "Enter", "Z"] },
  // Enter at start, then again at the start of the paragraph now below it.
  { name: "Enter twice at a paragraph start", keys: [..."ab", "Home", "Enter", "Enter", "Z"] },
];

describe("local and collab routes agree on the document a gesture leaves", () => {
  for (const { name, keys } of GESTURES) {
    it(name, async () => {
      const local = await runLocal(keys);
      const collab = await runCollab(keys, `parity-${name.replace(/\W+/g, "-")}`);
      expect(local, `local route diverged from the collab route for: ${name}`).toBe(collab);
    });
  }
});
