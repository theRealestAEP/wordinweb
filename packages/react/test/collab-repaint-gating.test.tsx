// @vitest-environment jsdom
/**
 * REPAINT GATING (the 500-page collab regression pin).
 *
 * DocxView answers every `renderSignal` move with a whole-document relayout
 * (forced-global; async behind an inert container past the >50-page
 * background-layout threshold). The signal therefore may move ONLY when the
 * rendered document changed outside the editor. The editor-driven typing path
 * (submit + own echo) mutates and paints the doc itself — riding the
 * every-onChange `version` counter as the signal queued a redundant
 * whole-document relayout per keystroke, which on a 500-page document meant a
 * multi-second, input-eating stall per keystroke ("the collab editor is
 * orders of magnitude slower").
 *
 * Pins, end to end through a real hub:
 *   1. an own submit + its echo move `version` but NOT `renderVersion`, and
 *      DocxView's received renderSignal stays put;
 *   2. a REMOTE edit moves `renderVersion` AND actually repaints (the text
 *      lands in the DOM) — gating must never eat remote repaints;
 *   3. a toolbar-style `submitOp` (applied by the connection, not the editor)
 *      moves `renderVersion`.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabEditor } from "../src/collab.js";
import type { CollabSession } from "../src/collab.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";

function docWith(text: string): Uint8Array {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
}
const provider: DocProvider = { load: () => docWith("HELLOWORLD") };
function factoryFor(hub: CollabHub) {
  let n = 0;
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `g${n++}`, send: (m: ServerMessage) => ls.forEach((l) => l({ data: JSON.stringify(m) })) };
    let opened = false;
    return { send: (d: string) => { void hub.handle(conn, JSON.parse(d)); },
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}
async function tick() { await act(async () => { await new Promise<void>((r) => setTimeout(r, 5)); }); }

/** The renderSignal DocxView actually received, read off the fiber tree. */
function receivedRenderSignal(container: HTMLElement): number {
  const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
  const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
  let guard = 0;
  while (stack.length && guard++ < 6000) {
    const f = stack.pop() as { memoizedProps?: { collab?: { renderSignal?: number } }; child?: unknown; sibling?: unknown } | null;
    if (!f) continue;
    const sig = f.memoizedProps?.collab?.renderSignal;
    if (typeof sig === "number") return sig;
    if (f.child) stack.push(f.child);
    if (f.sibling) stack.push(f.sibling);
  }
  throw new Error("no DocxView collab.renderSignal in the tree");
}

describe("collab repaint gating (renderVersion vs version)", () => {
  it("own typing does not move the repaint signal; remote and toolbar edits do", async () => {
    const hub = new CollabHub(provider);
    const factory = factoryFor(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    let session: CollabSession | null = null;
    await act(async () => {
      root.render(createElement(CollabEditor, {
        url: "ws://x", docId: "gate", clientId: "typist", createSocket: factory,
        onSession: (s) => { session = s; },
      }));
    });
    for (let i = 0; i < 20 && !container.textContent?.includes("HELLO"); i++) await tick();
    expect(container.textContent).toContain("HELLOWORLD");
    expect(session).not.toBeNull();

    // 1. OWN TYPING (editor-style preApplied submit + its echo): version moves,
    //    the repaint signal does not — the editor already painted this edit.
    const versionBefore = (session as unknown as CollabSession).version;
    const renderBefore = (session as unknown as CollabSession).renderVersion;
    const signalBefore = receivedRenderSignal(container);
    await act(async () => {
      (session as unknown as CollabSession).submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 5 }, text: "T" } as never);
    });
    for (let i = 0; i < 10; i++) await tick(); // echo round trip + re-renders
    const s = session as unknown as CollabSession;
    expect(s.version, "own submit + echo must still tick the bookkeeping version").toBeGreaterThan(versionBefore);
    expect(s.renderVersion, "own submit + echo must NOT move the repaint signal").toBe(renderBefore);
    expect(receivedRenderSignal(container), "DocxView must not receive a repaint for own typing").toBe(signalBefore);

    // 2. REMOTE EDIT: the repaint signal moves and the text actually paints.
    const editor = new CollabConnection(createWebSocketTransport(factory("ws://x") as never), "remote");
    editor.join("gate");
    await tick();
    await act(async () => {
      editor.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "-R-" } as never);
    });
    for (let i = 0; i < 20 && !container.textContent?.includes("-R-"); i++) await tick();
    expect(container.textContent, "a remote edit must still repaint").toContain("-R-");
    expect((session as unknown as CollabSession).renderVersion, "a remote edit must move the repaint signal").toBeGreaterThan(renderBefore);

    // 3. TOOLBAR OP (submitOp — applied by the connection, not the editor):
    //    the repaint signal moves.
    const renderMid = (session as unknown as CollabSession).renderVersion;
    await act(async () => {
      (session as unknown as CollabSession).submitOp({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 1 }, text: "O" } as never);
    });
    for (let i = 0; i < 10; i++) await tick();
    expect((session as unknown as CollabSession).renderVersion, "a connection-applied op must move the repaint signal").toBeGreaterThan(renderMid);

    await act(async () => { root.unmount(); });
  });
});
