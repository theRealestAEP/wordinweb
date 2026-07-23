// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { createElement } from "react";
import { createRoot } from "react-dom/client";
import { act } from "react";
import { useCollab } from "../src/collab.js";
import { CollabHub, blankProvider } from "@wordinweb/server";
import type { Connection, ServerMessage } from "@wordinweb/server";

/** A fake browser WebSocket bridged to an in-process CollabHub — useCollab's
 * real transport + connection run against a real server with no network. */
function makeFakeSocketFactory(hub: CollabHub) {
  let n = 0;
  return (_url: string) => {
    const msgListeners: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = {
      id: `c${n++}`,
      send: (msg: ServerMessage) => msgListeners.forEach((l) => l({ data: JSON.stringify(msg) })),
    };
    let opened = false;
    const pending: string[] = [];
    const socket = {
      send: (data: string) => { void hub.handle(conn, JSON.parse(data)); },
      addEventListener: (type: "message" | "open", cb: never) => {
        if (type === "message") msgListeners.push(cb as (ev: { data: unknown }) => void);
        else if (!opened) { opened = true; (cb as () => void)(); pending.splice(0); }
      },
    };
    void pending;
    return socket as unknown as WebSocket;
  };
}

/** Flush pending microtasks + one macrotask turn inside act. */
async function tick() {
  await act(async () => { await new Promise<void>((r) => setTimeout(r, 5)); });
}

function paraText(doc: { docRoot: { name: string; text: string; children: never[] } } | null): string {
  if (!doc) return "";
  let t = "";
  const walk = (el: { name: string; text: string; children: never[] }): void => {
    if (el.name === "w:t" || el.name.endsWith(":t")) t += el.text;
    (el.children as { name: string; text: string; children: never[] }[]).forEach(walk);
  };
  walk(doc.docRoot);
  return t;
}

describe("useCollab (jsdom React integration)", () => {
  it("joins a doc over an in-process hub and becomes ready with a live document", async () => {
    const hub = new CollabHub(blankProvider);
    const factory = makeFakeSocketFactory(hub);
    const captured: { ready: boolean; doc: unknown }[] = [];
    function Probe() {
      const s = useCollab({ url: "ws://x", docId: "d", clientId: "a", createSocket: factory });
      captured.push({ ready: s.ready, doc: s.doc });
      return createElement("div", null, s.ready ? "ready" : "connecting");
    }
    const root = createRoot(document.createElement("div"));
    await act(async () => { root.render(createElement(Probe)); });
    for (let i = 0; i < 10 && !captured.some((c) => c.ready); i++) await tick();
    expect(captured.some((c) => c.ready && c.doc != null)).toBe(true);
    await act(async () => { root.unmount(); });
  });

  it("one hook's edit reaches another hook on the same doc (inbound broadcast)", async () => {
    const hub = new CollabHub(blankProvider);
    const factory = makeFakeSocketFactory(hub);
    let submitA: ((i: never) => void) | null = null;
    let readyA = false;
    const bDocs: string[] = [];

    function A() {
      const s = useCollab({ url: "ws://x", docId: "shared", clientId: "a", createSocket: factory });
      submitA = s.submit as unknown as (i: never) => void;
      readyA = s.ready;
      return null;
    }
    function B() {
      const s = useCollab({ url: "ws://x", docId: "shared", clientId: "b", createSocket: factory });
      bDocs.push(paraText(s.doc as never));
      return null;
    }
    const rootA = createRoot(document.createElement("div"));
    const rootB = createRoot(document.createElement("div"));
    await act(async () => { rootA.render(createElement(A)); rootB.render(createElement(B)); });
    for (let i = 0; i < 10 && !readyA; i++) await tick();

    await act(async () => { submitA!({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "hi" } as never); });
    for (let i = 0; i < 10 && !bDocs.some((t) => t.includes("hi")); i++) await tick();

    expect(bDocs.some((t) => t.includes("hi"))).toBe(true);
    await act(async () => { rootA.unmount(); rootB.unmount(); });
  });
});

describe("useCollab presence (jsdom)", () => {
  it("a remote participant's cursor position reaches another hook's presence state", async () => {
    const hub = new CollabHub(blankProvider);
    const factory = makeFakeSocketFactory(hub);
    let setPresenceA: ((p: unknown) => void) | null = null;
    let readyA = false;
    const bPresence: Record<string, unknown>[] = [];

    function A() {
      const s = useCollab({ url: "ws://x", docId: "shared", clientId: "a", createSocket: factory });
      setPresenceA = s.setPresence as unknown as (p: unknown) => void;
      readyA = s.ready;
      return null;
    }
    function B() {
      const s = useCollab({ url: "ws://x", docId: "shared", clientId: "b", createSocket: factory });
      bPresence.push(s.presence);
      return null;
    }
    const rootA = createRoot(document.createElement("div"));
    const rootB = createRoot(document.createElement("div"));
    await act(async () => { rootA.render(createElement(A)); rootB.render(createElement(B)); });
    for (let i = 0; i < 10 && !readyA; i++) await tick();

    // A broadcasts a cursor position; B's presence state should reflect it.
    await act(async () => { setPresenceA!({ anchor: { blockId: 1, runId: 2, offset: 3 } }); });
    for (let i = 0; i < 10 && !bPresence.some((p) => Object.keys(p).length > 0); i++) await tick();

    const latest = bPresence[bPresence.length - 1];
    expect(Object.keys(latest).length).toBeGreaterThan(0); // B knows a remote cursor exists
    const pos = Object.values(latest)[0] as { anchor?: { offset: number } } | null;
    expect(pos?.anchor?.offset).toBe(3);
    await act(async () => { rootA.unmount(); rootB.unmount(); });
  });
});
