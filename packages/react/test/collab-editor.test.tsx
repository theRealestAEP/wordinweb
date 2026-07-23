// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";

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
    const conn: Connection = { id: `c${n++}`, send: (m: ServerMessage) => ls.forEach((l) => l({ data: JSON.stringify(m) })) };
    let opened = false;
    return { send: (d: string) => { void hub.handle(conn, JSON.parse(d)); },
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}
async function tick() { await act(async () => { await new Promise<void>((r) => setTimeout(r, 5)); }); }

describe("CollabEditor renders the collab document (jsdom)", () => {
  it("mounts, connects, and paints the reconciled document's text via DocxView", async () => {
    const hub = new CollabHub(provider);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(CollabEditor, { url: "ws://x", docId: "d", clientId: "a", createSocket: factoryFor(hub) }));
    });
    for (let i = 0; i < 15 && !container.textContent?.includes("HELLO"); i++) await tick();
    // The real DocxView painted the collab document's text into the DOM.
    expect(container.textContent).toContain("HELLO");
    expect(container.textContent).not.toContain("Connecting");
    // DocxView produced real DOM structure (not just a placeholder div).
    expect(container.querySelectorAll("*").length).toBeGreaterThan(3);
    await act(async () => { root.unmount(); });
  });
});

import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";

describe("CollabEditor repaints on a remote edit (full inbound loop)", () => {
  it("a second client's insert re-renders the editor's painted text", async () => {
    const hub = new CollabHub(provider);
    const factory = factoryFor(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(CollabEditor, { url: "ws://x", docId: "shared", clientId: "viewer", createSocket: factory }));
    });
    for (let i = 0; i < 15 && !container.textContent?.includes("HELLO"); i++) await tick();
    expect(container.textContent).toContain("HELLOWORLD");

    // A separate client joins the same doc and inserts text.
    const editor = new CollabConnection(createWebSocketTransport(factory("ws://x") as never), "editor");
    editor.join("shared");
    await tick();
    await act(async () => {
      editor.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 5 }, text: "-X-" } as never);
    });
    // The CollabEditor receives the broadcast, reconciles, and REPAINTS.
    for (let i = 0; i < 15 && !container.textContent?.includes("HELLO-X-WORLD"); i++) await tick();
    expect(container.textContent).toContain("HELLO-X-WORLD");
    await act(async () => { root.unmount(); });
  });
});

describe("CollabEditor draws remote presence carets (jsdom)", () => {
  it("renders a caret bar for a remote participant's cursor", async () => {
    const hub = new CollabHub(provider);
    const factory = factoryFor(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(CollabEditor, { url: "ws://x", docId: "shared", clientId: "viewer", createSocket: factory }));
    });
    for (let i = 0; i < 15 && !container.textContent?.includes("HELLO"); i++) await tick();

    // A remote editor joins and broadcasts a cursor position in the run (ids 1/2).
    const editor = new CollabConnection(createWebSocketTransport(factory("ws://x") as never), "peer");
    editor.join("shared");
    await tick();
    await act(async () => { editor.setPresence({ anchor: { blockId: 1, runId: 2, offset: 5 } }); });
    for (let i = 0; i < 15 && !container.querySelector(".dxw-presence-caret"); i++) await tick();

    const caret = container.querySelector<HTMLElement>(".dxw-presence-caret");
    expect(caret).toBeTruthy(); // a remote cursor is painted on the page
    expect(caret!.dataset.participant).toBeTruthy();
    await act(async () => { root.unmount(); });
  });
});
