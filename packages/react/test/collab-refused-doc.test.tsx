// @vitest-environment jsdom
//
// A refused session and the document it may leave behind.
//
// The split under test is the whole point:
//  - a POST-LIVE refusal (idle-timeout / session-expired) arrives with the
//    document in hand → it stays on screen, READ-ONLY, with the host's
//    refusal content overlaid;
//  - a FIRST-CONNECT refusal (room-full, already-open, …) arrives before
//    anything loaded → the refusal is the whole page, and no phantom empty
//    editor is faked behind it.
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabEditor, type CollabSession } from "../src/collab.js";
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

/** Hub-backed socket factory that also CAPTURES the server-side Connection,
 * so the test can push a server frame (the kick) at the client directly. */
function factoryFor(hub: CollabHub, conns: Connection[]) {
  let n = 0;
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `c${n++}`, send: (m: ServerMessage) => ls.forEach((l) => l({ data: JSON.stringify(m) })) };
    conns.push(conn);
    let opened = false;
    return { send: (d: string) => { void hub.handle(conn, JSON.parse(d)); },
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}

/** Socket that answers EVERY client frame with a refusal — the first-connect
 * refusal shape (nothing was ever loaded). */
function refusingFactory(reason: string) {
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    return {
      send: (_d: string) => queueMicrotask(() => ls.forEach((l) => l({ data: JSON.stringify({ t: "refused", reason }) }))),
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else (cb as () => void)(); },
    } as unknown as WebSocket;
  };
}

async function tick() { await act(async () => { await new Promise<void>((r) => setTimeout(r, 5)); }); }

const refusedContent = (reason: string, ctx?: { docVisible: boolean }) =>
  createElement("div", { "data-testid": "refusal", "data-doc-visible": String(ctx?.docVisible ?? false) }, `ended:${reason}`);

describe("refused WITH a document (post-live kick)", () => {
  it("keeps the document on screen behind the refusal content, read-only", async () => {
    const hub = new CollabHub(provider);
    const conns: Connection[] = [];
    let session: CollabSession | null = null;
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(CollabEditor, {
        url: "ws://x", docId: "d", clientId: "a",
        createSocket: factoryFor(hub, conns),
        onSession: (s) => { session = s; },
        refusedContent,
      }));
    });
    for (let i = 0; i < 15 && !container.textContent?.includes("HELLO"); i++) await tick();
    expect(container.textContent).toContain("HELLOWORLD");
    // PRECONDITION that keeps the read-only assertion below honest: while the
    // session is live and editable, the attached editor's IME textarea (the
    // text-input surface) exists. If this ever stops being the editable
    // marker, the "no textarea" assertion would go vacuous — this line is
    // what would catch that.
    expect(container.querySelector("textarea")).toBeTruthy();
    expect((session as unknown as CollabSession).writesBlocked).toBe(false);

    // The server ends the session: the idle-timeout kick frame.
    await act(async () => { conns[0].send({ t: "refused", reason: "idle-timeout" }); });
    for (let i = 0; i < 15 && !container.querySelector('[data-testid="refusal"]'); i++) await tick();
    // Wait for the read-only DocxView (re-parsed from bytes) to paint too.
    for (let i = 0; i < 15 && !container.textContent?.includes("HELLOWORLD"); i++) await tick();

    const refusal = container.querySelector<HTMLElement>('[data-testid="refusal"]');
    expect(refusal).toBeTruthy();
    expect(refusal!.textContent).toBe("ended:idle-timeout");
    // The host was told the document is visible behind its content.
    expect(refusal!.dataset.docVisible).toBe("true");
    // The document is still painted, behind an overlay-positioned refusal.
    expect(container.textContent).toContain("HELLOWORLD");
    expect(container.querySelector(".dxw-collab-refused-overlay")).toBeTruthy();
    expect(container.querySelector(".dxw-page")).toBeTruthy();
    // READ-ONLY, asserted directly: no editor is attached, so its IME
    // textarea (the only way keystrokes enter a document) does not exist,
    // and no element is focusable as an editing surface.
    expect(container.querySelector("textarea")).toBeNull();
    expect(container.querySelectorAll('[tabindex="0"]').length).toBe(0);
    // And the single write-gate predicate agrees: refusal folds into it.
    expect((session as unknown as CollabSession).writesBlocked).toBe(true);
    await act(async () => { root.unmount(); });
  });
});

describe("refused WITHOUT a document (first-connect refusal)", () => {
  it("renders the full-screen refusal with no phantom empty editor", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    await act(async () => {
      root.render(createElement(CollabEditor, {
        url: "ws://x", docId: "d", clientId: "a",
        createSocket: refusingFactory("room-full"),
        refusedContent,
      }));
    });
    for (let i = 0; i < 15 && !container.querySelector('[data-testid="refusal"]'); i++) await tick();

    const refusal = container.querySelector<HTMLElement>('[data-testid="refusal"]');
    expect(refusal).toBeTruthy();
    expect(refusal!.textContent).toBe("ended:room-full");
    // The host was told there is NO document behind its content.
    expect(refusal!.dataset.docVisible).toBe("false");
    // Full-screen refusal markup, exactly as before the overlay existed…
    expect(container.querySelector(".dxw-collab-refused")).toBeTruthy();
    expect(container.querySelector(".dxw-collab-refused-overlay")).toBeNull();
    // …and no faked empty page or editor surface behind it.
    expect(container.querySelector(".dxw-page")).toBeNull();
    expect(container.querySelector("textarea")).toBeNull();
    await act(async () => { root.unmount(); });
  });
});
