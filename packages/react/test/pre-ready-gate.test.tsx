// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";

/**
 * NO EDITABLE SURFACE BEFORE THE SESSION IS READY (B13).
 *
 * The user's decision for large-document joins is explicit: "a fresh client
 * can sit and load and be blocked from editing until caught up thats fine."
 * That is what makes the connection's pre-ready submit guard acceptable — a
 * submit arriving before the welcome cannot be honoured (no replica to apply
 * against, no confirmed seq to base on, no key to seal with in E2EE) and is
 * dropped, now loudly (see collab's submit-before-ready pins).
 *
 * The two halves have to agree, and this pins the editor half: while the
 * welcome is outstanding, CollabEditor renders the connecting notice and NO
 * DocxView, so there is nothing for a user to type into and nothing that
 * could reach the dropping guard. When the welcome lands, the editor appears
 * and editing works normally.
 *
 * Why pin something that already holds: the gate is one `if` in a component
 * with several early returns, and the failure it prevents is invisible —
 * keystrokes that vanish with the document looking fine. B13 spent a swarm
 * scenario and a night of hunting on a loss that presented as data loss.
 */

function docWith(text: string): Uint8Array {
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(xml),
  });
}

const provider: DocProvider = { load: () => docWith("HELLOWORLD") };

/** Socket factory whose SERVER-TO-CLIENT direction can be held, so the
 * welcome stays outstanding for as long as the test wants — the join window a
 * large or encrypted document opens for real. */
function heldFactory(hub: CollabHub): { factory: (url: string) => WebSocket; release: () => void } {
  let n = 0;
  const held: unknown[] = [];
  let holding = true;
  const listeners: ((ev: { data: unknown }) => void)[] = [];
  const factory = (_url: string): WebSocket => {
    const conn: Connection = {
      id: `c${n++}`,
      send: (m: ServerMessage) => {
        const data = JSON.stringify(m);
        if (holding) held.push(data);
        else listeners.forEach((l) => l({ data }));
      },
    };
    let opened = false;
    return {
      send: (d: string) => {
        void hub.handle(conn, JSON.parse(d));
      },
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") listeners.push(cb as never);
        else if (!opened) {
          opened = true;
          (cb as () => void)();
        }
      },
    } as unknown as WebSocket;
  };
  return {
    factory,
    release: () => {
      holding = false;
      for (const data of held.splice(0)) listeners.forEach((l) => l({ data }));
    },
  };
}

async function tick(): Promise<void> {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, 5));
  });
}

describe("the editor is not editable until the session is ready", () => {
  it("shows the connecting notice and mounts no editor while the welcome is outstanding", async () => {
    const hub = new CollabHub(provider);
    const { factory, release } = heldFactory(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(
        createElement(CollabEditor, { url: "ws://x", docId: "d", clientId: "a", createSocket: factory }),
      );
    });
    for (let i = 0; i < 5; i++) await tick();

    // THE GATE: no document surface exists yet, so there is nothing to type
    // into and no path from a keystroke to the connection's dropping guard.
    expect(container.textContent, "the connecting notice is what a joiner sees").toContain("Connecting");
    expect(container.textContent, "no document content is painted").not.toContain("HELLO");
    expect(
      container.querySelectorAll(".dxw-page").length,
      "DocxView is not mounted before ready",
    ).toBe(0);
    expect(
      container.querySelectorAll('[contenteditable="true"]').length,
      "nothing editable is on screen",
    ).toBe(0);

    // …and once the welcome lands, the editor appears and paints.
    release();
    for (let i = 0; i < 20 && !container.textContent?.includes("HELLO"); i++) await tick();
    expect(container.textContent, "the document paints once ready").toContain("HELLO");
    expect(container.textContent).not.toContain("Connecting");
    expect(container.querySelectorAll(".dxw-page").length, "DocxView mounted").toBeGreaterThan(0);

    await act(async () => {
      root.unmount();
    });
  });
});
