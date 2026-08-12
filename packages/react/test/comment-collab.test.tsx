// @vitest-environment jsdom
/**
 * #128 — adding a comment in a SHARED document.
 *
 * `api.addComment` opened with `if (collabRef.current?.submitOp) return false`,
 * which made the `commentRun` intent path below it unreachable: in a room the
 * comment shortcut, the toolbar Comment button, the Review tab's New Comment
 * and the AI's commentRun op all returned false and left the document
 * untouched. `commentRun` is a registered, replicated intent, so the early
 * return was the whole bug.
 *
 * The assertion that matters is byte convergence AFTER the echo settles: the
 * comment must exist on the author's replica AND on a second replica that
 * never ran the call, with both replicas serializing identically.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { CollabEditor } from "../src/collab.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";
import type { DocxViewApi } from "../src/index.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";

const DOCUMENT =
  `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:body><w:p><w:r><w:t xml:space="preserve">Comment the cat here</w:t></w:r></w:p></w:body></w:document>`;

const BYTES = zipSync({
  "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
  "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
  "word/document.xml": strToU8(DOCUMENT),
});

const provider: DocProvider = { load: () => BYTES };

let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `k${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => setTimeout(fn, delayMs);
  return () => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = {
      id: `${ns}c${n++}`,
      send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))),
    };
    let opened = false;
    return {
      send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") ls.push(cb as never);
        else if (!opened) { opened = true; (cb as () => void)(); }
      },
    } as unknown as WebSocket;
  };
}

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 40) { for (let i = 0; i < n; i++) await tick(); }

async function mount(hub: CollabHub, docId: string, clientId: string) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(CollabEditor, {
      url: "ws://x",
      docId,
      clientId,
      commentAuthor: "Alex Pickett",
      toolbar: false,
      createSocket: factoryFor(hub),
      onReady: (a: DocxViewApi) => { seen.api = a; },
      onSession: (s: { doc: DocxDocument | null }) => { if (s.doc) seen.doc = s.doc; },
    }));
  });
  for (let i = 0; i < 40 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  return {
    api: () => seen.api!,
    doc: () => seen.doc!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

/** Everything the replica would persist, as one string. */
function bytes(doc: DocxDocument): string {
  return doc.editableRoots().map((r) => serializeXml(r)).join("\n");
}

describe("addComment in a shared document (#128)", () => {
  it("rides the wire and converges byte for byte on every replica", async () => {
    const hub = new CollabHub(provider);
    const alice = await mount(hub, "comments", "alice");
    const bob = await mount(hub, "comments", "bob");
    await settle();

    await act(async () => { alice.api().find("cat"); });
    let added: boolean | undefined;
    await act(async () => { added = alice.api().addComment("look here"); });
    expect(added).toBe(true); // the bug returned false and did nothing
    await settle();

    // The comment landed on the author's replica...
    expect(alice.doc().comments.length).toBe(1);
    expect(alice.doc().comments[0].text).toContain("look here");
    // ...on the replica that never made the call...
    expect(bob.doc().comments.length).toBe(1);
    expect(bob.doc().comments[0].text).toContain("look here");
    // ...and both wrote identical XML (carried provenance, no local fork).
    expect(bytes(bob.doc())).toBe(bytes(alice.doc()));
    expect(bytes(alice.doc())).toContain("commentRangeStart");

    await alice.unmount();
    await bob.unmount();
  });

  it("the sibling comment operations never had the early return: resolve and edit converge too", async () => {
    const hub = new CollabHub(provider);
    const alice = await mount(hub, "comment-siblings", "alice");
    const bob = await mount(hub, "comment-siblings", "bob");
    await settle();

    await act(async () => { alice.api().find("cat"); });
    await act(async () => { alice.api().addComment("first pass"); });
    await settle();
    const id = alice.doc().comments[0].id;

    let edited: boolean | undefined;
    let resolved: boolean | undefined;
    await act(async () => { edited = alice.api().editComment(id, "second pass"); });
    await settle();
    await act(async () => { resolved = alice.api().resolveComment(id, true); });
    await settle();
    expect(edited).toBe(true);
    expect(resolved).toBe(true);

    expect(bob.doc().comments[0].text).toContain("second pass");
    expect(bob.doc().comments[0].resolved).toBe(true);
    expect(bytes(bob.doc())).toBe(bytes(alice.doc()));

    await alice.unmount();
    await bob.unmount();
  });
});
