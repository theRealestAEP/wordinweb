// @vitest-environment jsdom
/**
 * REMOTE SELECTION HIGHLIGHT (sender side, live): selecting text in one
 * CollabEditor must put `ranges` on the presence payload that leaves the
 * socket — in the WIRE basis the receiver decodes — and the peer must both
 * receive them and paint a `.dxw-presence-selection` rect.
 *
 * Same harness shape as presence-live-flow: keys (not clicks) drive the
 * selection, because jsdom's degenerate layout collapses every click to the
 * same spot, and the assertions are on the wire numbers rather than pixels.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { zipSync, strToU8 } from "fflate";

interface WirePresence {
  anchor: { blockId: number; runId: number; offset: number };
  ranges?: { blockId: number; runId: number; start: number; end: number }[];
}
interface SentPresence { sock: number; position: WirePresence | null }

function docWith(text: string): Uint8Array {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
}

let factorySeq = 0;
function factoryFor(hub: CollabHub, sent: SentPresence[], received: SentPresence[], delayMs = 2) {
  const ns = `s${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => (delayMs > 0 ? setTimeout(fn, delayMs) : fn());
  return (_url: string) => {
    const myn = n++;
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `${ns}c${myn}`, send: (m: ServerMessage) => {
      if ((m as { t?: string }).t === "presence") received.push({ sock: myn, position: (m as { position?: WirePresence | null }).position ?? null });
      defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) })));
    } };
    let opened = false;
    return {
      send: (d: string) => {
        const msg = JSON.parse(d);
        if (msg?.t === "presence") sent.push({ sock: myn, position: msg.position });
        defer(() => { void hub.handle(conn, msg); });
      },
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

async function mount(hub: CollabHub, sent: SentPresence[], received: SentPresence[], docId: string, clientId: string) {
  const factory = factoryFor(hub, sent, received);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  await act(async () => {
    root.render(createElement(CollabEditor, { url: "ws://x", docId, clientId, createSocket: factory }));
  });
  for (let i = 0; i < 40 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const target = () =>
    (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
  const keys = async (seq: (string | { key: string; shift: boolean })[]) => {
    await act(async () => {
      for (const k of seq) {
        const { key, shift } = typeof k === "string" ? { key: k, shift: false } : k;
        target().dispatchEvent(new KeyboardEvent("keydown", { key, shiftKey: shift, bubbles: true, cancelable: true }));
        await new Promise((r) => setTimeout(r, 3));
      }
    });
  };
  const clickPage = async () => {
    const surface = container.querySelector<HTMLElement>(".dxw-page")!.firstElementChild as HTMLElement;
    await act(async () => {
      const o = { bubbles: true, cancelable: true, clientX: 20, clientY: 12, button: 0 };
      surface.dispatchEvent(new MouseEvent("mousedown", o));
      surface.dispatchEvent(new MouseEvent("mouseup", o));
    });
    await tick();
  };
  const unmount = async () => { await act(async () => { root.unmount(); }); };
  return { container, keys, clickPage, unmount };
}

describe("live selection presence (A selects -> B receives ranges and paints them)", () => {
  it("a selection emits wire ranges alongside the caret, and the peer draws the highlight", async () => {
    const hub = new CollabHub({ load: () => docWith("HELLOWORLD") } as DocProvider);
    const sent: SentPresence[] = [];
    const received: SentPresence[] = [];
    const A = await mount(hub, sent, received, "shared", "alice");
    const B = await mount(hub, sent, received, "shared", "bob");
    for (let i = 0; i < 20 && !B.container.textContent?.includes("HELLO"); i++) await tick();

    await A.clickPage();
    await A.keys(["Home"]);
    // Caret-only presence first: no selection, so no ranges (backward shape).
    const caretOnly = sent.filter((s) => s.position).slice(-1)[0];
    expect(caretOnly?.position?.ranges, "a bare caret carries no ranges").toBeUndefined();

    // Shift+End selects the whole run: wire range [0, 10) of the only w:t.
    await A.keys([{ key: "End", shift: true }]);
    await tick();
    const withRanges = sent.filter((s) => s.position?.ranges).slice(-1)[0];
    expect(withRanges, "selecting emits a presence payload with ranges").toBeTruthy();
    const ranges = withRanges!.position!.ranges!;
    expect(ranges).toHaveLength(1);
    expect(ranges[0].start).toBe(0);
    expect(ranges[0].end).toBe(10); // "HELLOWORLD"
    // Ranges address the SAME run the caret does — one keyspace, one basis.
    expect(ranges[0].runId).toBe(withRanges!.position!.anchor.runId);
    expect(ranges[0].blockId).toBe(withRanges!.position!.anchor.blockId);

    // B's socket receives them, and B paints a highlight rect for alice.
    for (let i = 0; i < 40; i++) await tick();
    const bRanges = received.filter((r) => r.position?.ranges).slice(-1)[0];
    expect(bRanges?.position?.ranges?.[0]).toEqual(ranges[0]);
    const box = B.container.querySelector<HTMLElement>(".dxw-presence-selection");
    expect(box, "B paints A's selection highlight").toBeTruthy();
    expect(box!.dataset.participant).toBe("alice");
    expect(B.container.querySelector(".dxw-presence-caret"), "and still A's caret").toBeTruthy();

    // Collapsing the selection clears the highlight on B (ranges drop out).
    await A.keys(["Home"]);
    for (let i = 0; i < 40; i++) await tick();
    const last = sent.filter((s) => s.position).slice(-1)[0];
    expect(last?.position?.ranges, "collapsed caret sends no ranges").toBeUndefined();
    expect(B.container.querySelector(".dxw-presence-selection"), "B's highlight is gone").toBeNull();

    await A.unmount();
    await B.unmount();
  });
});
