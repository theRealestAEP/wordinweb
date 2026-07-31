// @vitest-environment jsdom
/**
 * DIAGNOSTIC (repro): live two-window presence — does moving the LOCAL caret
 * re-report presence on EVERY move (cadence), and does the encoded position
 * track the caret's real offset? Symptom under test: the remote caret "does not
 * appear to update as the sender moves" and sits near the line start.
 *
 * jsdom layout is degenerate (all clicks converge, per caret-move-repaint), so
 * we drive the caret with KEYS (arrows/Home/End), which move the model offset
 * deterministically, and assert on the ENCODED wire positions the sender emits
 * — not pixels. That isolates cadence + offset basis from geometry.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { zipSync, strToU8 } from "fflate";

function docWith(text: string): Uint8Array {
  const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
}

// A record of every presence position sent over any socket, tagged by socket #.
interface SentPresence { sock: number; position: { anchor: { blockId: number; runId: number; offset: number } } | null; }

let factorySeq = 0;
function factoryFor(hub: CollabHub, sent: SentPresence[], received: SentPresence[], delayMs = 2) {
  const ns = `f${factorySeq++}-`; // unique per factory so two mounts never collide conn ids
  let n = 0;
  const defer = (fn: () => void) => (delayMs > 0 ? setTimeout(fn, delayMs) : fn());
  return (_url: string) => {
    const myn = n++;
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `${ns}c${myn}`, send: (m: ServerMessage) => {
      if ((m as { t?: string }).t === "presence") received.push({ sock: myn, position: (m as { position?: SentPresence["position"] }).position ?? null });
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
  const keys = async (seq: string[]) => {
    await act(async () => {
      for (const key of seq) {
        target().dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
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
  // Reach the live collab session (presence map) through the React tree.
  const collabState = (): { presence?: Record<string, unknown>; participants?: unknown } | null => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 8000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { presence?: Record<string, unknown> } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const c = f.memoizedProps?.collab;
      if (c && c.presence && typeof c.presence === "object") return c as never;
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    return null;
  };
  const unmount = async () => { await act(async () => { root.unmount(); }); };
  return { container, keys, clickPage, collabState, unmount };
}

describe("live presence cadence + offset basis (two CollabEditors, A -> B)", () => {
  it("A re-reports presence on each caret move, with the offset tracking the caret", async () => {
    const hub = new CollabHub({ load: () => docWith("HELLOWORLD") } as DocProvider);
    const sent: SentPresence[] = [];
    const received: SentPresence[] = [];
    const A = await mount(hub, sent, received, "shared", "alice");
    const B = await mount(hub, sent, received, "shared", "bob");
    // Let both settle / see initial doc.
    for (let i = 0; i < 20 && !B.container.textContent?.includes("HELLO"); i++) await tick();

    // Focus A's caret in the doc, then walk it deterministically with keys.
    await A.clickPage();
    await A.keys(["End"]);          // caret to end of "HELLOWORLD" (offset 10)
    await tick();
    const afterEnd = sent.filter((s) => s.position).slice(-1)[0];
    await A.keys(["Home"]);         // caret to start (offset 0)
    await tick();
    const afterHome = sent.filter((s) => s.position).slice(-1)[0];
    await A.keys(["ArrowRight", "ArrowRight", "ArrowRight", "ArrowRight"]); // offset 4
    await tick();
    const afterRight4 = sent.filter((s) => s.position).slice(-1)[0];

    // Collect the distinct offsets A emitted over the whole walk.
    const offsets = sent.filter((s) => s.position).map((s) => s.position!.anchor.offset);
    void afterEnd; void afterHome;

    // CADENCE: the caret visited end (10), start (0), and mid (4); each distinct
    // stop must have been reported. If presence goes stale, these fail.
    expect(offsets, "End should report offset 10").toContain(10);
    expect(offsets, "Home should report offset 0").toContain(0);
    expect(offsets, "ArrowRight x4 should report offset 4").toContain(4);
    // The LAST reported position must reflect the final caret (offset 4), not a
    // stale early value.
    expect(afterRight4?.position?.anchor.offset, "final reported offset tracks the caret").toBe(4);

    // RELAY (authoritative): B's connection receives EVERY presence A emitted,
    // in order — including the final offset 4. This is the wire-level truth,
    // read at B's socket, so it isolates delivery from React/paint timing.
    for (let i = 0; i < 40; i++) await tick();
    const bRecv = received.filter((r) => r.position).map((r) => r.position!.anchor);
    // B received the whole caret walk (cadence survives end-to-end), not just
    // the first position: this is what "the remote caret tracks the sender"
    // requires at the transport layer.
    expect(bRecv.map((a) => a.offset), "B received the full caret path incl final 4").toEqual([0, 10, 0, 1, 2, 3, 4]);
    // ID PARITY: every relayed runId is A's runId 2 — the sidecar keeps the
    // stable-id meaning identical across replicas (no id drift for this doc).
    expect([...new Set(bRecv.map((a) => a.runId))], "same run id on both replicas").toEqual([afterRight4!.position!.anchor.runId]);

    // B applies presence into React state and paints a caret bar.
    const bPresence = B.collabState()?.presence as Record<string, { anchor: { blockId: number; runId: number; offset: number } } | null>;
    expect(bPresence?.alice, "B has A's presence entry in its collab state").toBeTruthy();
    expect(B.container.querySelector(".dxw-presence-caret"), "B paints A's caret").toBeTruthy();

    await A.unmount();
    await B.unmount();
  });
});
