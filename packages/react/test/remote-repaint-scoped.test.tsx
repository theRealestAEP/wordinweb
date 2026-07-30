// @vitest-environment jsdom
/**
 * THE OWNER'S SCENARIO, at test scale: user X watches a large document while
 * user Y edits it elsewhere. Before the dirty-scope threading, EVERY remote
 * intent forced a whole-document relayout — past 50 pages an async one behind
 * an inert (input-blocking) container. Now a remote text edit must repaint as
 * ONE incremental paragraph relayout, must never set the input-blocking busy
 * flag, and the view must stay converged with the server.
 *
 * The assertions are causal, not wall-clock: canvas measureText calls per
 * remote keystroke (the actual O(document) work), and the layout engine's own
 * incremental stats (__dxwPerf.incr: fast path taken, no fallback).
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { serializeXml, type Paragraph, type Run, type DocxDocument } from "@wordinweb/core";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";

const PARAS = 1400; // enough pages to be past BACKGROUND_LAYOUT_PAGE_THRESHOLD (50)

function bigDocBytes(paras: number): Uint8Array {
  const para = (i: number) =>
    `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i}: the quick brown fox jumps over the lazy dog while the committee deliberates at length. </w:t></w:r></w:p>`;
  let body = "";
  for (let i = 0; i < paras; i++) body += para(i);
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
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
    "word/document.xml": strToU8(documentXml),
  });
}

// Count canvas measureText calls — the causal cost of layout work. Wraps the
// fake context factories the shared setup installed (the measurer prefers
// OffscreenCanvas, so wrap both).
let measureCalls = 0;
function wrapGetContext(proto: { getContext: (...a: unknown[]) => unknown }): void {
  const orig = proto.getContext;
  proto.getContext = function (...args: unknown[]) {
    const ctx = orig.apply(this, args) as { measureText?: (t: string) => unknown } | null;
    if (ctx && typeof ctx.measureText === "function") {
      const m = ctx.measureText.bind(ctx);
      ctx.measureText = (text: string) => {
        measureCalls++;
        return m(text);
      };
    }
    return ctx;
  };
}
wrapGetContext(HTMLCanvasElement.prototype as never);
wrapGetContext((globalThis as { OffscreenCanvas: { prototype: never } }).OffscreenCanvas.prototype);

// jsdom has no layout: clientHeight is 0 everywhere, so the virtualizer would
// mount only a tail window and the edited page's DOM would never exist. Give
// every element a viewport so the first pages mount (mirrors the local-typing
// bench, scripts/bench-local-typing.mjs).
Object.defineProperty(HTMLElement.prototype, "clientHeight", {
  get() {
    return 1200;
  },
  configurable: true,
});

function factoryFor(hub: CollabHub) {
  let n = 0;
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = {
      id: `rr${n++}`,
      send: (m: ServerMessage) => ls.forEach((l) => l({ data: JSON.stringify(m) })),
    };
    let opened = false;
    return {
      send: (d: string) => {
        void hub.handle(conn, JSON.parse(d));
      },
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") ls.push(cb as never);
        else if (!opened) {
          opened = true;
          (cb as () => void)();
        }
      },
    } as unknown as WebSocket;
  };
}

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}
async function until(cond: () => boolean, label: string, tries = 400) {
  for (let i = 0; i < tries; i++) {
    if (cond()) return;
    await tick();
  }
  if (!cond()) throw new Error(`timeout: ${label}`);
}

/** The live collab session object DocxView received, via the fiber tree. */
function sessionOf(container: HTMLElement): { doc: DocxDocument | null } {
  const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
  const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
  let guard = 0;
  while (stack.length && guard++ < 20000) {
    const f = stack.pop() as {
      memoizedProps?: { collab?: { doc?: DocxDocument | null; presence?: unknown } };
      child?: unknown;
      sibling?: unknown;
    } | null;
    if (!f) continue;
    const c = f.memoizedProps?.collab;
    if (c && c.presence && typeof c.presence === "object") return c as never;
    if (f.child) stack.push(f.child);
    if (f.sibling) stack.push(f.sibling);
  }
  throw new Error("collab session not found in fiber tree");
}

function addrOf(doc: DocxDocument, i: number): { blockId: number; runId: number; len: number } {
  const para = doc.sections[0].blocks[i] as Paragraph;
  const run = para.children[0] as Run;
  const t = run.content.find((c) => c.kind === "text")!.srcT!;
  const ids = doc.stableIds!;
  return { blockId: ids.idOf(para.src!)!, runId: ids.idOf(run.src!)!, len: t.text.length };
}

type Perf = { incr?: { fallbackReason: string; hintFastPath: boolean; blocksLaid: number; pageShift: number } };

describe("remote edits repaint scoped on a large document (owner scenario)", () => {
  it("far remote keystrokes cost one paragraph, never block input, and converge", async () => {
    (globalThis as { __dxwPerf?: Perf }).__dxwPerf = {};
    const hub = new CollabHub({ load: () => bigDocBytes(PARAS) } as DocProvider);
    const factory = factoryFor(hub);

    // X — the watcher: a full CollabEditor (DocxView) on the big document.
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    measureCalls = 0;
    await act(async () => {
      root.render(
        createElement(CollabEditor, { url: "ws://x", docId: "big", clientId: "watcher", createSocket: factory }),
      );
    });
    await until(() => !!container.querySelector(".dxw-page"), "watcher paints");
    const mountMeasures = measureCalls; // the cost of one WHOLE-document layout
    const pages = container.querySelectorAll(".dxw-page").length;
    // The scenario must actually be past the background-layout threshold, or
    // the input-blocking assertion below is vacuous.
    expect(pages).toBeGreaterThan(50);

    // Record any input-blocking background layout on the watcher.
    let busySeen = 0;
    const mo = new MutationObserver((records) => {
      for (const r of records) {
        if ((r.target as HTMLElement).hasAttribute?.("data-dxw-layout-busy")) busySeen++;
      }
    });
    mo.observe(container, { subtree: true, attributes: true, attributeFilter: ["data-dxw-layout-busy", "inert"] });

    // Y — the editor: a bare connection typing into paragraph 2 (page 1),
    // far from wherever the watcher's viewport work would matter.
    const editor = new CollabConnection(createWebSocketTransport(factory("ws://x") as never), "typist");
    editor.join("big");
    await until(() => editor.ready, "editor welcome");
    const addr = addrOf(editor.doc!, 2);

    const session = sessionOf(container);
    const perf = (globalThis as { __dxwPerf: Perf }).__dxwPerf;
    const perKey: number[] = [];
    const KEYS = 8;
    for (let k = 0; k < KEYS; k++) {
      const marker = `X${k}Q`;
      const m0 = measureCalls;
      perf.incr = undefined;
      await act(async () => {
        editor.submit({
          kind: "insertText",
          at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len + "X0Q".length * k },
          text: marker,
        } as never);
      });
      // The watcher must actually PAINT the remote edit (freshness first).
      await until(() => !!container.textContent?.includes(marker), `watcher paints ${marker}`);
      perKey.push(measureCalls - m0);
      // The repaint went through the incremental engine's hint fast path —
      // one dirty block, no fallback to a full pass.
      expect(perf.incr?.fallbackReason).toBe("");
      expect(perf.incr?.hintFastPath).toBe(true);
      // Relayout is O(the edited page) — resume points are page tops, so the
      // dirty paragraph's whole page re-lays (~23 blocks here) — never
      // O(document).
      expect(perf.incr!.blocksLaid).toBeLessThan(PARAS / 10);
    }
    const median = [...perKey].sort((a, b) => a - b)[Math.floor(perKey.length / 2)];
    // CAUSAL bound: a remote keystroke must not re-measure the document.
    // Full layout measured `mountMeasures` (O(document)); a scoped repaint
    // re-breaks one paragraph. 20× headroom keeps this loose enough for
    // cache-warmth noise while failing the pre-fix behavior by orders of
    // magnitude (pre-fix: every keystroke ~= mountMeasures).
    expect(mountMeasures).toBeGreaterThan(5000); // sanity: the ratio means something
    expect(median).toBeLessThan(mountMeasures / 20);
    // No remote text edit may block input behind the async layout path.
    await tick(80); // let any (wrongly) queued background layout surface
    expect(busySeen).toBe(0);
    expect(container.inert).toBeFalsy();

    // STRUCTURAL: a remote Enter (splitParagraph) genuinely reflows forward.
    // It stays incremental (insertion fast path + shifted retained suffix)
    // and still must not block input; its honest cost is reported, bounded
    // only against the full-document cost.
    const [newBlockId, newRunId] = editor.allocIds(2);
    const mSplit0 = measureCalls;
    perf.incr = undefined;
    const blocksBefore = (sessionOf(container).doc as DocxDocument).sections[0].blocks.length;
    await act(async () => {
      editor.submit({
        kind: "splitParagraph",
        at: { blockId: addr.blockId, runId: addr.runId, offset: 3 },
        newBlockId,
        newRunId,
      } as never);
    });
    await until(
      () => (sessionOf(container).doc as DocxDocument).sections[0].blocks.length === blocksBefore + 1,
      "watcher applies split",
    );
    await tick(80);
    const splitMeasures = measureCalls - mSplit0;
    expect(perf.incr?.fallbackReason).toBe("");
    expect(perf.incr?.hintFastPath).toBe(true);
    expect(splitMeasures).toBeLessThan(mountMeasures / 10);
    expect(busySeen).toBe(0);

    // CONVERGENCE — the view optimization must never desync the model. The
    // editor's replica, the watcher's rendered doc, and a FRESH third client
    // (whose welcome is built from the server's canonical session) must be
    // byte-identical.
    const late = new CollabConnection(createWebSocketTransport(factory("ws://x") as never), "late");
    late.join("big");
    await until(() => late.ready, "late welcome");
    const watcherXml = serializeXml((sessionOf(container).doc as DocxDocument).docRoot);
    expect(watcherXml).toBe(serializeXml(editor.doc!.docRoot));
    expect(watcherXml).toBe(serializeXml(late.doc!.docRoot));
    // And the pixels followed the model: the last marker is on the page.
    expect(container.textContent).toContain(`X${KEYS - 1}Q`);

    console.log(
      `STRESS-METRIC remote-repaint-scoped paragraphs=${PARAS} pages=${pages} keys=${KEYS} ` +
        `mountMeasures=${mountMeasures} measurePerKeyMedian=${median} measurePerKeyMax=${Math.max(...perKey)} ` +
        `splitMeasures=${splitMeasures} busySeen=${busySeen}`,
    );

    mo.disconnect();
    await act(async () => {
      root.unmount();
    });
  }, 240_000);
});
