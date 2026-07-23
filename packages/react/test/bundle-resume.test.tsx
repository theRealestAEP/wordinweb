// @vitest-environment jsdom
/**
 * Bundle persistence + resume through the REACT layer (plan doc 12 §4/§5):
 * useCollab({ store }) must persist the confirmed state (flush on unmount —
 * the "close the laptop" path), resume seamlessly into a live session
 * (case 1), and on an epoch change (case 2) adopt the server state while
 * preserving the superseded bundle as a draft. Driven against the real
 * CollabHub over the same async factory transport the fidelity tests use.
 */
import { describe, expect, it } from "vitest";
import { createElement, act, useImperativeHandle, forwardRef, type Ref } from "react";
import { createRoot } from "react-dom/client";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import { InMemoryBundleStore } from "@wordinweb/collab/client";
import { useCollab, type CollabSession } from "../src/collab.js";

const provider: DocProvider = { load: () => blankDocxBytes() };

let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 2) {
  const ns = `f${factorySeq++}-`;
  let n = 0;
  const defer = (fn: () => void) => (delayMs > 0 ? setTimeout(fn, delayMs) : fn());
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = { id: `${ns}c${n++}`, send: (m: ServerMessage) => defer(() => ls.forEach((l) => l({ data: JSON.stringify(m) }))) };
    let opened = false;
    return { send: (d: string) => defer(() => { void hub.handle(conn, JSON.parse(d)); }),
      addEventListener: (t: "message" | "open", cb: never) => { if (t === "message") ls.push(cb as never); else if (!opened) { opened = true; (cb as () => void)(); } },
    } as unknown as WebSocket;
  };
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function until(pred: () => boolean, label: string) {
  for (let i = 0; i < 100 && !pred(); i++) await tick();
  if (!pred()) throw new Error(`timeout: ${label}`);
}

/** Headless probe: runs useCollab and hands the live session out via ref. */
const Probe = forwardRef(function Probe(
  props: { factory: (u: string) => WebSocket; docId: string; clientId: string; store: InMemoryBundleStore },
  ref: Ref<{ session: CollabSession }>,
) {
  const session = useCollab({
    url: "ws://x",
    docId: props.docId,
    clientId: props.clientId,
    createSocket: props.factory, // stable per mount — an inline factory would re-run the effect every render
    store: props.store,
  });
  useImperativeHandle(ref, () => ({ session }), [session]);
  return null;
});

async function mountProbe(hub: CollabHub, docId: string, clientId: string, store: InMemoryBundleStore) {
  const factory = factoryFor(hub);
  const holder: { current: { session: CollabSession } | null } = { current: null };
  const div = document.createElement("div");
  const root = createRoot(div);
  await act(async () => {
    root.render(createElement(Probe, { factory, docId, clientId, store, ref: (v: never) => (holder.current = v) }));
  });
  await until(() => holder.current?.session.ready === true, "session ready");
  return {
    get session() { return holder.current!.session; },
    unmount: async () => { await act(async () => root.unmount()); await tick(20); },
  };
}

const docText = (s: CollabSession): string => {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(s.doc!.docRoot as never);
};
const insertAt = (offset: number, text: string) =>
  ({ kind: "insertText", at: { blockId: 1, runId: 2, offset }, text }) as never;

describe("useCollab({ store }) — persistence + resume (doc 12 §4/§5)", () => {
  it("persists on unmount; a later mount resumes and NEW edits are not self-deduped (watermark)", async () => {
    const hub = new CollabHub(provider);
    const store = new InMemoryBundleStore();

    // Session A: type, close the tab.
    const a = await mountProbe(hub, "d", "alice", store);
    await act(async () => a.session.submitOp(insertAt(0, "hello")));
    await until(() => a.session.version > 1, "echo");
    await a.unmount();
    const bundle = await store.get("d");
    expect(bundle).toBeTruthy();
    expect(bundle!.clientSeq).toBe(1);
    expect(bundle!.genesisId).toBeTruthy();

    // Later: same browser (same store, same clientId) reopens the doc.
    const b = await mountProbe(hub, "d", "alice", store);
    expect(b.session.epochChanged).toBeNull(); // case 1: same epoch, seamless
    // Watermark regression: this NEW edit must sequence, not dedup away.
    await act(async () => b.session.submitOp(insertAt(5, "!")));
    await until(() => b.session.version > 1, "second echo");
    expect(docText(b.session)).toContain("hello!");
    await b.unmount();
  });

  it("case 2: resume into a re-seeded epoch adopts server state and saves the old bundle as a draft", async () => {
    const store = new InMemoryBundleStore();

    // Epoch 1 on hub1: alice works, tab closes, bundle persisted.
    const hub1 = new CollabHub(provider);
    const a = await mountProbe(hub1, "d", "alice", store);
    await act(async () => a.session.submitOp(insertAt(0, "mine")));
    await until(() => a.session.version > 1, "echo1");
    await a.unmount();
    const oldBundle = (await store.get("d"))!;

    // The server dies (zero custody: hub2 knows nothing); bob re-seeds the
    // SAME docId from his copy — a NEW epoch with different content.
    const hub2 = new CollabHub(null, undefined, undefined, () => 0, () => "g_bob");
    hub2.seed("d", blankDocxBytes());
    const bobStore = new InMemoryBundleStore();
    const b = await mountProbe(hub2, "d", "bob", bobStore);
    await act(async () => b.session.submitOp(insertAt(0, "theirs")));
    await until(() => b.session.version > 1, "echo2");

    // Alice returns with her old-epoch bundle.
    const a2 = await mountProbe(hub2, "d", "alice", store);
    await until(() => a2.session.epochChanged !== null, "epoch change surfaced");
    expect(a2.session.epochChanged!.from).toBe(oldBundle.genesisId);
    expect(a2.session.epochChanged!.to).toBe("g_bob");
    // She sees the server's state (bob's), not a silent merge with hers.
    expect(docText(a2.session)).toContain("theirs");
    expect(docText(a2.session)).not.toContain("mine");
    // Her superseded copy is preserved as a draft, byte-for-byte.
    await until(() => store.writes >= 2, "draft written");
    const draft = await store.get(`d#draft-${oldBundle.genesisId}`);
    expect(draft).toBeTruthy();
    expect(Buffer.from(draft!.confirmedBytes).equals(Buffer.from(oldBundle.confirmedBytes))).toBe(true);
    await a2.unmount();
  });
});
