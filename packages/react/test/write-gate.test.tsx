// @vitest-environment jsdom
/**
 * USER-REPORTED: the read-only banner was on screen saying "edits are paused"
 * WHILE the user typed and their text appeared. The banner annotated; it did
 * not gate. Every one of those keystrokes applied to the local document, was
 * refused by the server, and was healed away — silent data loss that looked
 * like it worked.
 *
 * THE RULE: any state where the server will not accept this client's writes
 * renders the editor in VIEWER MODE, never merely annotated.
 *
 * The load-bearing assertion here is BYTE-UNCHANGED, not "did not sync". The
 * bug was the LOCAL MUTATION — an assertion about syncing passed throughout the
 * period the bug existed, because the refused edit genuinely never reached
 * anyone else. Only the local document's bytes tell the two apart.
 */
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { CollabEditor } from "../src/collab.js";
import { CollabHub, blankDocxBytes, type DocProvider, type Connection, type ServerMessage } from "@wordinweb/server";
import type { DocxDocument } from "@wordinweb/core";

const glob = globalThis as unknown as Record<string, unknown>;
glob.createImageBitmap ??= async () => ({ width: 64, height: 48, close() {} });
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:write-gate";
  URL.revokeObjectURL = () => {};
}
async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }
async function settle(n = 25) { for (let i = 0; i < n; i++) await tick(); }

const provider: DocProvider = { load: () => blankDocxBytes() };
let seq = 0;
function factoryFor(hub: CollabHub, legacy = false) {
  const ns = `w${seq++}-`;
  let n = 0;
  const defer = (fn: () => void) => setTimeout(fn, 2);
  return (_url: string) => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = {
      id: `${ns}c${n++}`,
      send: (m: ServerMessage) => defer(() => {
        // `legacy` simulates a server built before write status existed: the
        // roster arrives with the field ABSENT. That is the branch the
        // fallback contract exists for, and it cannot be produced by the real
        // hub, which always publishes.
        const out = legacy && (m as { t: string }).t === "roster"
          ? { ...m, roster: (m as unknown as { roster: Record<string, unknown>[] }).roster.map(({ write, ...rest }) => rest) }
          : m;
        ls.forEach((l) => l({ data: JSON.stringify(out) }));
      }),
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

type Session = { writesBlocked: boolean; readOnlyBlocked: boolean; retryWrites: () => void };

async function mount(hub: CollabHub, docId: string, clientId: string, legacy = false) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let session: Session | null = null;
  await act(async () => {
    root.render(createElement(CollabEditor, {
      url: "ws://x", docId, clientId, createSocket: factoryFor(hub, legacy), editable: true,
      onSession: (s: never) => { session = s; },
    }));
  });
  for (let i = 0; i < 40 && !container.querySelector(".dxw-page"); i++) await tick();
  const liveDoc = (): DocxDocument => {
    const key = Object.keys(container).find((k) => k.startsWith("__reactContainer$"))!;
    const stack: unknown[] = [(container as unknown as Record<string, unknown>)[key]];
    let guard = 0;
    while (stack.length && guard++ < 5000) {
      const f = stack.pop() as { memoizedProps?: { collab?: { doc?: DocxDocument } }; child?: unknown; sibling?: unknown } | null;
      if (!f) continue;
      const d = f.memoizedProps?.collab?.doc;
      if (d) return d;
      if (f.child) stack.push(f.child);
      if (f.sibling) stack.push(f.sibling);
    }
    throw new Error("no live doc");
  };
  const type = async (text: string) => {
    const page = container.querySelector<HTMLElement>(".dxw-page");
    if (!page) return;
    const span = page.querySelector("span") ?? page;
    await act(async () => {
      const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      span.dispatchEvent(new MouseEvent("mousedown", o));
      span.dispatchEvent(new MouseEvent("mouseup", o));
    });
    await tick();
    const target = (container.contains(document.activeElement)
      ? (document.activeElement as HTMLElement)
      : container.querySelector("textarea")) ?? container;
    for (const ch of text) {
      await act(async () => {
        target.dispatchEvent(new KeyboardEvent("keydown", { key: ch, bubbles: true, cancelable: true }));
      });
    }
    await settle(10);
  };
  return {
    container, liveDoc, type,
    /** The editor's writable surface — absent in viewer mode. */
    editable: () => !!container.querySelector("textarea"),
    session: () => session!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

/**
 * Put the SERVER in a state where it refuses this client's writes — AND re-fan
 * the roster, which is what the real admin path does. Mutating the flag alone
 * would leave every client holding a stale "allowed" status, a state the
 * server never actually produces.
 */
function refan(hub: CollabHub, docId: string) {
  const h = hub as unknown as { rooms: Map<string, unknown>; broadcastRoster: (r: unknown) => void };
  h.broadcastRoster(h.rooms.get(docId));
}
function lockRoom(hub: CollabHub, docId: string) {
  const room = (hub as unknown as { rooms: Map<string, { readOnly?: boolean }> }).rooms.get(docId)!;
  room.readOnly = true;
  refan(hub, docId);
}
function demote(hub: CollabHub, docId: string, clientId: string) {
  const room = (hub as unknown as { rooms: Map<string, { demoted?: Set<string> }> }).rooms.get(docId)!;
  room.demoted = new Set([clientId]);
  refan(hub, docId);
}
function unlockRoom(hub: CollabHub, docId: string) {
  const room = (hub as unknown as { rooms: Map<string, { readOnly?: boolean }> }).rooms.get(docId)!;
  room.readOnly = false;
  refan(hub, docId);
}

describe("the server's write status closes the first-edit gap", () => {
  it("joining an ALREADY-LOCKED room gates before any edit is made — nothing dies to discover it", async () => {
    // THE POINT OF THE SIGNAL. Before it, a client could only learn it was
    // blocked by making an edit and having it refused: the keystroke applied
    // locally, appeared on screen, and was healed away. Gating fixed edits
    // 2..N; only a status present AT JOIN fixes edit 1.
    const hub = new CollabHub(provider);
    // Lock the room BEFORE anyone joins, so the very first roster this client
    // receives already says so.
    const seeded = await mount(hub, "prelock", "seed");
    await settle();
    lockRoom(hub, "prelock");
    await seeded.unmount();

    const ed = await mount(hub, "prelock", "late");
    await settle();
    expect(ed.session().writesBlocked, "the join itself must reveal the block").toBe(true);
    expect(ed.editable(), "so the editor never offers a writable surface").toBe(false);
    // THE ASSERTION THAT MAKES IT THE FIRST-EDIT FIX: no edit was ever refused.
    // If this flag were set, the gate would have come from a dead keystroke —
    // which is the old behaviour wearing the new behaviour's result.
    expect(ed.session().readOnlyBlocked, "no edit should have had to die for this").toBe(false);

    const frozen = Buffer.from(ed.liveDoc().save());
    await ed.type("FIRST");
    await settle(20);
    expect(Buffer.from(ed.liveDoc().save()), "the first keystroke must not mutate anything").toEqual(frozen);
    expect(ed.session().readOnlyBlocked, "and still nothing had to be refused").toBe(false);
    await ed.unmount();
  });

  it("reports WHY, so the copy can stop guessing", async () => {
    // owner-lock may lift; viewer-role never will. A UI that cannot tell them
    // apart has to tell one of the two groups something false.
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "why", "someone");
    await settle();
    expect(ed.session().writeStatus).toBe("allowed");
    demote(hub, "why", "someone");
    await ed.type("A");
    await settle(20);
    expect(ed.session().writesBlocked).toBe(true);
    expect(ed.session().writeStatus).toBe("demoted");
    await ed.unmount();
  });
});

describe("writes the server will refuse ⇒ viewer mode, never a banner over a live editor", () => {
  it("owner read-only lock: the status gates it, and typing leaves the document BYTE-UNCHANGED", async () => {
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "ro", "editor");
    await settle();
    expect(ed.editable(), "precondition: the editor starts writable").toBe(true);
    lockRoom(hub, "ro");
    await settle(10);

    // No edit was needed: the lock re-fanned and the editor closed itself.
    expect(ed.session().writesBlocked).toBe(true);
    expect(ed.session().writeStatus).toBe("owner-lock");
    expect(ed.editable(), "the editor must drop to viewer mode").toBe(false);

    // BYTE-UNCHANGED is the claim: "it didn't sync" was true even while the
    // bug existed, so only the local document's bytes tell the two apart.
    const frozen = Buffer.from(ed.liveDoc().save());
    await ed.type("ABCDEFGH");
    await settle(20);
    expect(Buffer.from(ed.liveDoc().save()), "no keystroke may mutate the local document").toEqual(frozen);
    await ed.unmount();
  });

  it("a LIFT arrives on its own — no click, no reload", async () => {
    // The exit criterion for the signal: before it, a lifted lock left the
    // client stranded in viewer mode because nothing on the wire announced it.
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "lift", "editor");
    await settle();
    lockRoom(hub, "lift");
    await settle(10);
    expect(ed.editable()).toBe(false);

    unlockRoom(hub, "lift");
    await settle(10);
    expect(ed.session().writeStatus).toBe("allowed");
    expect(ed.editable(), "the lift must restore the editor by itself").toBe(true);
    const before = Buffer.from(ed.liveDoc().save());
    await ed.type("AGAIN");
    await settle(20);
    expect(Buffer.from(ed.liveDoc().save()), "and typing must work again").not.toEqual(before);
    await ed.unmount();
  });

  it("per-client demotion gates identically — one predicate, not a read-only special case", async () => {
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "dem", "viewer-1");
    await settle();
    demote(hub, "dem", "viewer-1");
    await settle(10);
    expect(ed.session().writesBlocked).toBe(true);
    expect(ed.session().writeStatus).toBe("demoted");
    expect(ed.editable()).toBe(false);
    const frozen = Buffer.from(ed.liveDoc().save());
    await ed.type("XYZ");
    await settle(20);
    expect(Buffer.from(ed.liveDoc().save())).toEqual(frozen);
    await ed.unmount();
  });

  it("OLDER SERVER (no status published): falls back to refusal-driven gating, never to 'allowed'", async () => {
    // The contract's other branch. Absent must not read as permission — a
    // permissive default would put the user straight back in the
    // first-edit-lost loop against a server that still refuses the writes.
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "legacy", "editor", true);
    await settle();
    expect(ed.session().writeStatus, "this fixture must actually publish nothing").toBeUndefined();
    expect(ed.editable(), "an absent status must not block a client that may write").toBe(true);

    lockRoom(hub, "legacy");
    await settle(10);
    // Nothing announced it, so the FIRST edit is still the discovery — the
    // known cost of the fallback path, and why the signal exists.
    expect(ed.editable(), "still writable: nothing has told this client yet").toBe(true);
    await ed.type("A");
    await settle(20);
    expect(ed.session().writesBlocked, "the refusal must raise the fallback gate").toBe(true);
    expect(ed.editable()).toBe(false);

    const frozen = Buffer.from(ed.liveDoc().save());
    await ed.type("BCDEFGH");
    await settle(20);
    expect(Buffer.from(ed.liveDoc().save()), "and everything after it is refused input").toEqual(frozen);

    // retryWrites is the ONLY escape here: the block is sticky and no lift is
    // announced. This is why it survives as belt-and-braces.
    await act(async () => { ed.session().retryWrites(); });
    await settle(5);
    expect(ed.editable(), "the manual override must restore the surface").toBe(true);
    await ed.unmount();
  });

  it("RATE-LIMIT must not gate — it means slow down, and the connection re-drives it", async () => {
    // The mistake a future "gate on every refusal" refactor would make.
    // Throttling is not refusal: the pending queue is re-sent after backoff,
    // and dropping the user into viewer mode would break bursty typing.
    const hub = new CollabHub(provider);
    const ed = await mount(hub, "rl", "editor");
    await settle();
    const before = Buffer.from(ed.liveDoc().save());
    await ed.type("HELLO");
    await settle(20);
    expect(ed.session().writesBlocked, "normal typing must never raise the gate").toBe(false);
    expect(ed.editable()).toBe(true);
    expect(Buffer.from(ed.liveDoc().save()), "and it must actually have typed").not.toEqual(before);
    await ed.unmount();
  });
});
