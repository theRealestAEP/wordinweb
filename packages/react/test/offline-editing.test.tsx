// @vitest-environment jsdom
/**
 * OFFLINE EDITING END TO END (plan doc 12 §5 "offline mode is first-class" +
 * doc 15 §2 the offline tail / arrival ladder).
 *
 * The contract under test, stated as the refusal/unreachable split:
 *
 *  - The server SAID NO (owner lock, demotion, viewer link, refusal) —
 *    writes stay blocked, offline or not. write-gate.test.tsx pins the
 *    online half; the offline half is pinned here.
 *  - The server is UNREACHABLE — the editor stays writable; every edit
 *    applies locally, lands in the durable offline tail, survives a reload,
 *    and replays on reconnect: silently after an automatic same-page retry,
 *    after a merge check for a stable-URL return, as suggestions (diverged,
 *    small tail), or as a draft (diverged, large tail).
 *  - The tail is BOUNDED: at the cap recording stops by loudly gating the
 *    editor, never by silently discarding keystrokes.
 *
 * The load-bearing assertions are byte-identical documents across replicas
 * after replay, and the store's contents across simulated reloads — banner
 * state alone proves nothing (see the write-gate history).
 */
import { describe, expect, it } from "vitest";
import { createElement, act, forwardRef, useImperativeHandle, type Ref } from "react";
import { createRoot } from "react-dom/client";
import { CollabHub, blankDocxBytes, type Connection, type ServerMessage } from "@wordinweb/server";
import { InMemoryBundleStore } from "@wordinweb/collab/client";
import { serializeXml } from "@wordinweb/core";
import { useCollab, type CollabSession } from "../src/collab.js";

/** Same staged-death factory as reconnect.test.tsx, plus a swappable delay.
 * Connection ids are NAMESPACED per factory (write-gate style): two probes
 * share one hub here, and colliding `Connection.id`s would cross their
 * per-connection state server-side. */
let factorySeq = 0;
function factoryFor(hub: CollabHub, delayMs = 1) {
  const ns = `o${factorySeq++}-`;
  const sockets: { kill: () => void; blackhole: () => void }[] = [];
  const outage = { value: false };
  let n = 0;
  const factory = (_url: string): WebSocket => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    let dead = false;
    const gone = () => dead || outage.value;
    let opened = false;
    const conn: Connection = {
      id: `${ns}c${n++}`,
      send: (m: ServerMessage) => {
        if (gone()) return;
        const data = JSON.stringify(m);
        setTimeout(() => {
          if (!gone()) ls.forEach((l) => l({ data }));
        }, delayMs);
      },
    };
    const sock = {
      onclose: null as null | (() => void),
      onerror: null as null | (() => void),
      send: (d: string) => {
        if (gone()) return;
        setTimeout(() => {
          if (!gone()) void hub.handle(conn, JSON.parse(d) as never);
        }, delayMs);
      },
      close: () => {
        dead = true;
      },
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") ls.push(cb as never);
        else if (!opened) {
          opened = true;
          (cb as () => void)();
        }
      },
    };
    sockets.push({
      kill: () => {
        dead = true;
        sock.onclose?.();
      },
      blackhole: () => {
        dead = true;
      },
    });
    return sock as unknown as WebSocket;
  };
  return { factory, sockets, outage };
}

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}
async function until(pred: () => boolean, label: string, n = 1200) {
  for (let i = 0; i < n && !pred(); i++) await tick();
  if (!pred()) throw new Error(`timeout: ${label}`);
}

const FAST = { intervalMs: 20, timeoutMs: 30, maxRetries: 3 };

const Probe = forwardRef(function Probe(
  props: {
    factory: (u: string) => WebSocket;
    store: InMemoryBundleStore;
    clientId?: string;
    offlineTailCap?: number;
    startOffline?: boolean;
  },
  ref: Ref<{ session: CollabSession }>,
) {
  const session = useCollab({
    url: "ws://x",
    docId: "d",
    clientId: props.clientId ?? "alice",
    createSocket: props.factory,
    store: props.store,
    liveness: FAST,
    offlineTailCap: props.offlineTailCap,
    startOffline: props.startOffline,
  });
  useImperativeHandle(ref, () => ({ session }), [session]);
  return null;
});

async function mount(
  hub: CollabHub,
  store: InMemoryBundleStore,
  opts: { clientId?: string; offlineTailCap?: number; delayMs?: number; startOffline?: boolean } = {},
) {
  const f = factoryFor(hub, opts.delayMs ?? 1);
  const holder: { current: { session: CollabSession } | null } = { current: null };
  const root = createRoot(document.createElement("div"));
  await act(async () => {
    root.render(
      createElement(Probe, {
        factory: f.factory,
        store,
        clientId: opts.clientId,
        offlineTailCap: opts.offlineTailCap,
        startOffline: opts.startOffline,
        ref: (v: never) => (holder.current = v),
      }),
    );
  });
  await until(() => holder.current?.session.ready === true, "initial welcome");
  return {
    ...f,
    get session() {
      return holder.current!.session;
    },
    unmount: async () => {
      await act(async () => root.unmount());
      await tick(20);
    },
  };
}

const seeded = () => {
  const hub = new CollabHub(null);
  hub.seed("d", blankDocxBytes());
  return hub;
};
const insertAt = (offset: number, text: string) =>
  ({ kind: "insertText", at: { blockId: 1, runId: 2, offset }, text }) as never;
const docText = (s: CollabSession): string => {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(s.doc!.docRoot as never);
};
const bytes = (s: CollabSession) => Buffer.from(s.doc!.save());
/** Synchronous peek at the stored tail (the store's API is async; polling
 * predicates need a sync read — the in-memory map is right there). */
const storedTail = (store: InMemoryBundleStore, id = "d"): unknown[] | undefined =>
  (store as unknown as { bundles: Map<string, { offlineTail?: unknown[] }> }).bundles.get(id)?.offlineTail;

describe("offline capture is durable (the live data-loss bug)", () => {
  it("opens a stored room copy offline without creating a socket, then rejoins through the normal merge check", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const live = await mount(hub, store);
    await act(async () => live.session.submitOp(insertAt(0, "BASE")));
    await until(() => live.session.version > 1 && docText(live.session) === "BASE", "base edit echoed");
    await until(
      () => {
        const bundle = (store as unknown as { bundles: Map<string, { pending: unknown[]; confirmedSeq: number }> }).bundles.get("d");
        return bundle?.pending.length === 0 && bundle.confirmedSeq === 1;
      },
      "confirmed bundle persisted",
    );
    await live.unmount();

    const offline = await mount(hub, store, { startOffline: true });
    expect(offline.sockets, "opening the browser copy must not join the room").toHaveLength(0);
    expect(offline.session.connection).toBe("lost");
    expect(offline.session.offline).toEqual({ editsHeld: 0, capped: false });
    expect(docText(offline.session)).toBe("BASE");

    await act(async () => offline.session.submitOp(insertAt(4, "+OFF")));
    expect(docText(offline.session)).toBe("BASE+OFF");
    await until(() => storedTail(store)?.length === 1, "offline edit persisted");

    await act(async () => offline.session.reconnect());
    await until(() => offline.sockets.length === 1, "rejoin creates the first socket");
    await until(() => offline.session.arrival?.mode === "replay", "rejoin offers the merge check");
    await act(async () => offline.session.reconcile("replay"));
    await until(() => offline.session.connection === "live" && docText(offline.session) === "BASE+OFF", "offline edit merged");
    await offline.unmount();
  });

  it("offline edits persist through the throttled writer WITHOUT an unmount, tagged with their epoch", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    await act(async () => p.session.submitOp(insertAt(0, "LIVE")));
    await until(() => p.session.version > 1, "edit echoed");
    await tick(30); // let the persister write the confirmed state
    const genesis = (await store.get("d"))!.genesisId;

    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");
    await act(async () => p.session.submitOp(insertAt(4, "-OFF")));
    expect(p.session.offline!.editsHeld).toBe(1);

    // No unmount, no pagehide — the trailing throttle alone must land it,
    // because an OS kill grants no goodbye write.
    await until(
      () => storedTail(store)?.length === 1,
      "throttled offline write landed",
    );
    const stored = (await store.get("d"))!;
    expect(stored.offlineTail).toHaveLength(1);
    expect(stored.offlineTail![0]).toMatchObject({ kind: "insertText", text: "-OFF" });
    // The epoch stamp is what lets the NEXT session decide silent-replay vs
    // arrival ladder without guessing.
    expect(stored.offlineTailEpoch).toBe(genesis);
    expect(stored.offlineBaseSeq).toBe(stored.confirmedSeq);
    await p.unmount();
  });

  it("the PRE-APPLIED path (editor typing) records without re-applying — no doubled keystrokes", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    await act(async () => p.session.submitOp(insertAt(0, "AB")));
    await until(() => p.session.version > 1, "edit echoed");
    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");

    // The editor applies to the live doc itself and THEN emits — so capture
    // must record only. Re-applying here would double the keystroke (the
    // same bug the live submitPreApplied path exists to avoid).
    const frozen = bytes(p.session);
    await act(async () => p.session.submit(insertAt(2, "C")));
    expect(p.session.offline!.editsHeld, "recorded").toBe(1);
    expect(Buffer.compare(bytes(p.session), frozen), "not re-applied").toBe(0);
    await until(() => storedTail(store)?.length === 1, "persisted");
    await p.unmount();
  });

  it("edits captured WHILE the resume read is in flight survive it — the stored tail must not clobber memory", async () => {
    // The narrow race the merge guard exists for: a reconnect's store read
    // resolves AFTER the user typed another offline edit. The stored tail
    // trails by the throttle window, so overwriting the in-memory tail with
    // it would silently drop that edit.
    const hub = seeded();
    class GatedStore extends InMemoryBundleStore {
      gate: Promise<void> | null = null;
      async get(id: string) {
        if (this.gate) await this.gate;
        return super.get(id);
      }
    }
    const store = new GatedStore();
    const p = await mount(hub, store);
    await act(async () => p.session.submitOp(insertAt(0, "S")));
    await until(() => p.session.version > 1, "seed echoed");
    await tick(30); // persister writes the confirmed state

    let release!: () => void;
    store.gate = new Promise<void>((r) => (release = r));
    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");
    // Past the backoff: the reconnect effect has torn down (flushed) and is
    // now parked on the GATED store read. This edit lands in memory only.
    await tick(600);
    await act(async () => p.session.submitOp(insertAt(1, "X")));
    expect(p.session.offline!.editsHeld).toBe(1);

    release();
    store.gate = null;
    await tick(50); // the resume-load runs against a store with NO tail
    expect(p.session.offline!.editsHeld, "the in-flight read must not clobber the fresher memory").toBe(1);

    // And it still replays: the edit was kept, not merely counted.
    p.outage.value = false;
    await until(() => p.session.connection === "live" && p.session.offline === null, "recovered");
    await until(() => docText(p.session) === "SX", "the raced edit replayed");
    await p.unmount();
  });

  it("a reload (fresh mount, same store) still holds the tail and still offers/replays it", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    await act(async () => p.session.submitOp(insertAt(0, "BASE")));
    await until(() => p.session.version > 1, "edit echoed");
    await tick(50); // persist the drained live queue before the offline copy starts
    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");
    await act(async () => p.session.submitOp(insertAt(4, "+TAIL")));
    await p.unmount(); // the close-the-laptop path: teardown flush

    const stored = (await store.get("d"))!;
    expect(stored.offlineTail, "the tail survived the reload").toHaveLength(1);
    expect(stored.offlineBaseSeq, "the stable URL keeps its room base").toBeTypeOf("number");
    expect(stored.pending, "the offline copy has no in-flight live edits").toHaveLength(0);

    // Same epoch is still live on the hub. A fresh stable-URL mount restores
    // the recorded base sequence and asks before it merges the tail.
    const p2 = await mount(hub, store);
    await until(() => p2.session.arrival?.mode === "replay", "stable URL offers a merge check");
    expect(p2.session.arrival).toMatchObject({ mode: "replay", tailLength: 1 });
    await act(async () => p2.session.reconcile("replay"));
    await until(() => docText(p2.session).includes("+TAIL"), "tail replayed after reload");
    await until(() => !storedTail(store)?.length, "stored tail erased after replay");
    expect(docText(p2.session)).toBe("BASE+TAIL");
    await p2.unmount();
  });
});

describe("arrival outcome 1 — same epoch: silent replay, byte-identical convergence", () => {
  it("type offline → reconnect → both clients and the server converge byte-identically", async () => {
    const hub = seeded();
    const storeA = new InMemoryBundleStore();
    const storeB = new InMemoryBundleStore();
    const a = await mount(hub, storeA, { clientId: "alice" });
    const b = await mount(hub, storeB, { clientId: "bob" });

    await act(async () => a.session.submitOp(insertAt(0, "S")));
    await until(() => docText(b.session) === "S", "seed reached bob");

    // Alice's network dies (only hers); she keeps typing.
    a.outage.value = true;
    await until(() => a.session.offline !== null, "alice offline");
    for (const [i, ch] of [..."123"].entries()) {
      await act(async () => a.session.submitOp(insertAt(1 + i, ch)));
    }
    expect(a.session.offline!.editsHeld).toBe(3);
    expect(docText(a.session), "offline edits are visible to alice").toBe("S123");
    expect(docText(b.session), "and invisible to bob until replay").toBe("S");

    // Network returns: the tail replays silently (same epoch — to the room
    // alice is just a client that was disconnected with a big pending queue).
    a.outage.value = false;
    await until(() => a.session.connection === "live" && a.session.offline === null, "alice recovered");
    await until(() => docText(b.session) === "S123", "bob received the replayed tail");
    await until(() => Buffer.compare(bytes(a.session), bytes(b.session)) === 0, "byte-identical");
    expect(docText(a.session)).toBe("S123");

    // THE DRAIN DISCIPLINE IS OBSERVABLE IN THE LOG (doc 03 one-in-flight /
    // doc 15 §2 "replayed with drains"): each replayed intent must be BASED
    // past its predecessor's sequenced entry — the client waited for the
    // echo before submitting the next. A fire-all-at-once replay shares one
    // base across the tail, which is exactly the multi-pending shape the
    // gated sibling math exists for and this replay must never produce.
    const entries = (hub as unknown as {
      rooms: Map<string, { session: { entriesSince: (s: number) => { seq: number; kind: string; intent?: { text?: string; base: number } }[] } }>;
    }).rooms.get("d")!.session.entriesSince(0);
    const replayed = entries.filter((e) => e.kind === "applied" && ["1", "2", "3"].includes(e.intent?.text ?? ""));
    expect(replayed).toHaveLength(3);
    for (let i = 1; i < replayed.length; i++) {
      expect(
        replayed[i].intent!.base,
        "each replayed intent is based past the previous one's seq (one-in-flight)",
      ).toBeGreaterThanOrEqual(replayed[i - 1].seq);
    }
    await a.unmount();
    await b.unmount();
  });

  it("EXACTLY-ONCE: killed mid-replay, the resume finishes the tail without duplicating anything", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    // A slow transport so the drained replay takes long enough to interrupt.
    const p = await mount(hub, store, { delayMs: 12 });
    await act(async () => p.session.submitOp(insertAt(0, "S")));
    await until(() => p.session.version > 1, "seed echoed");

    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");
    for (const [i, ch] of [..."ABCDEF"].entries()) {
      await act(async () => p.session.submitOp(insertAt(1 + i, ch)));
    }
    expect(p.session.offline!.editsHeld).toBe(6);

    // Reconnect and KILL the tab midway through the drained replay.
    p.outage.value = false;
    await until(() => p.session.connection === "live", "reconnected");
    await until(() => {
      const left = storedTail(store)?.length ?? 6;
      return left > 0 && left < 6;
    }, "replay is mid-flight (some drained, some not)");
    await p.unmount(); // the kill: teardown flush persists pending + remainder

    // The resumed session must finish the job: pending replays verbatim
    // (deduped by clientId/clientSeq — the server-side guarantee this test
    // exists to prove), the tail remainder replays via the same silent path,
    // and NOTHING lands twice.
    const p2 = await mount(hub, store, { delayMs: 1 });
    await until(() => docText(p2.session) === "SABCDEF", "every offline edit landed exactly once");
    // A second joiner from nothing sees the same bytes — the sequencer holds
    // exactly one copy of each edit.
    const fresh = await mount(hub, new InMemoryBundleStore(), { clientId: "carol" });
    await until(() => Buffer.compare(bytes(p2.session), bytes(fresh.session)) === 0, "byte-identical with a fresh joiner");
    expect(docText(fresh.session)).toBe("SABCDEF");
    await p2.unmount();
    await fresh.unmount();
  });
});

describe("arrival outcomes 2+3 — diverged epoch: suggestions and draft", () => {
  /** Build a bundle with a real offline tail on hub1, typed while offline. */
  async function offlineBundle(store: InMemoryBundleStore, offlineEdits: (s: CollabSession) => Promise<void>) {
    const hub1 = seeded();
    const p = await mount(hub1, store);
    await act(async () => p.session.submitOp(insertAt(0, "mine")));
    await until(() => p.session.version > 1, "edit echoed");
    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");
    await offlineEdits(p.session);
    await p.unmount();
    return (await store.get("d"))!;
  }
  /** A different-epoch hub for the same docId (bob re-seeded — doc 12 §5.2). */
  const reseeded = () => {
    const hub2 = new CollabHub(null, undefined, undefined, () => 0, () => "g_two");
    hub2.seed("d", blankDocxBytes());
    return hub2;
  };

  it("suggest: Enter and its following text stay linked; unsupported structure remains in the draft", async () => {
    const store = new InMemoryBundleStore();
    const old = await offlineBundle(store, async (s) => {
      const [nb, nr] = s.allocIds(2);
      await act(async () =>
        s.submitOp({ kind: "splitParagraph", at: { blockId: 1, runId: 2, offset: 0 }, newBlockId: nb, newRunId: nr } as never),
      );
      await act(async () => s.submitOp({ kind: "insertText", at: { blockId: nb, runId: nr, offset: 0 }, text: "NEW LINE" } as never));
      // Formatting has no tracked suggestion form. Count it and retain the
      // complete original tail in the banked draft.
      await act(async () => s.submitOp({ kind: "formatRun", blockId: 1, runId: 2, patch: { bold: true } } as never));
    });
    expect(old.offlineTail).toHaveLength(3);

    const hub2 = reseeded();
    const p2 = await mount(hub2, store);
    await until(() => p2.session.arrival !== null, "arrival offered");
    expect(p2.session.arrival).toMatchObject({ mode: "suggest", tailLength: 3, structural: 1 });
    expect(p2.session.epochChanged, "the fork was surfaced, not merged").not.toBeNull();
    const paragraphsBefore = (serializeXml(p2.session.doc!.docRoot).match(/<w:p[ >]/g) ?? []).length;
    const hubSession = (hub2 as unknown as { rooms: Map<string, { session: { seq: number } }> }).rooms.get("d")!.session;
    const seqBefore = hubSession.seq;

    await act(async () => p2.session.reconcile("suggest"));
    await until(() => docText(p2.session).includes("NEW LINE"), "text after the suggested Enter landed");
    // The replay must run the WHOLE tail (structural entries popped, not
    // stalled) and erase it — otherwise the assertions below could pass on
    // a half-finished replay.
    await until(() => !storedTail(store)?.length, "tail fully drained and erased");
    // Landed as a TRACKED CHANGE attributed to the returning editor — the
    // crowd reviews it; nothing was bulldozed.
    const xml = serializeXml(p2.session.doc!.docRoot);
    expect(xml).toContain("w:ins");
    expect(xml).toContain('w:author="alice"');
    expect(xml).toMatch(/<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:ins /);

    // A second client sees byte-identical content (the replay used the
    // ordinary sequencer path, not a local shortcut).
    const other = await mount(hub2, new InMemoryBundleStore(), { clientId: "bob" });
    await until(() => Buffer.compare(bytes(p2.session), bytes(other.session)) === 0, "converged");

    // The structural edit did NOT vanish: the banked draft holds the whole
    // tail, unsupported formatting included.
    const draft = await store.get(`d#draft-${old.genesisId}`);
    expect(draft, "draft slot banked on epoch change").toBeTruthy();
    expect(draft!.offlineTail?.some((i) => (i as { kind: string }).kind === "formatRun")).toBe(true);
    // The live split carries a tracked paragraph mark. It remains reversible.
    const paragraphsAfter = (serializeXml(p2.session.doc!.docRoot).match(/<w:p[ >]/g) ?? []).length;
    expect(paragraphsAfter, "the suggested Enter creates one reviewable paragraph").toBe(paragraphsBefore + 1);
    // Split and following text sequence. Unsupported formatting stays local.
    expect(hubSession.seq).toBe(seqBefore + 2);
    await p2.unmount();
    await other.unmount();
  });

  it("draft: a tail past the suggest threshold is offered draft-only, and the draft keeps everything", async () => {
    const store = new InMemoryBundleStore();
    const old = await offlineBundle(store, async (s) => {
      for (let i = 0; i < 51; i++) {
        await act(async () => s.submitOp(insertAt(4 + i, String.fromCharCode(97 + (i % 26)))));
      }
    });
    expect(old.offlineTail).toHaveLength(51);

    const hub2 = reseeded();
    const p2 = await mount(hub2, store);
    await until(() => p2.session.arrival !== null, "arrival offered");
    expect(p2.session.arrival!.mode, "51 > threshold(50) ⇒ draft only").toBe("draft");
    expect(p2.session.arrival!.tailLength).toBe(51);

    await act(async () => p2.session.reconcile("draft"));
    expect(p2.session.arrival).toBeNull();
    // The draft slot holds the superseded copy AND the full tail; the live
    // slot's tail is erased so it can never be re-offered or double-kept.
    const draft = await store.get(`d#draft-${old.genesisId}`);
    expect(draft!.offlineTail).toHaveLength(51);
    await until(() => !storedTail(store)?.length, "live tail erased");
    // Nothing leaked into the live document.
    expect(docText(p2.session)).not.toContain("abc");
    await p2.unmount();
  });
});

describe("the tail cap (doc 15 §4.3)", () => {
  it("recording stops AT the cap by gating the editor loudly — capped state, writesBlocked, durable tail intact", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store, { offlineTailCap: 3 });
    await act(async () => p.session.submitOp(insertAt(0, "S")));
    await until(() => p.session.version > 1, "seed echoed");
    p.outage.value = true;
    await until(() => p.session.offline !== null, "offline");
    // Wait out the teardown/rebuild so the stale ready connection is gone
    // (submits that race the drop window land in `pending`, which is safe
    // but a different path than the one under test).
    await tick(500);

    for (const [i, ch] of [..."abc"].entries()) {
      await act(async () => p.session.submitOp(insertAt(1 + i, ch)));
    }
    expect(p.session.offline).toEqual({ editsHeld: 3, capped: true });
    expect(p.session.writesBlocked, "the cap folds into the single write gate").toBe(true);

    // The next edit is NOT captured and NOT silently applied: it falls to
    // the not-ready connection's loud droppedPreReady backstop, and the
    // editor surface is already withdrawn by writesBlocked anyway.
    const frozen = bytes(p.session);
    const dropsBefore = p.session.droppedPreReady;
    await act(async () => p.session.submitOp(insertAt(4, "X")));
    expect(p.session.offline!.editsHeld, "nothing further recorded").toBe(3);
    expect(Buffer.compare(bytes(p.session), frozen), "nothing further applied").toBe(0);
    expect(p.session.droppedPreReady, "the overflow was counted, not swallowed").toBe(dropsBefore + 1);

    // The three captured edits stay durable.
    await until(() => storedTail(store)?.length === 3, "capped tail persisted");
    await p.unmount();
  });
});

describe("refusals still block — offline editing punches no hole", () => {
  function refan(hub: CollabHub) {
    const h = hub as unknown as { rooms: Map<string, unknown>; broadcastRoster: (r: unknown) => void };
    h.broadcastRoster(h.rooms.get("d"));
  }
  it("a DEMOTED participant stays blocked when their socket drops — no offline capture opens up", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    (hub as unknown as { rooms: Map<string, { demoted?: Set<string> }> }).rooms.get("d")!.demoted = new Set(["alice"]);
    refan(hub);
    await until(() => p.session.writesBlocked, "demotion gated");
    expect(p.session.writeStatus).toBe("demoted");

    p.outage.value = true;
    await until(() => p.session.connection !== "live", "offline");
    // The refusal outranks the transport fault: still gated, and the
    // offline state is NOT offered (a UI reading it would falsely promise
    // "your changes are being kept").
    expect(p.session.writesBlocked, "still blocked while offline").toBe(true);
    expect(p.session.offline, "no offline-capture promise for a refused writer").toBeNull();
    await p.unmount();
  });

  it("an OWNER-LOCKED room stays blocked when the socket drops", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    (hub as unknown as { rooms: Map<string, { readOnly?: boolean }> }).rooms.get("d")!.readOnly = true;
    refan(hub);
    await until(() => p.session.writesBlocked, "lock gated");
    expect(p.session.writeStatus).toBe("owner-lock");

    p.outage.value = true;
    await until(() => p.session.connection !== "live", "offline");
    expect(p.session.writesBlocked).toBe(true);
    expect(p.session.offline).toBeNull();
    await p.unmount();
  });

  it("a REFUSED session stays blocked and never reports offline", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    // The kick refusal path: the server ends the session for this client.
    await act(async () => {
      (hub as unknown as { rooms: Map<string, { conns: Set<{ send: (m: unknown) => void }> }> }).rooms
        .get("d")!
        .conns.forEach((c) => c.send({ t: "refused", reason: "kicked" }));
    });
    await until(() => p.session.refused === "kicked", "refusal surfaced");
    expect(p.session.writesBlocked, "a refusal blocks, full stop").toBe(true);
    p.outage.value = true;
    await tick(100);
    expect(p.session.offline, "a dead session never turns into offline editing").toBeNull();
    await p.unmount();
  });
});

describe("seeded fuzz: random offline tails against a concurrently-mutating room", () => {
  const mulberry32 = (seed: number) => () => {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };

  for (const seed of [7, 41]) {
    it(`seed ${seed}: offline edits + concurrent live edits converge byte-identically`, async () => {
      const rnd = mulberry32(seed);
      const hub = seeded();
      const a = await mount(hub, new InMemoryBundleStore(), { clientId: "alice" });
      const b = await mount(hub, new InMemoryBundleStore(), { clientId: "bob" });
      await act(async () => a.session.submitOp(insertAt(0, "SEED")));
      await until(() => docText(b.session) === "SEED", "seeded");

      a.outage.value = true;
      await until(() => a.session.offline !== null, "alice offline");

      const offlineOps = 8 + Math.floor(rnd() * 8);
      for (let i = 0; i < offlineOps; i++) {
        const text = docText(a.session);
        if (rnd() < 0.75 || text.length < 2) {
          const at = Math.floor(rnd() * (text.length + 1));
          await act(async () => a.session.submitOp(insertAt(at, String.fromCharCode(65 + Math.floor(rnd() * 26)))));
        } else {
          const start = Math.floor(rnd() * (text.length - 1));
          await act(async () =>
            a.session.submitOp({ kind: "deleteText", blockId: 1, runId: 2, start, end: start + 1 } as never),
          );
        }
        // The crowd keeps typing while alice is away.
        if (i % 2 === 0) {
          const bt = docText(b.session);
          await act(async () =>
            b.session.submitOp(insertAt(Math.floor(rnd() * (bt.length + 1)), String.fromCharCode(97 + i % 26))),
          );
        }
      }
      expect(a.session.offline!.editsHeld, "every offline op was recorded").toBeGreaterThan(0);

      a.outage.value = false;
      await until(() => a.session.connection === "live" && a.session.offline === null, "alice recovered");
      // Converged when both replicas hold the same bytes (byte-identical is
      // the contract; text equality alone would miss structural divergence).
      await until(
        () => Buffer.compare(bytes(a.session), bytes(b.session)) === 0 && docText(a.session).length > 0,
        "byte-identical convergence after replay",
        800,
      );
      // And the server agrees with both (three-way, not just pairwise).
      const c = await mount(hub, new InMemoryBundleStore(), { clientId: "carol" });
      await until(() => Buffer.compare(bytes(c.session), bytes(b.session)) === 0, "fresh joiner byte-identical");
      await a.unmount();
      await b.unmount();
      await c.unmount();
    });
  }
});


/**
 * A WEDGED STORE MUST NOT WEDGE THE CONNECTION.
 *
 * IndexedDB does not reject when the quota is exhausted or storage is blocked
 * — it simply never settles. The join used to `await store.get(docId)` with no
 * bound, and the hello is sent only after that resolves, so the socket opened,
 * no hello was ever sent, the server saw a room with no participants and
 * evicted it on the empty-room grace, and the client sat on "Connecting…"
 * forever. Observed in production, with nothing in any log naming the cause:
 * `conn-open`, then `room-evict reason=empty`, and no `conn-close`.
 *
 * Losing the resume costs a pending queue and one cold rejoin. Never
 * connecting costs everything.
 */
describe("a storage read that never settles", () => {
  class DeadStore extends InMemoryBundleStore {
    override get(): Promise<never> {
      return new Promise<never>(() => {}); // never settles, exactly like a wedged IDB
    }
  }

  it("still joins — cold — instead of hanging on Connecting", async () => {
    const hub = seeded();
    const store = new DeadStore();
    // mount() waits for session.ready, so reaching this line at all is the
    // assertion: before the deadline it timed out here.
    // Generous poll budget: the join deliberately waits out the storage
    // deadline before proceeding, so this must outlast it.
    const a = await mount(hub, store);
    try {
      expect(a.session.ready, "the session must connect despite dead storage").toBe(true);
      // And it says so, rather than quietly looking like an empty first visit.
      // `storeSlow`, NOT `persistErrors`: nothing failed to WRITE, and the demo
      // renders persistErrors as "storage is full or blocked" — a different and
      // wrong claim for a read that timed out. Conflating them told one user
      // their storage was full when it held a few megabytes.
      expect(a.session.storeSlow, "the unresponsive store is reported").toBe(true);
      expect(a.session.persistErrors, "nothing failed to write").toBe(0);
    } finally {
      await a.unmount();
    }
  }, 30_000);
});
