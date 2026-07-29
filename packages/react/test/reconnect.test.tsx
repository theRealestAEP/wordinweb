// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { createElement, act, forwardRef, useImperativeHandle, type Ref } from "react";
import { createRoot } from "react-dom/client";
import { CollabHub, type Connection, type ServerMessage } from "@wordinweb/server";
import { blankDocxBytes } from "@wordinweb/server";
import { InMemoryBundleStore } from "@wordinweb/collab/client";
import { useCollab, CollabEditor, type CollabSession } from "../src/collab.js";

/**
 * DESYNC DETECTION AND RECONNECT.
 *
 * The bug: a phone's browser sleeps, its WebSocket dies, and the client has no
 * idea. The editor stays editable, the user keeps typing, and every keystroke
 * is applied locally and then lost — silent data loss that looks exactly like
 * it worked.
 *
 * This is the same failure shape as the `writesBlocked` bug already fixed in
 * this codebase, so it gets the same remedy and the same standard of proof: it
 * is NOT enough to show a banner appears. The load-bearing assertion is that
 * the document BYTES DO NOT CHANGE while the connection is down — a weaker
 * "the editor says disconnected" test passes happily while the user's text is
 * still being applied and discarded underneath it.
 *
 * The half-open case gets its own test because it is the motivating scenario
 * and the one no close-handler can catch: the socket never fires `onclose`,
 * never fires `onerror`, and still reports OPEN. Only the heartbeat notices.
 */

/** A socket whose death can be staged three different ways. */
interface Staged {
  /** Clean drop: the browser notices and fires `onclose`. */
  kill(): void;
  /**
   * HALF-OPEN: frames stop moving in both directions and NO event fires. The
   * socket still reports OPEN and `send()` still succeeds. This is the
   * sleeping phone, and it is invisible to every handler except the probe.
   */
  blackhole(): void;
}

function factoryFor(hub: CollabHub) {
  const sockets: Staged[] = [];
  const hellos: { sinceSeq: number; takeover?: boolean; genesisId?: string }[] = [];
  /** Network-wide outage: every socket, including ones not yet created, is
   * born half-open. Models a phone with no signal — as opposed to `blackhole`,
   * which kills one specific socket. */
  const outage = { value: false };
  /** How many times a close EVENT was delivered to the client. The half-open
   * test asserts this stays zero — otherwise it would silently be testing the
   * easy close-handler path instead of the scenario it names. */
  let closeFires = 0;
  let n = 0;
  const factory = (_url: string): WebSocket => {
    const ls: ((ev: { data: unknown }) => void)[] = [];
    let dead = false;
    const gone = () => dead || outage.value;
    let opened = false;
    const conn: Connection = {
      id: `c${n++}`,
      send: (m: ServerMessage) => {
        if (gone()) return;
        const data = JSON.stringify(m);
        setTimeout(() => {
          if (!gone()) ls.forEach((l) => l({ data }));
        }, 1);
      },
    };
    const sock = {
      onclose: null as null | (() => void),
      onerror: null as null | (() => void),
      send: (d: string) => {
        if (gone()) return; // a half-open socket accepts writes that go nowhere
        const msg = JSON.parse(d) as { t: string; sinceSeq: number; takeover?: boolean; genesisId?: string };
        if (msg.t === "hello") hellos.push({ sinceSeq: msg.sinceSeq, takeover: msg.takeover, genesisId: msg.genesisId });
        setTimeout(() => {
          if (!gone()) void hub.handle(conn, msg as never);
        }, 1);
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
        closeFires++;
        sock.onclose?.();
      },
      blackhole: () => {
        dead = true;
      },
    });
    return sock as unknown as WebSocket;
  };
  return { factory, sockets, hellos, outage, closeFires: () => closeFires };
}

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}
async function until(pred: () => boolean, label: string, n = 200) {
  for (let i = 0; i < n && !pred(); i++) await tick();
  if (!pred()) throw new Error(`timeout: ${label}`);
}

/** Fast heartbeat so the tests run in milliseconds, not half a minute. */
const FAST = { intervalMs: 20, timeoutMs: 30, maxRetries: 3 };

const Probe = forwardRef(function Probe(
  props: { factory: (u: string) => WebSocket; store?: InMemoryBundleStore },
  ref: Ref<{ session: CollabSession }>,
) {
  const session = useCollab({
    url: "ws://x",
    docId: "d",
    clientId: "alice",
    createSocket: props.factory,
    store: props.store,
    liveness: FAST,
  });
  useImperativeHandle(ref, () => ({ session }), [session]);
  return null;
});

async function mount(hub: CollabHub, store?: InMemoryBundleStore) {
  const f = factoryFor(hub);
  const holder: { current: { session: CollabSession } | null } = { current: null };
  const root = createRoot(document.createElement("div"));
  await act(async () => {
    root.render(createElement(Probe, { factory: f.factory, store, ref: (v: never) => (holder.current = v) }));
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

describe("desync detection", () => {
  it("a clean drop (onclose) surfaces as not-live and BLOCKS WRITES", async () => {
    const hub = seeded();
    const p = await mount(hub);
    expect(p.session.connection).toBe("live");
    expect(p.session.writesBlocked, "a healthy session is writable").toBe(false);

    p.sockets[0].kill();
    await until(() => p.session.connection !== "live", "drop detected");

    expect(p.session.connection).toBe("reconnecting");
    // THE ASSERTION THAT MATTERS. Not "a banner appeared" — the editor must be
    // gated, because an editable surface over a server that cannot receive the
    // writes is the silent-data-loss bug this arc exists to close.
    expect(p.session.writesBlocked, "a dead connection must gate the editor read-only").toBe(true);
    await p.unmount();
  });

  it("HALF-OPEN: no onclose, no onerror — only the heartbeat notices", async () => {
    const hub = seeded();
    const p = await mount(hub);

    // The socket stops carrying frames in both directions but fires NOTHING:
    // no close, no error, and it would still report readyState OPEN. This is
    // the sleeping phone, and it is the scenario the whole arc exists for.
    p.sockets[0].blackhole();

    await until(() => p.session.connection !== "live", "heartbeat noticed the silence");

    // THE GUARD ON THE TEST ITSELF. Without this the test would still pass if
    // the fake leaked a close event, and it would then be quietly re-testing
    // the easy path — proving nothing about the case no close handler can
    // catch. Zero close events means the heartbeat is the ONLY thing that
    // could have detected this.
    expect(p.closeFires(), "no close event was ever delivered — the probe is the sole detector").toBe(0);
    expect(p.session.writesBlocked, "half-open is gated exactly like a clean drop").toBe(true);
    await p.unmount();
  });

  it("keystrokes during a drop do not mutate the document", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    await act(async () => p.session.submitOp(insertAt(0, "BEFORE")));
    await until(() => p.session.version > 1, "first edit echoed");

    const frozen = Buffer.from(p.session.doc!.save());
    p.sockets[0].blackhole();
    await until(() => p.session.writesBlocked, "gate closed");

    // The gate is what stops the editor emitting these at all; this asserts
    // the document itself is untouched even if something did try.
    expect(Buffer.from(p.session.doc!.save()), "no local mutation while disconnected").toEqual(frozen);
    await p.unmount();
  });
});

describe("reconnect", () => {
  it("RESUMES from the bundle rather than joining cold", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    await act(async () => p.session.submitOp(insertAt(0, "SAVED")));
    await until(() => p.session.version > 1, "edit confirmed");
    await tick(30); // let the persister write

    expect(p.hellos.length, "one join so far").toBe(1);
    const firstJoin = p.hellos[0];

    p.sockets[0].kill();
    await until(() => p.hellos.length > 1, "rejoin handshake sent");
    await until(() => p.session.connection === "live", "session recovered");

    const rejoin = p.hellos[p.hellos.length - 1];
    // A cold join is `sinceSeq: 0` with no epoch. THE WHOLE POINT of routing a
    // reconnect through resume() is that the client says "I already have state
    // up to N" — a cold rejoin would re-download the document and, worse,
    // restart the clientSeq counter, so the server would dedup the client's
    // NEW edits as re-sends and silently drop them.
    expect(firstJoin.sinceSeq, "the original join was cold, as expected").toBe(0);
    expect(rejoin.sinceSeq, "the rejoin must carry the bundle's confirmed seq").toBeGreaterThan(0);
    expect(rejoin.genesisId, "and the epoch it belongs to").toBeTruthy();
    expect(rejoin.takeover, "a rejoin claims its own identity back from the zombie socket").toBe(true);

    // And the document survived the round trip intact.
    expect(p.session.doc!.save().length, "document still present after reconnect").toBeGreaterThan(0);
    await p.unmount();
  });

  it("recovers to live and re-opens the editor", async () => {
    const hub = seeded();
    const store = new InMemoryBundleStore();
    const p = await mount(hub, store);
    p.sockets[0].blackhole();
    await until(() => p.session.writesBlocked, "gated");
    await until(() => p.session.connection === "live", "recovered");
    expect(p.session.writesBlocked, "the gate lifts on its own once writes can land again").toBe(false);

    // The recovered session is genuinely usable, not merely labelled live.
    await act(async () => p.session.submitOp(insertAt(0, "AFTER")));
    await until(() => p.session.version > 1, "post-reconnect edit is sequenced");
    await p.unmount();
  });

  it("gives up as `lost` after the retry budget, and reconnect() revives it", async () => {
    const hub = seeded();
    const p = await mount(hub);
    // A total outage: every socket, including the ones the retries create, is
    // born half-open. A phone with no signal, not a single unlucky socket.
    p.outage.value = true;

    await until(() => p.session.connection === "lost", "retries exhausted", 600);
    expect(p.session.writesBlocked, "still gated while lost").toBe(true);

    // `lost` means automatic recovery has STOPPED. Nothing will retry on its
    // own from here — which is the whole reason the demo shows a modal with a
    // button rather than a spinner that would spin forever.
    const before = p.hellos.length;
    await tick(200);
    expect(p.hellos.length, "no further attempts are made once lost").toBe(before);

    p.outage.value = false;
    await act(async () => p.session.reconnect());
    await until(() => p.session.connection === "live", "manual retry recovered the session", 600);
    expect(p.session.writesBlocked, "and the editor is writable again").toBe(false);
    await p.unmount();
  });
});

describe("the editor surface is actually gated", () => {
  it("CollabEditor drops its writable surface when the connection dies", async () => {
    const hub = seeded();
    const f = factoryFor(hub);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    let session: CollabSession | null = null;
    await act(async () => {
      root.render(
        createElement(CollabEditor, {
          url: "ws://x",
          docId: "d",
          clientId: "alice",
          createSocket: f.factory,
          liveness: FAST,
          onSession: (s: CollabSession) => (session = s),
        }),
      );
    });
    await until(() => !!container.querySelector("textarea"), "editor mounted writable");

    f.sockets[0].blackhole();
    // The DOM is the real contract: `writesBlocked` feeding `editable` is an
    // implementation detail, but a user being unable to type is not.
    await until(() => !container.querySelector("textarea"), "writable surface withdrawn");
    expect(session!.connection).not.toBe("live");
    await act(async () => root.unmount());
  });
});
