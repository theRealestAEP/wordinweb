// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { createElement, act } from "react";
import { createRoot } from "react-dom/client";
import { useCollab, type CollabSession } from "../src/collab.js";
import { CollabHub, blankProvider, type Connection, type ServerMessage } from "@wordinweb/server";

/**
 * SESSION-ENDING COUNTDOWNS at the react layer (server lifecycle arc).
 *
 * The connection already turns the wire messages into callbacks
 * (collab/test/session-warning.test.ts); this pins what the HOOK does with
 * them, which is where the two things a countdown must never do live:
 *
 *  - outlive its session (a clock still promising an ending that has already
 *    happened, sitting beside the "this session ended" screen), and
 *  - lose an uncancellable deadline to a cancellable one (the overlap case in
 *    the third test — the failure a single warning slot produces).
 *
 * The transport is a real CollabHub over a fake socket, so the join/welcome
 * path is genuine; only the lifecycle messages are injected, because the
 * alternative is a test that waits out a real idle timeout.
 */

function harness() {
  const hub = new CollabHub(blankProvider);
  /** Push a raw ServerMessage down the client's socket (the server's timers
   * are what send these for real — nothing else about the room is faked). */
  let inject: (m: unknown) => void = () => {};
  const factory = (_url: string) => {
    const listeners: ((ev: { data: unknown }) => void)[] = [];
    const conn: Connection = {
      id: "c0",
      send: (m: ServerMessage) => listeners.forEach((l) => l({ data: JSON.stringify(m) })),
    };
    inject = (m: unknown) => listeners.forEach((l) => l({ data: JSON.stringify(m) }));
    let opened = false;
    return {
      send: (data: string) => { void hub.handle(conn, JSON.parse(data)); },
      addEventListener: (t: "message" | "open", cb: never) => {
        if (t === "message") listeners.push(cb as (ev: { data: unknown }) => void);
        else if (!opened) { opened = true; (cb as () => void)(); }
      },
    } as unknown as WebSocket;
  };
  const seen: CollabSession[] = [];
  function Probe() {
    seen.push(useCollab({ url: "ws://x", docId: "d", clientId: "a", createSocket: factory }));
    return null;
  }
  const root = createRoot(document.createElement("div"));
  return {
    root,
    Probe,
    /** Deliver, then let React settle — the state update is the thing under
     * test, so it happens inside `act` rather than racing the assertions. */
    inject: async (m: unknown) => {
      await act(async () => { inject(m); await new Promise<void>((r) => setTimeout(r, 5)); });
    },
    latest: () => seen[seen.length - 1],
  };
}

async function tick() {
  await act(async () => { await new Promise<void>((r) => setTimeout(r, 5)); });
}

async function joined() {
  const h = harness();
  await act(async () => { h.root.render(createElement(h.Probe)); });
  for (let i = 0; i < 20 && !h.latest()?.ready; i++) await tick();
  expect(h.latest().ready).toBe(true);
  return h;
}

describe("useCollab — announced session endings", () => {
  it("surfaces the countdown, and takes it down when the deadline is cancelled", async () => {
    const h = await joined();
    await h.inject({ t: "session-warning", reason: "idle", inMs: 57_000 });
    await tick();
    expect(h.latest().sessionWarning?.reason).toBe("idle");
    // Measured from the announced remainder, not invented: within a tick of
    // the number the server sent.
    expect(h.latest().sessionWarning!.inMs).toBeGreaterThan(50_000);
    expect(h.latest().sessionWarning!.inMs).toBeLessThanOrEqual(57_000);
    // A warning is NOT a refusal — the session stays fully usable through the
    // grace period, and the editor must keep rendering.
    expect(h.latest().refused).toBeNull();
    expect(h.latest().ready).toBe(true);

    await h.inject({ t: "session-warning-cleared", reason: "idle" });
    await tick();
    expect(h.latest().sessionWarning).toBeNull();
    expect(h.latest().ready).toBe(true);
    await act(async () => { h.root.unmount(); });
  });

  it("the kick that follows REPLACES the countdown — never both at once", async () => {
    const h = await joined();
    await h.inject({ t: "session-warning", reason: "idle", inMs: 60_000 });
    await tick();
    expect(h.latest().sessionWarning).not.toBeNull();
    // The deadline arriving. A countdown left standing beside the ended-session
    // screen would be promising an ending that already happened.
    await h.inject({ t: "refused", reason: "idle-timeout" });
    await tick();
    expect(h.latest().refused).toBe("idle-timeout");
    expect(h.latest().sessionWarning).toBeNull();
    await act(async () => { h.root.unmount(); });
  });

  it("clearing the idle deadline REVEALS the lifetime one still pending under it", async () => {
    const h = await joined();
    // A room near its age cap that also goes quiet: both deadlines pending.
    await h.inject({ t: "session-warning", reason: "lifetime", inMs: 300_000 });
    await tick();
    expect(h.latest().sessionWarning?.reason).toBe("lifetime");
    await h.inject({ t: "session-warning", reason: "idle", inMs: 45_000 });
    await tick();
    // The SOONER ending is the one shown — and it is the cancellable one, so
    // it is also the one worth acting on.
    expect(h.latest().sessionWarning?.reason).toBe("idle");

    // Someone types; the server takes the idle deadline back. The lifetime cap
    // has not moved — nothing can move it — so the countdown must return to it
    // rather than clearing the screen.
    await h.inject({ t: "session-warning-cleared", reason: "idle" });
    await tick();
    expect(h.latest().sessionWarning?.reason).toBe("lifetime");
    // …with the time it ACTUALLY has left, not the remainder it was announced
    // with several messages ago.
    expect(h.latest().sessionWarning!.inMs).toBeLessThanOrEqual(300_000);
    expect(h.latest().sessionWarning!.inMs).toBeGreaterThan(290_000);
    await act(async () => { h.root.unmount(); });
  });
});
