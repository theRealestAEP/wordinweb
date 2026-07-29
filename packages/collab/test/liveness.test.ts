import { describe, expect, it } from "vitest";
import { LivenessMonitor, monitorTransport } from "../src/liveness.js";
import type { ClientTransport } from "../src/connection.js";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";

/**
 * The heartbeat that catches a HALF-OPEN socket.
 *
 * The bug these pin: a phone browser sleeps, its WebSocket dies at the far
 * end, and nothing local notices — `onclose` never fires, `onerror` never
 * fires, `readyState` still reads OPEN, and every `send()` succeeds into a
 * void. The editor stays editable, the user keeps typing, and the edits go
 * nowhere. Only an unanswered round trip can detect that, which is why these
 * tests never simulate a close event: a monitor that only worked when the
 * socket announced its own death would pass a close-driven test and still
 * leave the motivating scenario completely unhandled.
 *
 * Time is injected (matching BundlePersister's harness) so "the deadline
 * passed" is a deterministic statement rather than a sleep.
 */

/** Fake clock + timer queue: nothing fires until `advance` says so. */
function clock() {
  let now = 0;
  const timers: { at: number; fn: () => void }[] = [];
  return {
    opts: {
      setTimer: (fn: () => void, ms: number) => {
        const t = { at: now + ms, fn };
        timers.push(t);
        return t;
      },
      clearTimer: (t: unknown) => {
        const i = timers.indexOf(t as never);
        if (i >= 0) timers.splice(i, 1);
      },
    },
    /**
     * Move the clock forward, firing timers IN CHRONOLOGICAL ORDER and moving
     * `now` to each one's due time as it runs.
     *
     * Not `now += ms` then sweep: a probe firing inside the jump arms its
     * deadline relative to `now`, so a clock already teleported to the end of
     * the span would place that deadline in the future and it would never
     * fire. That artifact made a genuinely dead socket look alive in the test
     * while the production monitor was correct.
     */
    advance(ms: number) {
      const target = now + ms;
      for (;;) {
        const due = timers.filter((t) => t.at <= target).sort((a, b) => a.at - b.at)[0];
        if (!due) break;
        timers.splice(timers.indexOf(due), 1);
        now = Math.max(now, due.at);
        due.fn();
      }
      now = target;
    },
    get pending() {
      return timers.length;
    },
  };
}

function harness(opts: { intervalMs?: number; timeoutMs?: number } = {}) {
  const c = clock();
  const sent: ClientMessage[] = [];
  let deaths = 0;
  const m = new LivenessMonitor(
    (msg) => sent.push(msg),
    () => deaths++,
    { intervalMs: 1000, timeoutMs: 500, ...opts, ...c.opts },
  );
  m.start();
  const pings = () => sent.filter((s) => s.t === "ping") as { t: "ping"; nonce: number }[];
  return { m, sent, pings, advance: c.advance, deaths: () => deaths, pending: () => c.pending };
}

describe("LivenessMonitor — half-open detection", () => {
  it("a socket that answers stays alive indefinitely (no false positives)", () => {
    const h = harness();
    for (let i = 0; i < 20; i++) {
      h.advance(1000); // the interval elapses and a probe goes out
      const last = h.pings()[h.pings().length - 1];
      h.m.noteInbound({ t: "pong", nonce: last.nonce }); // answered in time
    }
    expect(h.pings().length, "one probe per interval, re-armed after each answer").toBe(20);
    expect(h.deaths(), "an answering socket is never declared dead").toBe(0);
    // And the deadline really was live each round rather than never armed:
    // stop answering and the very next one kills it.
    h.advance(1000);
    h.advance(500);
    expect(h.deaths(), "the same monitor dies as soon as an answer is missed").toBe(1);
  });

  it("HALF-OPEN: no close event, no error, no pong — the heartbeat alone notices", () => {
    const h = harness();
    // Nothing is ever delivered to noteInbound: this is the sleeping phone.
    // The socket object would still report OPEN and send() would still succeed.
    h.advance(1000);
    expect(h.pings().length, "a probe went out").toBe(1);
    expect(h.deaths(), "the deadline has not passed yet — no premature verdict").toBe(0);
    h.advance(500);
    expect(h.deaths(), "an unanswered probe is the only available death signal").toBe(1);
  });

  it("declares death exactly once and stops probing afterwards", () => {
    const h = harness();
    h.advance(1_000_000);
    expect(h.deaths(), "one verdict per monitor, not one per elapsed interval").toBe(1);
    const after = h.pings().length;
    h.advance(1_000_000);
    expect(h.pings().length, "a dead monitor stops probing (the reconnect owns what follows)").toBe(after);
  });

  it("a STALE pong does not satisfy the probe in flight", () => {
    const h = harness();
    h.advance(1000);
    const live = h.pings()[0].nonce;
    // A pong for a probe that is not the current one — e.g. one that arrived
    // late, after its own deadline already passed. Without nonce matching it
    // would silently reset the deadline and mask a genuinely dead socket.
    h.m.noteInbound({ t: "pong", nonce: live + 99 });
    h.advance(500);
    expect(h.deaths(), "only the answer to the CURRENT probe counts").toBe(1);
  });

  it("ONLY a pong proves life — inbound broadcasts do not", () => {
    const h = harness();
    h.advance(1000);
    // A room where broadcasts still arrive but our sends no longer land: the
    // inbound half works, the outbound half is dead. Treating any inbound
    // frame as proof of life (the usual heartbeat shortcut) would call this
    // healthy — and it is precisely the user-visible bug, since the screen
    // keeps updating with everyone else's edits while ours go nowhere.
    h.m.noteInbound({ t: "broadcast", entries: [] } as ServerMessage);
    h.m.noteInbound({ t: "roster", roster: [] } as ServerMessage);
    h.advance(500);
    expect(h.deaths(), "a working inbound half does not prove writes can land").toBe(1);
  });

  it("probe() answers a wake-up immediately instead of waiting for the interval", () => {
    const h = harness({ intervalMs: 15_000 });
    h.advance(10); // nowhere near the interval
    expect(h.pings().length).toBe(0);
    h.m.probe(); // visibilitychange
    expect(h.pings().length, "the wake-up probe does not wait out the interval").toBe(1);
  });

  it("probe() while one is in flight does not extend its deadline", () => {
    const h = harness();
    h.m.probe();
    h.advance(400);
    h.m.probe(); // a second focus event just before the deadline
    h.m.probe();
    expect(h.pings().length, "no duplicate probes").toBe(1);
    h.advance(100);
    expect(h.deaths(), "repeated tab-switching cannot postpone the verdict forever").toBe(1);
  });

  it("stop() leaves no timers and no verdict behind", () => {
    const h = harness();
    h.advance(1000);
    h.m.stop();
    expect(h.pending(), "an unmounted session must not leave a probe armed").toBe(0);
    h.advance(1_000_000);
    expect(h.deaths(), "a stopped monitor never reports").toBe(0);
  });
});

describe("monitorTransport — teeing without stealing the handler", () => {
  function fake(): { inner: ClientTransport; sent: ClientMessage[]; deliver: (m: ServerMessage) => void } {
    const sent: ClientMessage[] = [];
    let cb: ((m: ServerMessage) => void) | null = null;
    return {
      inner: { send: (m) => sent.push(m), onMessage: (h) => (cb = h) },
      sent,
      deliver: (m) => cb?.(m),
    };
  }

  it("forwards every real message to the consumer and swallows only pongs", () => {
    const f = fake();
    const { transport } = monitorTransport(f.inner, () => {});
    const seen: ServerMessage[] = [];
    transport.onMessage((m) => seen.push(m));
    f.deliver({ t: "broadcast", entries: [] } as ServerMessage);
    f.deliver({ t: "pong", nonce: 1 });
    f.deliver({ t: "roster", roster: [] } as ServerMessage);
    expect(seen.map((m) => m.t), "the document channel is untouched; pong is consumed").toEqual([
      "broadcast",
      "roster",
    ]);
  });

  it("the monitor sees pongs that the consumer never does", () => {
    const c = clock();
    const f = fake();
    let deaths = 0;
    const { transport, monitor } = monitorTransport(f.inner, () => deaths++, {
      intervalMs: 1000,
      timeoutMs: 500,
      ...c.opts,
    });
    // The connection registers ITS handler — the single-handler contract of
    // ClientTransport means a monitor that called onMessage itself would have
    // silently replaced this one and stopped the document updating.
    const seen: ServerMessage[] = [];
    transport.onMessage((m) => seen.push(m));
    monitor.start();
    c.advance(1000);
    const nonce = (f.sent.find((s) => s.t === "ping") as { nonce: number }).nonce;
    f.deliver({ t: "pong", nonce });
    c.advance(500);
    expect(deaths, "the teed pong reached the monitor").toBe(0);
    expect(seen.length, "and never reached the consumer").toBe(0);
  });
});
