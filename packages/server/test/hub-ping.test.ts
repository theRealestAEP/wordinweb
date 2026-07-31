import { describe, expect, it } from "vitest";
import { CollabHub, Connection } from "../src/hub.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";
import type { ServerMessage } from "@wordinweb/collab/server";
import { blankDocxBytes } from "../src/blank.js";

/**
 * The server half of the liveness probe.
 *
 * Two properties matter here and they pull in opposite directions, which is
 * why they are pinned separately:
 *
 *  1. A ping is ALWAYS answered. The client is required to read an unanswered
 *     probe as a dead socket, so any condition under which the server stays
 *     silent becomes a false disconnect on a perfectly healthy session.
 *  2. A ping is NEVER activity. The idle timeout exists to reap rooms nobody
 *     is using, and a heartbeat that reset its clock would make every
 *     abandoned tab immortal — the textbook way an activity-based timeout is
 *     defeated, and the hub's own notes say so.
 */

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
  last(): ServerMessage {
    for (let i = this.received.length - 1; i >= 0; i--) {
      const t = this.received[i].t;
      if (t !== "roster" && t !== "checkpointer") return this.received[i];
    }
    return this.received[this.received.length - 1];
  }
}

const hello = (clientId: string, over: Record<string, unknown> = {}) =>
  ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, ...over }) as never;

const seeded = () => {
  const hub = new CollabHub(null);
  hub.seed("d", blankDocxBytes());
  return hub;
};

/** The room's idle clock, read straight off the hub. */
const lastActivity = (hub: CollabHub): number =>
  (hub as unknown as { rooms: Map<string, { lastActivityAt: number }> }).rooms.get("d")!.lastActivityAt;

describe("hub ping/pong", () => {
  it("answers a ping with a pong carrying the same nonce", async () => {
    const hub = seeded();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    await hub.handle(a, { t: "ping", nonce: 4242 });
    expect(a.last()).toEqual({ t: "pong", nonce: 4242 });
  });

  it("echoes each nonce distinctly, so a client can match answer to probe", async () => {
    const hub = seeded();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    await hub.handle(a, { t: "ping", nonce: 1 });
    await hub.handle(a, { t: "ping", nonce: 2 });
    const pongs = a.received.filter((m) => m.t === "pong");
    expect(pongs, "a stale answer must be distinguishable from the current one").toEqual([
      { t: "pong", nonce: 1 },
      { t: "pong", nonce: 2 },
    ]);
  });

  it("answers even BEFORE hello — liveness is not a privilege", async () => {
    const hub = seeded();
    const a = new FakeConn("a");
    // No hello: this socket has no room, no identity and no authorisation.
    // It is still entitled to know whether it is connected, and gating the
    // answer would make an unjoined-but-healthy socket indistinguishable from
    // a dead one — the client would tear down a connection that works.
    await hub.handle(a, { t: "ping", nonce: 7 });
    expect(a.last()).toEqual({ t: "pong", nonce: 7 });
  });

  it("a ping is NOT activity: the idle countdown keeps running", async () => {
    let now = 0;
    const hub = new CollabHub(null, undefined, undefined, () => now);
    hub.seed("d", blankDocxBytes());
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice")); // a join IS activity
    const joinedAt = lastActivity(hub);

    now = 60_000;
    await hub.handle(a, { t: "ping", nonce: 1 });
    expect(a.last().t, "still answered").toBe("pong");
    expect(lastActivity(hub), "a heartbeat must not hold an abandoned room open").toBe(joinedAt);

    // Contrast: a real edit DOES move the clock, so the assertion above is
    // about pings specifically and not about a clock that never moves.
    now = 90_000;
    await hub.handle(a, {
      t: "submit",
      intent: { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "x", clientId: "alice", clientSeq: 1, base: 0 },
    } as never);
    expect(lastActivity(hub), "an accepted edit is the canonical activity").toBe(90_000);
  });

  it("pings do not consume the submit rate-limit budget", async () => {
    const hub = seeded();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    // Far more pings than the 800-token burst. If they drew from the same
    // bucket, a heavy typist's heartbeat could be dropped and the client would
    // diagnose a false disconnect mid-session.
    for (let i = 0; i < 2000; i++) await hub.handle(a, { t: "ping", nonce: i });
    const pongs = a.received.filter((m) => m.t === "pong");
    expect(pongs.length, "every probe answered").toBe(2000);
    // And a submit right after is still accepted.
    await hub.handle(a, {
      t: "submit",
      intent: { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "y", clientId: "alice", clientSeq: 1, base: 0 },
    } as never);
    expect(a.last().t, "the edit path is unaffected by heartbeat volume").not.toBe("refused");
  });
});

describe("protocol version fence (the ping/pong bump)", () => {
  it("refuses the PREVIOUS protocol version at hello", async () => {
    const hub = seeded();
    const a = new FakeConn("a");
    // This is what makes "no pong ⇒ dead socket" safe to assert. A v2 client
    // would talk to a server that answers pings; a v2 SERVER would ignore
    // them and a v3 client would read the silence as death on a healthy
    // session. The handshake fence is what guarantees that pairing never
    // exists — so it must actually reject, not merely differ.
    await hub.handle(a, hello("alice", { protocolVersion: PROTOCOL_VERSION - 1 }));
    expect(a.last()).toEqual({ t: "refused", reason: "version-mismatch" });
  });

  it("PROTOCOL_VERSION includes the ping/pong bump", () => {
    expect(PROTOCOL_VERSION, "ping/pong requires protocol version 3 or newer").toBeGreaterThanOrEqual(3);
  });
});
