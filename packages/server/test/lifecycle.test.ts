import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection, DocProvider } from "../src/hub.js";
import { ServerMessage, PROTOCOL_VERSION } from "@wordinweb/collab/server";
import { normalizeLimits } from "../src/limits.js";
import { MetricsObservability } from "../src/observability.js";

/**
 * SERVER LIFECYCLE — the idle timeout and the absolute room lifetime.
 *
 * Every deadline here is driven by the hub's INJECTED CLOCK, so a twelve-hour
 * cap is tested by assigning to a variable. Nothing sleeps, nothing is
 * timing-dependent, and the tests state the policy rather than approximating
 * it: a test that waits on wall time for a 10-minute timeout is a test nobody
 * runs.
 */

function blankDoc(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
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

const provider: DocProvider = { load: () => blankDoc("hi") };

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
  ofType<T extends ServerMessage["t"]>(t: T): Extract<ServerMessage, { t: T }>[] {
    return this.received.filter((m) => m.t === t) as Extract<ServerMessage, { t: T }>[];
  }
}

const IDLE = 10 * 60_000;
const IDLE_WARN = 60_000;
const LIFETIME = 12 * 60 * 60_000;
const LIFETIME_WARN = 5 * 60_000;

const LIMITS = normalizeLimits({
  lifecycle: {
    idleTimeoutMs: IDLE,
    idleWarningMs: IDLE_WARN,
    roomMaxLifetimeMs: LIFETIME,
    lifetimeWarningMs: LIFETIME_WARN,
  },
});

/**
 * The cap in isolation, with the idle timeout switched off. Used by the tests
 * that are ABOUT the lifetime deadline: with both clocks running, a room left
 * quiet for twelve hours idles out long before the cap, and the test would
 * pass or fail for reasons that have nothing to do with what it claims to
 * check (the first draft of two of these did exactly that).
 */
const LIFETIME_ONLY = normalizeLimits({
  lifecycle: { idleTimeoutMs: 0, roomMaxLifetimeMs: LIFETIME, lifetimeWarningMs: LIFETIME_WARN },
});

/** A hub with the lifecycle limits above and a caller-controlled clock. */
function makeHub(now: () => number, limits = LIMITS) {
  return new CollabHub(provider, undefined, undefined, now, undefined, undefined, limits);
}

async function join(hub: CollabHub, conn: FakeConn, clientId: string, docId = "d"): Promise<void> {
  await hub.handle(conn, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId, clientId, sinceSeq: 0 });
}

/** One accepted edit — the canonical "qualifying activity". */
async function edit(hub: CollabHub, conn: FakeConn, clientId: string, clientSeq: number): Promise<void> {
  await hub.handle(conn, {
    t: "submit",
    intent: { kind: "insertText", clientId, clientSeq, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "x" },
  });
}

describe("idle timeout", () => {
  it("warns at the lead time, then kicks everyone and evicts the room", async () => {
    let now = 0;
    const hub = makeHub(() => now);
    const a = new FakeConn("a");
    await join(hub, a, "alice");

    // Well inside the window: no warning, no eviction.
    now += IDLE - IDLE_WARN - 1;
    expect(hub.sweepLifecycle()).toEqual([]);
    expect(a.ofType("session-warning")).toHaveLength(0);

    // Crossing into the lead time arms the warning — once.
    now += 1;
    expect(hub.sweepLifecycle()).toEqual([]);
    const warnings = a.ofType("session-warning");
    expect(warnings).toHaveLength(1);
    expect(warnings[0].reason).toBe("idle");
    // The countdown carries the MEASURED remainder, so a client counting down
    // from it lands on the real ending rather than on a configured constant.
    expect(warnings[0].inMs).toBe(IDLE_WARN);
    // A second sweep inside the same idle stretch must not re-warn.
    now += 1000;
    hub.sweepLifecycle();
    expect(a.ofType("session-warning")).toHaveLength(1);

    // At the deadline: kicked with the distinct reason, room gone.
    now += IDLE_WARN;
    expect(hub.sweepLifecycle()).toEqual([{ docId: "d", reason: "idle-timeout" }]);
    expect(a.received.at(-1)).toEqual({ t: "refused", reason: "idle-timeout" });
    expect(hub.activeDocs()).toEqual([]);
  });

  it("an accepted edit cancels the countdown, tells the room, and resets the clock", async () => {
    let now = 0;
    const hub = makeHub(() => now);
    const a = new FakeConn("a");
    await join(hub, a, "alice");

    now += IDLE - IDLE_WARN;
    hub.sweepLifecycle();
    expect(a.ofType("session-warning")).toHaveLength(1);

    // Someone types. The countdown must visibly END — a UI left showing
    // "ending in 12s" over a healthy session is its own bug.
    await edit(hub, a, "alice", 1);
    expect(a.ofType("session-warning-cleared")).toEqual([{ t: "session-warning-cleared", reason: "idle" }]);

    // And the clock genuinely restarted: the old deadline passes untouched.
    const lastActivity = now;
    now = lastActivity + IDLE_WARN + 1;
    expect(hub.sweepLifecycle()).toEqual([]);
    expect(hub.activeDocs()).toEqual(["d"]);
    // A full fresh window later, it warns again — the second stretch is armed
    // independently of the first.
    now = lastActivity + IDLE - IDLE_WARN;
    hub.sweepLifecycle();
    expect(a.ofType("session-warning")).toHaveLength(2);
  });

  it("PRESENCE IS NOT ACTIVITY — a room of twitching forgotten tabs still dies", async () => {
    // The decision the entire timeout hangs on. A forgotten tab emits
    // presence on every mouse move and focus change; counting it would make
    // the timeout dead letter, so this is pinned rather than commented.
    let now = 0;
    const hub = makeHub(() => now);
    const a = new FakeConn("a");
    await join(hub, a, "alice");

    for (let i = 0; i < 20; i++) {
      now += 30_000;
      await hub.handle(a, { t: "presence", position: { anchor: { blockId: 1, runId: 2, offset: i } } });
      hub.sweepLifecycle();
    }
    // 10 minutes of continuous presence traffic later, the room is gone.
    expect(hub.activeDocs()).toEqual([]);
    expect(a.received.at(-1)).toEqual({ t: "refused", reason: "idle-timeout" });
  });

  it("a JOIN counts as activity, so a newcomer is not evicted out from under", async () => {
    let now = 0;
    const hub = makeHub(() => now);
    const a = new FakeConn("a");
    await join(hub, a, "alice");

    // Nine minutes of silence, then someone opens the link.
    now += IDLE - IDLE_WARN;
    hub.sweepLifecycle();
    const b = new FakeConn("b");
    await join(hub, b, "bob");

    // The newcomer's arrival reset the clock: the original deadline passes
    // and the session is still alive, with the countdown withdrawn.
    now += IDLE_WARN + 1;
    expect(hub.sweepLifecycle()).toEqual([]);
    expect(hub.activeDocs()).toEqual(["d"]);
    expect(a.ofType("session-warning-cleared")).toHaveLength(1);
  });
});

describe("absolute room lifetime", () => {
  it("warns at the lead time and ends the session at the cap, whatever the activity", async () => {
    let now = 0;
    const hub = makeHub(() => now);
    const a = new FakeConn("a");
    await join(hub, a, "alice");

    // A keepalive script: one edit every five minutes, forever. This defeats
    // the idle timeout completely — and that is exactly why the absolute cap
    // exists, so the assertion is that it does NOT save the room.
    let seq = 1;
    for (let t = 0; t < LIFETIME - LIFETIME_WARN; t += 5 * 60_000) {
      now += 5 * 60_000;
      await edit(hub, a, "alice", seq++);
      hub.sweepLifecycle();
    }
    const warnings = a.ofType("session-warning");
    expect(warnings).toHaveLength(1);
    expect(warnings[0].reason).toBe("lifetime");
    expect(warnings[0].inMs).toBeLessThanOrEqual(LIFETIME_WARN);

    // Kept editing right through the warning — the deadline does not move.
    now += LIFETIME_WARN;
    await edit(hub, a, "alice", seq++);
    expect(hub.sweepLifecycle()).toEqual([{ docId: "d", reason: "session-expired" }]);
    expect(a.received.at(-1)).toEqual({ t: "refused", reason: "session-expired" });
    expect(hub.activeDocs()).toEqual([]);
  });

  it("the lifetime warning is NOT cancellable by activity", async () => {
    let now = 0;
    const hub = makeHub(() => now, LIFETIME_ONLY);
    const a = new FakeConn("a");
    await join(hub, a, "alice");
    now = LIFETIME - LIFETIME_WARN;
    hub.sweepLifecycle();
    expect(a.ofType("session-warning")).toHaveLength(1);

    await edit(hub, a, "alice", 1);
    // Activity clears the IDLE countdown only; nothing withdraws a lifetime
    // warning, because nothing can move the deadline it announces.
    expect(a.ofType("session-warning-cleared")).toHaveLength(0);
  });

  it("a client joining inside the final window is told about the deadline too", async () => {
    let now = 0;
    const hub = makeHub(() => now, LIFETIME_ONLY);
    const a = new FakeConn("a");
    await join(hub, a, "alice");
    now = LIFETIME - LIFETIME_WARN;
    hub.sweepLifecycle(); // broadcast to whoever was present

    // Bob arrives four minutes before the end. He missed the broadcast, but
    // the warning is a property of the ROOM — a session that simply vanishes
    // on a new joiner is the failure this catches.
    now += 60_000;
    const b = new FakeConn("b");
    await join(hub, b, "bob");
    const warnings = b.ofType("session-warning");
    expect(warnings).toHaveLength(1);
    expect(warnings[0].reason).toBe("lifetime");
    expect(warnings[0].inMs).toBe(LIFETIME_WARN - 60_000);
  });

  it("lifetime wins over idle when both deadlines land in the same sweep", async () => {
    // Both expire together; the reason an operator sees must be the true
    // cause, and the cap is the one that cannot be argued with.
    let now = 0;
    const hub = makeHub(() => now);
    const a = new FakeConn("a");
    await join(hub, a, "alice");
    now = LIFETIME;
    expect(hub.sweepLifecycle()).toEqual([{ docId: "d", reason: "session-expired" }]);
  });
});

describe("lifecycle configuration", () => {
  it("0 disables a deadline entirely (limits.ts convention)", async () => {
    let now = 0;
    const hub = makeHub(
      () => now,
      normalizeLimits({ lifecycle: { idleTimeoutMs: 0, roomMaxLifetimeMs: 0, emptyRoomTtlMs: 0 } }),
    );
    const a = new FakeConn("a");
    await join(hub, a, "alice");
    now = 100 * LIFETIME;
    expect(hub.sweepLifecycle()).toEqual([]);
    // …and the empty-room sweep is disabled by the same convention.
    hub.disconnect(a);
    now += 100 * LIFETIME;
    expect(hub.sweepRooms()).toEqual([]);
    expect(hub.activeDocs()).toEqual(["d"]);
  });

  it("evictions and warnings are tallied by reason in the observability snapshot", async () => {
    let now = 0;
    const lines: string[] = [];
    const obs = new MetricsObservability({ level: "silent", out: (l) => lines.push(l), now: () => now });
    const hub = new CollabHub(provider, undefined, undefined, () => now, undefined, obs, LIMITS);
    const a = new FakeConn("a");
    await join(hub, a, "alice");

    now = IDLE - IDLE_WARN;
    hub.sweepLifecycle(); // warns
    now = IDLE;
    hub.sweepLifecycle(); // evicts

    // A second room, ended by everyone simply leaving.
    const b = new FakeConn("b");
    await join(hub, b, "bob", "d2");
    hub.disconnect(b);
    now += LIMITS.lifecycle.emptyRoomTtlMs;
    hub.sweepRooms();

    const snap = obs.snapshot();
    expect(snap.warningsByReason).toEqual({ idle: 1 });
    // "Rooms are evicting" is useless on its own; the split is the signal.
    expect(snap.evictionsByReason).toEqual({ "idle-timeout": 1, empty: 1 });
    expect(snap.counters.roomsEvicted).toBe(2);
    expect(snap.gauges.rooms).toBe(0);
  });
});
