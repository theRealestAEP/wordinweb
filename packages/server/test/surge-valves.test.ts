import { describe, expect, it } from "vitest";
import { CollabHub, type Connection } from "../src/hub.js";
import { normalizeLimits, mediaWireCap, DEFAULT_LIMITS, envInt } from "../src/limits.js";
import { MAX_IMAGE_BYTES } from "@wordinweb/collab/server";
import { IpGuard } from "../src/ip-guard.js";
import { MetricsObservability } from "../src/observability.js";
import { handleSeedRequest } from "../src/seed-http.js";
import { blankDocxBytes } from "../src/blank.js";
import { attachWebSocketServer, type WsServer, type WsSocket, type WsRequest } from "../src/ws.js";
import { PROTOCOL_VERSION, type ServerMessage } from "@wordinweb/collab/server";

/**
 * SURGE VALVES — the limits that bound what an ALREADY-ADMITTED session can
 * spend.
 *
 * The per-IP caps answer "how many things may one address start". Every valve
 * here closes a hole where that question is answered correctly, every existing
 * limit is satisfied, and the server still dies. They are the difference
 * between surviving a mixed legitimate/hostile load and surviving a polite one.
 */

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
}

const CHECKPOINT = { seq: 0, iv: "aXY=", ciphertext: "Y2lwaGVy" };

async function joinEnc(hub: CollabHub, conn: FakeConn, clientId: string, docId = "d"): Promise<void> {
  await hub.handle(conn, {
    t: "hello",
    protocolVersion: PROTOCOL_VERSION,
    docId,
    clientId,
    sinceSeq: 0,
    engineVersion: "e4",
  });
}

describe("per-room log byte budget", () => {
  /** Envelopes big enough that a handful crosses a deliberately tiny budget. */
  function envelope(clientId: string, clientSeq: number, bytes: number) {
    return { clientId, clientSeq, base: 0, iv: "aXY=", ciphertext: "x".repeat(bytes) };
  }

  it("ends the session when one room's retained log outgrows its budget", async () => {
    const limits = normalizeLimits({ surge: { roomLogMaxBytes: 1000 } });
    const obs = new MetricsObservability({ level: "silent", out: () => {} });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, obs, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await joinEnc(hub, a, "alice");
    await joinEnc(hub, b, "bob");

    // Two 400-byte envelopes: inside budget, room healthy.
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 1, 400) });
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 2, 400) });
    expect(hub.activeDocs()).toEqual(["d"]);

    // The third crosses 1000 bytes. Everyone is kicked with the distinct
    // reason and the room is deleted — this is not a rate limit that refuses
    // one message, it is a room that has stopped being a document.
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 3, 400) });
    expect(a.received.at(-1)).toEqual({ t: "refused", reason: "log-overflow" });
    expect(b.received.at(-1)).toEqual({ t: "refused", reason: "log-overflow" });
    expect(hub.activeDocs()).toEqual([]);

    const snap = obs.snapshot();
    expect(snap.evictionsByReason).toEqual({ "log-overflow": 1 });
    expect(snap.kicksByReason["log-overflow"]).toBe(2);
  });

  it("the envelope that crosses the line is still broadcast before the room ends", async () => {
    // The room is ending either way; withholding the last sequenced entry
    // would leave peers with a tail the server already numbered.
    const limits = normalizeLimits({ surge: { roomLogMaxBytes: 500 } });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const a = new FakeConn("a");
    await joinEnc(hub, a, "alice");
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 1, 600) });
    const kinds = a.received.map((m) => m.t);
    expect(kinds).toContain("broadcast-enc");
    expect(kinds.indexOf("broadcast-enc")).toBeLessThan(kinds.lastIndexOf("refused"));
  });

  it("a checkpoint that prunes the log resets the budget", async () => {
    // Otherwise a long, healthy, heavily-checkpointed session would eventually
    // trip a limit meant for a room that never goes quiescent.
    const limits = normalizeLimits({ surge: { roomLogMaxBytes: 1000 } });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const a = new FakeConn("a");
    await joinEnc(hub, a, "alice");
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 1, 400) });
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 2, 400) });
    // Quiescent checkpoint at the log head prunes everything before it.
    await hub.handle(a, { t: "checkpoint", checkpoint: { seq: 2, iv: "aXY=", ciphertext: "bmV3" } });
    // With the budget reset, two more 400-byte envelopes are fine.
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 3, 400) });
    await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", 4, 400) });
    expect(hub.activeDocs()).toEqual(["d"]);
  });

  it("0 disables the budget", async () => {
    const limits = normalizeLimits({ surge: { roomLogMaxBytes: 0 } });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const a = new FakeConn("a");
    await joinEnc(hub, a, "alice");
    for (let i = 1; i <= 20; i++) {
      await hub.handle(a, { t: "submit-enc", envelope: envelope("alice", i, 5000) });
    }
    expect(hub.activeDocs()).toEqual(["d"]);
  });
});

describe("global ceilings", () => {
  it("refuses new rooms with 503 server-full, leaving existing rooms untouched", () => {
    const limits = normalizeLimits({ surge: { maxRooms: 2 }, ip: { seedPerMin: 0, maxDocsPerIp: 0 } });
    const obs = new MetricsObservability({ level: "silent", out: () => {} });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, obs, limits);
    const docx = Buffer.from(blankDocxBytes()).toString("base64");
    let n = 0;
    const seed = () => handleSeedRequest(hub, { method: "POST", body: { docx } }, { mintDocId: () => `d${n++}` });

    expect(seed().status).toBe(201);
    expect(seed().status).toBe(201);
    const full = seed();
    // 503, NOT 429: this is not the caller's fault and backing off harder does
    // not help them specifically. The status is the difference between "add
    // capacity" and "find the abuser".
    expect(full.status).toBe(503);
    expect(full.body).toEqual({ error: "server-full" });
    // The sessions already running are entirely unaffected — that is the whole
    // point of a fail-fast valve.
    expect(hub.activeDocs()).toHaveLength(2);
    expect(obs.snapshot().capacityByReason).toEqual({ "server-full": 1 });
    // Capacity is not a quota: it returns as rooms go away.
    hub.sweepRooms(0);
  });

  it("refuses sockets past the global ceiling at accept", () => {
    const obs = new MetricsObservability({ level: "silent", out: () => {} });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, obs);
    const wss = new FakeWss();
    attachWebSocketServer(wss, hub, obs, { maxConns: 2 });

    const s1 = wss.connect("1.1.1.1");
    const s2 = wss.connect("2.2.2.2");
    // A DIFFERENT address each time — per-IP caps are no defence against a
    // botnet, which is exactly why this ceiling exists.
    const s3 = wss.connect("3.3.3.3");
    expect(s1.closed).toBe(false);
    expect(s2.closed).toBe(false);
    expect(s3.closed).toBe(true);
    expect(s3.frames()).toEqual([{ t: "refused", reason: "server-busy" }]);
    expect(obs.snapshot().capacityByReason).toEqual({ "server-busy": 1 });

    // A socket closing frees a slot; the ceiling is concurrency, not a quota.
    s1.handlers.close?.();
    expect(wss.connect("4.4.4.4").closed).toBe(false);
  });
});

describe("slow-consumer backpressure", () => {
  /**
   * Driven through a REAL broadcast rather than by calling a send wrapper the
   * test built itself: the backpressure check lives inside the Connection the
   * adapter constructs, so the only honest way to exercise it is to make the
   * hub fan out to that Connection. Two sockets join an encrypted room and one
   * submits; the other is the one that has stopped reading.
   */
  async function room(maxBufferedBytes: number) {
    const obs = new MetricsObservability({ level: "silent", out: () => {} });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, obs);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const wss = new FakeWss();
    attachWebSocketServer(wss, hub, obs, { maxBufferedBytes });
    const writer = wss.connect("1.1.1.1");
    const reader = wss.connect("2.2.2.2");
    await deliver(writer, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "w", sinceSeq: 0, engineVersion: "e4" });
    await deliver(reader, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "r", sinceSeq: 0, engineVersion: "e4" });
    return { hub, obs, writer, reader };
  }

  it("terminates a peer that stops draining, and leaves the healthy peer alone", async () => {
    const { obs, writer, reader } = await room(1000);
    const before = reader.sent.length;

    // The reader stops reading; its queue climbs past the ceiling.
    reader.bufferedAmount = 5000;
    await deliver(writer, {
      t: "submit-enc",
      envelope: { clientId: "w", clientSeq: 1, base: 0, iv: "aXY=", ciphertext: "aGVsbG8=" },
    });

    // The broadcast reached the healthy socket and terminated the stuck one
    // instead of queueing more into a buffer nobody is draining.
    expect(reader.sent.length).toBe(before);
    expect(reader.terminated).toBe(true);
    expect(writer.terminated).toBe(false);
    expect(writer.frames().some((m) => m.t === "broadcast-enc")).toBe(true);
    expect(obs.snapshot().kicksByReason).toEqual({ backpressure: 1 });

    // It stays dead: a later fan-out is a no-op, not a second termination.
    await deliver(writer, {
      t: "submit-enc",
      envelope: { clientId: "w", clientSeq: 2, base: 0, iv: "aXY=", ciphertext: "aGVsbG8=" },
    });
    expect(reader.sent.length).toBe(before);
    expect(obs.snapshot().kicksByReason).toEqual({ backpressure: 1 });
  });

  it("0 disables the ceiling", async () => {
    const { reader, writer } = await room(0);
    const before = reader.sent.length;
    reader.bufferedAmount = 999_999_999;
    await deliver(writer, {
      t: "submit-enc",
      envelope: { clientId: "w", clientSeq: 1, base: 0, iv: "aXY=", ciphertext: "aGVsbG8=" },
    });
    expect(reader.terminated).toBe(false);
    expect(reader.sent.length).toBeGreaterThan(before);
  });
});

describe("media limits", () => {
  const bytesOf = (n: number) => new Uint8Array(n);

  async function seededHub(maxBlobBytes: number) {
    const limits = normalizeLimits({ media: { maxBlobBytes }, ip: { seedPerMin: 0, maxDocsPerIp: 0 } });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    return { hub, limits };
  }

  it("accepts a blob at the cap and refuses one byte over it", async () => {
    const { hub } = await seededHub(1024);
    // Content-addressed: the sha must match, so hash what we send.
    const at = bytesOf(1024);
    const over = bytesOf(1025);
    expect(await hub.mediaUpload("d", await sha256Hex(at), at)).toBe(201);
    expect(await hub.mediaUpload("d", await sha256Hex(over), over)).toBe(413);
  });

  it("THE COUPLING: the socket-level guard always sits above the blob cap", () => {
    // The ordering bug this pins is invisible until someone RAISES the cap:
    // if the wire guard ever falls at or below it, a legal upload is destroyed
    // mid-flight and the client sees a dropped connection instead of a 413
    // saying how big is too big. Checked across the range an operator might
    // plausibly configure, including well past the default.
    for (const maxBlobBytes of [1, 1024, 5 * 1024 * 1024, 10 * 1024 * 1024, 64 * 1024 * 1024]) {
      const limits = normalizeLimits({ media: { maxBlobBytes } });
      expect(mediaWireCap(limits.media)).toBeGreaterThan(limits.media.maxBlobBytes);
    }
  });

  it("a staged, never-referenced blob evicts on the SHORT configured clock", async () => {
    let now = 0;
    const limits = normalizeLimits({ media: { stageTtlMs: 100, ttlMs: 10_000 } });
    const hub = new CollabHub(null, undefined, undefined, () => now, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const blob = bytesOf(64);
    const sha = await sha256Hex(blob);
    expect(await hub.mediaUpload("d", sha, blob)).toBe(201);
    // Deliberately NOT probed in the middle: `mediaDownload` PROMOTES a staged
    // blob, so a presence check would move it onto the long clock and the test
    // would be measuring its own probe. (The first draft of this test did
    // exactly that and failed for that reason.)
    now += 101;
    hub.sweepMedia();
    expect(hub.mediaDownload("d", sha)).toBeNull();
  });

  it("a download promotes a staged blob onto the LONG clock", async () => {
    // The behaviour that broke the test above, pinned deliberately: an upload
    // precedes the intent referencing it, and the download IS the observable
    // claim that somebody wants it (in encrypted rooms the server cannot read
    // the referencing intent, so this is the only signal it gets).
    let now = 0;
    const limits = normalizeLimits({ media: { stageTtlMs: 100, ttlMs: 10_000 } });
    const hub = new CollabHub(null, undefined, undefined, () => now, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", CHECKPOINT);
    const blob = bytesOf(64);
    const sha = await sha256Hex(blob);
    await hub.mediaUpload("d", sha, blob);
    expect(hub.mediaDownload("d", sha)).not.toBeNull(); // promotes
    now += 101; // well past the STAGE ttl…
    hub.sweepMedia();
    expect(hub.mediaDownload("d", sha)).not.toBeNull(); // …and it survives
    now += 10_001; // past the promoted ttl, with no further downloads
    hub.sweepMedia();
    expect(hub.mediaDownload("d", sha)).toBeNull();
  });

  it("CLAMPS a configured cap above the wire ceiling, so the relay cannot accept what the wire refuses", async () => {
    // The bug this closes: the relay accepted an upload the validator then
    // rejected as "bad size" — identically on every replica, so no image, no
    // error, nothing to diagnose. Configuring above the ceiling is now
    // impossible rather than merely discouraged.
    const absurd = normalizeLimits({ media: { maxBlobBytes: MAX_IMAGE_BYTES * 4 } });
    expect(absurd.media.maxBlobBytes).toBe(MAX_IMAGE_BYTES);
    // Every construction path, not just the environment one: an embedder or a
    // test wiring limits directly used to bypass the clamp entirely, which
    // left the window open for exactly the callers least likely to notice.
    expect(normalizeLimits().media.maxBlobBytes).toBeLessThanOrEqual(MAX_IMAGE_BYTES);
    // Asserted on numbers alone — allocating a ceiling-sized buffer here would
    // burn 256MB to re-prove what the 1KB cap test above already covers.
  });

  it("a cap BELOW the ceiling is left exactly as configured", () => {
    // The clamp must not become a floor: dev's 50MB and production's 5MB are
    // both legitimate and must survive untouched.
    expect(normalizeLimits({ media: { maxBlobBytes: 50 * 1024 * 1024 } }).media.maxBlobBytes).toBe(50 * 1024 * 1024);
    expect(normalizeLimits({ media: { maxBlobBytes: 5 * 1024 * 1024 } }).media.maxBlobBytes).toBe(5 * 1024 * 1024);
  });

  it("ships the documented defaults, including the user's 5 MB public cap", () => {
    expect(DEFAULT_LIMITS.media.maxBlobBytes).toBe(5 * 1024 * 1024);
    expect(DEFAULT_LIMITS.media.stageTtlMs).toBe(60_000);
    expect(DEFAULT_LIMITS.media.ttlMs).toBe(300_000);
    expect(DEFAULT_LIMITS.surge.wsMaxPayloadBytes).toBe(512 * 1024);
    expect(DEFAULT_LIMITS.surge.wsMaxBufferedBytes).toBe(4 * 1024 * 1024);
    expect(DEFAULT_LIMITS.surge.roomLogMaxBytes).toBe(64 * 1024 * 1024);
    expect(DEFAULT_LIMITS.surge.maxRooms).toBe(2000);
    expect(DEFAULT_LIMITS.surge.maxConns).toBe(5000);
  });

  it("the frame cap comfortably exceeds the largest legal envelope", () => {
    // 256 KB of ciphertext plus base64 and JSON overhead has to fit, or the
    // frame cap silently rejects traffic the envelope cap deliberately allows
    // — a contradiction between two limits that would look like random
    // disconnects under heavy paste.
    expect(DEFAULT_LIMITS.surge.wsMaxPayloadBytes).toBeGreaterThan(256 * 1024);
  });

  it("reads every new knob from the environment", () => {
    const saved = { ...process.env };
    try {
      process.env.WW_MEDIA_MAX_BLOB_BYTES = "777";
      expect(envInt("WW_MEDIA_MAX_BLOB_BYTES", 1)).toBe(777);
    } finally {
      process.env = saved;
    }
  });
});

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as unknown as ArrayBuffer);
  let hex = "";
  for (const b of new Uint8Array(digest)) hex += b.toString(16).padStart(2, "0");
  return hex;
}

/** Minimal ws-shaped fakes; the adapter is structural by design. */
class FakeSocket implements WsSocket {
  sent: string[] = [];
  closed = false;
  terminated = false;
  bufferedAmount = 0;
  handlers: { message?: (d: unknown) => void; close?: () => void } = {};
  send(data: string): void {
    this.sent.push(data);
  }
  on(event: "message" | "close", cb: never): void {
    if (event === "message") this.handlers.message = cb;
    else this.handlers.close = cb;
  }
  close(): void {
    this.closed = true;
  }
  terminate(): void {
    this.terminated = true;
    this.closed = true;
  }
  frames(): ServerMessage[] {
    return this.sent.map((s) => JSON.parse(s) as ServerMessage);
  }
}

class FakeWss implements WsServer {
  private cb?: (socket: WsSocket, req?: WsRequest) => void;
  on(_event: "connection", cb: (socket: WsSocket, req?: WsRequest) => void): void {
    this.cb = cb;
  }
  connect(peer: string): FakeSocket {
    const s = new FakeSocket();
    this.cb?.(s, { socket: { remoteAddress: peer }, headers: {} });
    return s;
  }
}

/**
 * Push a client frame in through the socket's own message handler — the same
 * door a real peer uses — and let the adapter's `void hub.handle(...)` settle.
 * Going through the handler rather than calling `hub.handle` directly is what
 * makes these tests exercise the ADAPTER (frame parsing, the Connection
 * wrapper, the backpressure check) instead of only the hub.
 */
async function deliver(socket: FakeSocket, msg: unknown): Promise<void> {
  socket.handlers.message?.(JSON.stringify(msg));
  await new Promise((r) => setTimeout(r, 0));
}
