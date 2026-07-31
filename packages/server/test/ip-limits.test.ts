import { describe, expect, it } from "vitest";
import { clientIp, normalizeIp, normalizeLimits, envInt, DEFAULT_LIMITS } from "../src/limits.js";
import { IpGuard } from "../src/ip-guard.js";
import { CollabHub, type Connection } from "../src/hub.js";
import { handleSeedRequest } from "../src/seed-http.js";
import { blankDocxBytes } from "../src/blank.js";
import { attachWebSocketServer, type WsServer, type WsSocket, type WsRequest } from "../src/ws.js";
import { MetricsObservability } from "../src/observability.js";
import type { ServerMessage } from "@wordinweb/collab/server";

/**
 * PER-IP ABUSE PROTECTION — and above all, the address extraction it rests
 * on. Every limit in this file is exactly as strong as `clientIp`, so that
 * function gets the adversarial tests rather than the happy ones: a limiter
 * that works until someone sends a header is not a limiter.
 */

describe("clientIp — the trust-proxy boundary", () => {
  const withHeader = (xff: string | undefined, peer = "203.0.113.7"): WsRequest => ({
    socket: { remoteAddress: peer },
    headers: xff === undefined ? {} : { "x-forwarded-for": xff },
  });

  it("ignores X-Forwarded-For entirely when no proxy is trusted (the default)", () => {
    // A header from an untrusted peer is just a string the peer chose. With
    // hops=0 it must not influence the bucket at all.
    expect(clientIp(withHeader("1.2.3.4"), 0)).toBe("203.0.113.7");
    expect(clientIp(withHeader("1.2.3.4, 5.6.7.8"), 0)).toBe("203.0.113.7");
  });

  it("takes the RIGHT-most entry with one trusted hop — the address our own proxy observed", () => {
    // Caddy appends the peer it saw, so a single-proxy chain is
    // `<client>` and the right-most IS the client.
    expect(clientIp(withHeader("198.51.100.9"), 1)).toBe("198.51.100.9");
  });

  it("THE SPOOF CASE: a client-supplied X-Forwarded-For cannot choose its own bucket", () => {
    // The attack the right-most rule exists to defeat. A client sends
    // `X-Forwarded-For: 9.9.9.9` and Caddy forwards `9.9.9.9, <real client>`.
    // Reading the LEFT-most entry — the intuitive "the original client" —
    // would hand every request a fresh attacker-chosen bucket and silently
    // disable every per-IP limit in this file.
    const spoofed = withHeader("9.9.9.9, 198.51.100.9");
    expect(clientIp(spoofed, 1)).toBe("198.51.100.9");
    expect(clientIp(spoofed, 1)).not.toBe("9.9.9.9");

    // Even a long forged chain cannot push the real address out of position:
    // it is always exactly one from the right with one trusted hop.
    const longForgery = withHeader("1.1.1.1, 2.2.2.2, 3.3.3.3, 198.51.100.9");
    expect(clientIp(longForgery, 1)).toBe("198.51.100.9");
  });

  it("counts hops from the right, so two trusted proxies read one entry further left", () => {
    // `<client>, <cdn-observed>` arriving through our own proxy.
    expect(clientIp(withHeader("198.51.100.9, 70.70.70.70"), 2)).toBe("198.51.100.9");
  });

  it("misconfiguration degrades toward OVER-limiting, never toward no limiting", () => {
    // Trust configured but the proxy is not actually forwarding: fall back to
    // the peer (everyone shares the proxy's bucket — limits fire too often,
    // which is loud and safe) rather than to "no limit at all".
    expect(clientIp(withHeader(undefined), 1)).toBe("203.0.113.7");
    // More hops configured than entries present: take the oldest entry we
    // have. Still a bucket, still not attacker-chosen per-request.
    expect(clientIp(withHeader("198.51.100.9"), 5)).toBe("198.51.100.9");
    // No request at all (a test fake, a non-HTTP transport): empty key, which
    // the guard treats as "don't guess" rather than as one shared bucket.
    expect(clientIp(undefined, 1)).toBe("");
  });
});

describe("normalizeIp — one client, one bucket", () => {
  it("reduces IPv4-mapped IPv6 to its IPv4 form", () => {
    // Node hands back `::ffff:1.2.3.4` on a dual-stack listener; without this
    // the same client occupies two buckets depending on how it connected.
    expect(normalizeIp("::ffff:203.0.113.7")).toBe("203.0.113.7");
  });

  it("strips brackets, ports and zone ids", () => {
    expect(normalizeIp("[2001:db8::1]")).toBe("2001:db8:0:0::/64");
    expect(normalizeIp("fe80::1%eth0")).toBe("fe80:0:0:0::/64");
    expect(normalizeIp("  203.0.113.7  ")).toBe("203.0.113.7");
  });

  it("BUCKETS IPv6 BY /64 — otherwise every limit here is IPv4-only", () => {
    // A residential IPv6 allocation is 2^64 addresses. Per-exact-address caps
    // would be bypassed by picking a new address per request, so the prefix a
    // single subscriber is delegated is the honest unit.
    const a = normalizeIp("2001:db8:1234:5678:aaaa:bbbb:cccc:dddd");
    const b = normalizeIp("2001:db8:1234:5678:9999:8888:7777:6666");
    expect(a).toBe(b);
    // A genuinely different customer prefix stays a different bucket.
    expect(normalizeIp("2001:db8:1234:9999::1")).not.toBe(a);
  });
});

describe("IpGuard", () => {
  it("rate-limits seeds with a refilling token bucket", () => {
    let now = 0;
    const guard = new IpGuard({ seedPerMin: 3, maxDocsPerIp: 0, maxConnsPerIp: 0, trustProxyHops: 0 }, () => now);
    for (let i = 0; i < 3; i++) expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: true });
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: false, reason: "ip-seed-limit" });
    // Another address is unaffected — the bucket is per-IP, not global.
    expect(guard.allowSeed("5.6.7.8")).toEqual({ ok: true });
    // Twenty seconds refills exactly one token at 3/min.
    now += 20_000;
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: true });
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: false, reason: "ip-seed-limit" });
  });

  it("caps LIVE rooms per address and releases the slot on eviction", () => {
    const guard = new IpGuard({ seedPerMin: 0, maxDocsPerIp: 2, maxConnsPerIp: 0, trustProxyHops: 0 });
    guard.noteRoom("1.2.3.4", "d1");
    guard.noteRoom("1.2.3.4", "d2");
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: false, reason: "ip-doc-limit" });
    // It is a CONCURRENCY cap, not a quota: closing one frees the slot.
    guard.releaseRoom("d1");
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: true });
  });

  it("release is idempotent — a double-decrement would shrink the budget forever", () => {
    // Eviction paths overlap (an idle kick empties a room the empty sweep may
    // also visit). A leaked decrement here surfaces later as "this network
    // can't create documents any more", which is miserable to diagnose.
    const guard = new IpGuard({ seedPerMin: 0, maxDocsPerIp: 1, maxConnsPerIp: 0, trustProxyHops: 0 });
    guard.noteRoom("1.2.3.4", "d1");
    guard.releaseRoom("d1");
    guard.releaseRoom("d1");
    guard.releaseRoom("d1");
    guard.noteRoom("1.2.3.4", "d2");
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: false, reason: "ip-doc-limit" });
  });

  it("does not spend a seed token when the doc cap already refused", () => {
    // Otherwise an address parked at its room limit also burns its rate
    // budget, and the reason reported would flip to the wrong limit.
    const guard = new IpGuard({ seedPerMin: 2, maxDocsPerIp: 1, maxConnsPerIp: 0, trustProxyHops: 0 });
    guard.noteRoom("1.2.3.4", "d1");
    for (let i = 0; i < 5; i++) expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: false, reason: "ip-doc-limit" });
    guard.releaseRoom("d1");
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: true });
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: true });
    expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: false, reason: "ip-seed-limit" });
  });

  it("caps concurrent connections and counts distinct address buckets", () => {
    const guard = new IpGuard({ seedPerMin: 0, maxDocsPerIp: 0, maxConnsPerIp: 2, trustProxyHops: 0 });
    expect(guard.openConn("1.2.3.4", "c1")).toEqual({ ok: true });
    expect(guard.openConn("1.2.3.4", "c2")).toEqual({ ok: true });
    expect(guard.openConn("1.2.3.4", "c3")).toEqual({ ok: false, reason: "ip-conn-limit" });
    expect(guard.openConn("5.6.7.8", "c4")).toEqual({ ok: true });
    expect(guard.distinctIps()).toBe(2);
    guard.closeConn("c1");
    guard.closeConn("c1"); // idempotent
    expect(guard.openConn("1.2.3.4", "c5")).toEqual({ ok: true });
    expect(guard.openConn("1.2.3.4", "c6")).toEqual({ ok: false, reason: "ip-conn-limit" });
  });

  it("0 disables a limit, and an unknown address is never guessed at", () => {
    const guard = new IpGuard({ seedPerMin: 0, maxDocsPerIp: 0, maxConnsPerIp: 0, trustProxyHops: 0 });
    for (let i = 0; i < 100; i++) expect(guard.allowSeed("1.2.3.4")).toEqual({ ok: true });
    for (let i = 0; i < 100; i++) expect(guard.openConn("1.2.3.4", `c${i}`)).toEqual({ ok: true });
    // An empty key (no transport, a unix socket, a test fake) is exempt
    // rather than lumped into one shared bucket.
    const strict = new IpGuard({ seedPerMin: 1, maxDocsPerIp: 1, maxConnsPerIp: 1, trustProxyHops: 0 });
    for (let i = 0; i < 10; i++) expect(strict.allowSeed("")).toEqual({ ok: true });
  });
});

describe("configuration parsing", () => {
  it("falls back to the default for anything unusable rather than throwing", () => {
    // A typo in an env var must not stop a server from booting.
    const saved = { ...process.env };
    try {
      process.env.WW_TEST_LIMIT = "10min";
      expect(envInt("WW_TEST_LIMIT", 42)).toBe(42);
      process.env.WW_TEST_LIMIT = "-5";
      expect(envInt("WW_TEST_LIMIT", 42)).toBe(42);
      process.env.WW_TEST_LIMIT = "";
      expect(envInt("WW_TEST_LIMIT", 42)).toBe(42);
      process.env.WW_TEST_LIMIT = "900";
      expect(envInt("WW_TEST_LIMIT", 42)).toBe(900);
      delete process.env.WW_TEST_LIMIT;
      expect(envInt("WW_TEST_LIMIT", 42)).toBe(42);
    } finally {
      process.env = saved;
    }
  });

  it("clamps a warning lead time that would fire at creation", () => {
    const l = normalizeLimits({ lifecycle: { idleTimeoutMs: 60_000, idleWarningMs: 600_000 } });
    expect(l.lifecycle.idleWarningMs).toBe(60_000);
  });

  it("ships the documented defaults", () => {
    // The compose/README env tables are transcribed from these numbers, so a
    // silent change to one should break something.
    expect(DEFAULT_LIMITS.lifecycle.emptyRoomTtlMs).toBe(60_000);
    expect(DEFAULT_LIMITS.lifecycle.idleTimeoutMs).toBe(600_000);
    expect(DEFAULT_LIMITS.lifecycle.roomMaxLifetimeMs).toBe(43_200_000);
    expect(DEFAULT_LIMITS.ip.seedPerMin).toBe(10);
    expect(DEFAULT_LIMITS.ip.maxDocsPerIp).toBe(25);
    expect(DEFAULT_LIMITS.ip.maxConnsPerIp).toBe(50);
    expect(DEFAULT_LIMITS.ip.trustProxyHops).toBe(0);
  });
});

describe("seed path enforcement", () => {
  const limits = normalizeLimits({ ip: { seedPerMin: 2, maxDocsPerIp: 1 } });

  it("refuses an over-rate seed with 429, distinctly from the 409 first-wins case", () => {
    const guard = new IpGuard(limits.ip);
    const obs = new MetricsObservability({ level: "silent", out: () => {} });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, obs, limits, guard);
    const docx = Buffer.from(blankDocxBytes()).toString("base64");
    let n = 0;
    const seed = (ip: string) =>
      handleSeedRequest(hub, { method: "POST", body: { docx } }, { creatorIp: ip, mintDocId: () => `d${n++}` });

    expect(seed("1.2.3.4").status).toBe(201);
    // Second attempt hits the LIVE-DOC cap (1), not the rate cap.
    const capped = seed("1.2.3.4");
    expect(capped.status).toBe(429);
    expect(capped.body).toEqual({ error: "ip-doc-limit" });

    // A different address is unaffected, and its own 409 path still works:
    // first-wins carries the incumbent epoch, throttles carry no epoch.
    expect(seed("5.6.7.8").status).toBe(201);
    const clash = handleSeedRequest(hub, { method: "PUT", docId: "d0", body: { docx } }, { creatorIp: "9.9.9.9" });
    expect(clash.status).toBe(409);
    expect(clash.body.error).toBe("exists");
    expect(clash.body.genesisId).toBeTruthy();

    expect(obs.snapshot().ipLimitsByReason).toEqual({ "ip-doc-limit": 1 });
  });

  it("evicting a room returns the creator's slot", () => {
    let now = 0;
    const guard = new IpGuard(limits.ip, () => now);
    const hub = new CollabHub(null, undefined, undefined, () => now, undefined, undefined, limits, guard);
    const docx = Buffer.from(blankDocxBytes()).toString("base64");
    let n = 0;
    const seed = () =>
      handleSeedRequest(hub, { method: "POST", body: { docx } }, { creatorIp: "1.2.3.4", mintDocId: () => `d${n++}` });

    expect(seed().status).toBe(201);
    expect(seed().status).toBe(429); // at the cap of 1
    // A seeded room nobody joins evicts on the empty-room clock; the slot
    // must come back with it, or the cap becomes a permanent quota.
    now += DEFAULT_LIMITS.lifecycle.emptyRoomTtlMs;
    expect(hub.sweepRooms()).toEqual(["d0"]);
    expect(seed().status).toBe(201);
  });
});

describe("transport connection cap", () => {
  /** Minimal ws-shaped fakes; the adapter is structural by design. */
  class FakeSocket implements WsSocket {
    sent: string[] = [];
    closed = false;
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

  it("refuses over-cap sockets AT OPEN, with a reason, and hangs up", () => {
    const guard = new IpGuard({ seedPerMin: 0, maxDocsPerIp: 0, maxConnsPerIp: 2, trustProxyHops: 0 });
    const obs = new MetricsObservability({ level: "silent", out: () => {} });
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, obs);
    const wss = new FakeWss();
    attachWebSocketServer(wss, hub, obs, { guard, trustProxyHops: 0 });

    const s1 = wss.connect("1.2.3.4");
    const s2 = wss.connect("1.2.3.4");
    const s3 = wss.connect("1.2.3.4");
    expect(s1.closed).toBe(false);
    expect(s2.closed).toBe(false);
    // The third is told why and hung up on — a refusal frame the peer can
    // ignore is not a limit.
    expect(s3.closed).toBe(true);
    expect(s3.frames()).toEqual([{ t: "refused", reason: "ip-conn-limit" }]);

    // A different address is unaffected…
    expect(wss.connect("5.6.7.8").closed).toBe(false);
    // …and closing one frees the slot.
    s1.handlers.close?.();
    expect(wss.connect("1.2.3.4").closed).toBe(false);

    const snap = obs.snapshot();
    expect(snap.ipLimitsByReason).toEqual({ "ip-conn-limit": 1 });
    // A socket hung up at the door was never a connection this server
    // served: the live gauge counts only admitted ones.
    expect(snap.counters.connectionsOpened).toBe(4);
    expect(snap.gauges.liveConnections).toBe(3);
  });

  it("reads the proxy header only when hops are trusted", () => {
    // Same two sockets, same forged header, opposite outcomes — this is the
    // trust-proxy switch doing its job end to end at the transport.
    const makeStack = (hops: number) => {
      const guard = new IpGuard({ seedPerMin: 0, maxDocsPerIp: 0, maxConnsPerIp: 1, trustProxyHops: hops });
      const hub = new CollabHub(null);
      const wss = new FakeWss();
      attachWebSocketServer(wss, hub, undefined, { guard, trustProxyHops: hops });
      return { wss, guard };
    };

    // hops=0: the header is ignored, both sockets share the peer's bucket,
    // the second is refused.
    const untrusting = makeStack(0);
    const a1 = new FakeSocket();
    const a2 = new FakeSocket();
    const forged = (n: number): WsRequest => ({
      socket: { remoteAddress: "203.0.113.7" },
      headers: { "x-forwarded-for": `10.0.0.${n}` },
    });
    (untrusting.wss as unknown as { cb?: (s: WsSocket, r?: WsRequest) => void }).cb?.(a1, forged(1));
    (untrusting.wss as unknown as { cb?: (s: WsSocket, r?: WsRequest) => void }).cb?.(a2, forged(2));
    expect(a2.closed).toBe(true);
    expect(untrusting.guard.distinctIps()).toBe(1);
  });
});
