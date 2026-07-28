import { type IpLimits, DEFAULT_IP_LIMITS } from "./limits.js";

/**
 * Per-IP abuse protection for an unauthenticated, zero-custody host.
 *
 * WHY THIS EXISTS AS ITS OWN OBJECT: the three limits are enforced at three
 * different layers — connection count at the transport when a socket opens,
 * seed rate and live-room count at the HTTP seed path — and splitting the
 * POLICY across those layers is how limits drift apart from each other. All
 * the state and every decision lives here; the layers just ask.
 *
 * WHAT AN IP IS WORTH HERE: not much, and the code should not pretend
 * otherwise. There are no accounts in this product, so the address is the
 * only actor identity available. It over-groups (an office is one address —
 * see the NAT caveat on `maxDocsPerIp`) and under-groups (a botnet is
 * thousands). These limits stop one script from consuming the host; they are
 * not a fairness mechanism and not a security boundary against a determined
 * distributed attacker.
 *
 * ZERO CUSTODY: an address is personal data, so it never leaves this object.
 * Nothing here logs one, returns one, or puts one in a snapshot — callers get
 * decisions and counts. The only strings that escape are the fixed refusal
 * vocabulary below.
 */

/** Refusal reasons this guard produces — fixed vocabulary, safe to log and
 * tally by. Distinct per limit so a spike identifies WHICH one is firing. */
export type IpLimitReason = "ip-seed-limit" | "ip-doc-limit" | "ip-conn-limit";

export type IpDecision = { ok: true } | { ok: false; reason: IpLimitReason };

const ALLOW: IpDecision = { ok: true };

/**
 * How long an idle token bucket is kept before the sweep drops it. Buckets
 * are tiny, but a server that allocates one per address and never frees them
 * has traded a flood limit for a memory leak — which is the same denial of
 * service by a slower route.
 */
const BUCKET_IDLE_MS = 10 * 60_000;

export class IpGuard {
  /** Seed token buckets, keyed by normalized address. */
  private seedBuckets = new Map<string, { tokens: number; last: number }>();
  /** Live rooms per address (creator attribution), and the reverse map so
   * eviction can release without knowing who created what. */
  private docsByIp = new Map<string, number>();
  private ipByDoc = new Map<string, string>();
  /** Open sockets per address, and the reverse map keyed by connection id. */
  private connsByIp = new Map<string, number>();
  private ipByConn = new Map<string, string>();

  constructor(
    private limits: IpLimits = DEFAULT_IP_LIMITS,
    /** Injected clock (epoch ms), matching the hub's — tests drive both from
     * one variable and never wait on wall time. */
    private now: () => number = () => Date.now(),
  ) {}

  /**
   * May this address create another room? Checks the live-room cap first and
   * the token bucket second, and DOES NOT SPEND a token when the cap already
   * refuses — otherwise an address parked at its room limit would also burn
   * its seed budget, and the reason reported for the next attempt would flip
   * to the wrong one.
   *
   * A token is spent on every allowed call, so the caller must only ask when
   * it is actually about to seed.
   */
  allowSeed(ip: string): IpDecision {
    if (!ip) return ALLOW; // unknown address (test fake, unix socket): don't guess
    if (this.limits.maxDocsPerIp > 0 && (this.docsByIp.get(ip) ?? 0) >= this.limits.maxDocsPerIp) {
      return { ok: false, reason: "ip-doc-limit" };
    }
    if (this.limits.seedPerMin > 0 && !this.spendSeedToken(ip)) {
      return { ok: false, reason: "ip-seed-limit" };
    }
    return ALLOW;
  }

  /** Continuous-refill token bucket: `seedPerMin` per minute, burst of the
   * same size. Returns false when empty (nothing is spent). */
  private spendSeedToken(ip: string): boolean {
    const now = this.now();
    const burst = this.limits.seedPerMin;
    let b = this.seedBuckets.get(ip);
    if (!b) {
      b = { tokens: burst, last: now };
      this.seedBuckets.set(ip, b);
    }
    b.tokens = Math.min(burst, b.tokens + ((now - b.last) / 60_000) * this.limits.seedPerMin);
    b.last = now;
    if (b.tokens < 1) return false;
    b.tokens -= 1;
    return true;
  }

  /** Record that `ip` created the room `docId` (called after a seed wins). */
  noteRoom(ip: string, docId: string): void {
    if (!ip) return;
    // A re-seed of a docId already attributed elsewhere would double-count;
    // release first so the mapping stays one-to-one.
    this.releaseRoom(docId);
    this.ipByDoc.set(docId, ip);
    this.docsByIp.set(ip, (this.docsByIp.get(ip) ?? 0) + 1);
  }

  /** Release a room's slot (eviction, by any reason). Idempotent — eviction
   * paths overlap (an idle kick empties a room that the empty sweep may also
   * visit), and a decrement that ran twice would let an address exceed its
   * cap forever. */
  releaseRoom(docId: string): void {
    const ip = this.ipByDoc.get(docId);
    if (ip === undefined) return;
    this.ipByDoc.delete(docId);
    const n = (this.docsByIp.get(ip) ?? 0) - 1;
    if (n > 0) this.docsByIp.set(ip, n);
    else this.docsByIp.delete(ip);
  }

  /**
   * May this address open another socket, and if so, register it.
   *
   * Called when the SOCKET OPENS, not at hello: the resource being spent is
   * the socket itself, so a check at hello would leave a connect-and-say-
   * nothing flood completely unbounded. Registration happens here too, so
   * callers cannot allow-then-forget-to-record.
   */
  openConn(ip: string, connId: string): IpDecision {
    if (!ip) return ALLOW;
    if (this.limits.maxConnsPerIp > 0 && (this.connsByIp.get(ip) ?? 0) >= this.limits.maxConnsPerIp) {
      return { ok: false, reason: "ip-conn-limit" };
    }
    this.ipByConn.set(connId, ip);
    this.connsByIp.set(ip, (this.connsByIp.get(ip) ?? 0) + 1);
    return ALLOW;
  }

  /** Release a socket's slot. Idempotent, for the same reason releaseRoom is:
   * a leaked increment here permanently shrinks an address's budget, and the
   * symptom (a user who "can't connect from this network any more") is
   * miserable to diagnose. */
  closeConn(connId: string): void {
    const ip = this.ipByConn.get(connId);
    if (ip === undefined) return;
    this.ipByConn.delete(connId);
    const n = (this.connsByIp.get(ip) ?? 0) - 1;
    if (n > 0) this.connsByIp.set(ip, n);
    else this.connsByIp.delete(ip);
  }

  /**
   * Number of DISTINCT addresses currently holding sockets.
   *
   * The diagnostic for the proxy misconfiguration: a deployment running
   * behind Caddy with WW_TRUST_PROXY unset attributes every connection to the
   * proxy, so this reads 1 while `ip-conn-limit` refusals climb. That pair of
   * signals says "trust-proxy is wrong" and nothing else does — which is why
   * a count nobody would otherwise export is worth exporting. Safe: a
   * cardinality, never an address.
   */
  distinctIps(): number {
    return this.connsByIp.size;
  }

  /** Drop token buckets that have been idle long enough to be full again;
   * called from the same periodic sweep as the room timers. Bounded memory is
   * part of the protection, not housekeeping. */
  sweep(): void {
    const now = this.now();
    for (const [ip, b] of [...this.seedBuckets]) {
      if (now - b.last >= BUCKET_IDLE_MS) this.seedBuckets.delete(ip);
    }
  }
}

/** A guard that permits everything — the default wherever one is optional, so
 * an un-wired hub behaves exactly as it did before this arc. */
export const NO_OP_IP_GUARD = new IpGuard({
  seedPerMin: 0,
  maxDocsPerIp: 0,
  maxConnsPerIp: 0,
  trustProxyHops: 0,
});
