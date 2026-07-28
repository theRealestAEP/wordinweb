import { MAX_IMAGE_BYTES } from "@wordinweb/collab/server";
/**
 * Server lifecycle + abuse-protection configuration.
 *
 * ONE PLACE. Every tunable the host has, its default, and the env var that
 * overrides it live here — not scattered as magic numbers across hub.ts,
 * ws.ts and cli.ts. An operator reading this file learns the entire surface,
 * and the compose/README env tables are transcribed from it.
 *
 * THE CONVENTIONS, applied uniformly:
 *
 *  - **0 disables** a limit. Every knob below is a threshold, and "no
 *    threshold" is a legitimate configuration (a private deployment that
 *    wants rooms to live until the process dies, a load test that must not
 *    be throttled). Encoding that as 0 rather than a sentinel like -1 or the
 *    absence of the var means the disable is visible in `docker inspect`.
 *  - **A bad value falls back to the default, never throws.** Same rule
 *    envLogLevel already follows: a typo in an env var must not stop a
 *    server from booting. A server that refuses to start because someone
 *    wrote `WW_IDLE_TIMEOUT_MS=10min` is a server that is down.
 *  - **Everything is injectable.** The hub takes a resolved `HubLimits`, so
 *    tests set exact numbers and drive the existing injected clock instead
 *    of waiting on wall time. Nothing here reads `process.env` except
 *    {@link envLimits}, which is called once at startup.
 */

/** Timers governing a room's life. All epoch-ms durations; 0 disables. */
export interface LifecycleLimits {
  /**
   * Last participant disconnects → the room (document, log, dedup state,
   * everything) is deleted after this long. Long enough to survive a page
   * refresh or a flaky reconnect, short enough that "everyone left" promptly
   * means "the server holds nothing". This is the zero-custody promise's
   * actual enforcement point.
   */
  emptyRoomTtlMs: number;
  /**
   * No QUALIFYING ACTIVITY from anyone for this long → every participant is
   * kicked (`idle-timeout`) and the room evicts, even though sockets are
   * still open.
   *
   * Qualifying activity is deliberately NOT presence — see
   * {@link LIFECYCLE_NOTES.activity}. A room full of forgotten tabs is
   * exactly what this reclaims.
   */
  idleTimeoutMs: number;
  /**
   * How long before idle expiry the `session-warning` goes out, so the UI can
   * run a countdown and a returning human can save the session by typing.
   * Must be < idleTimeoutMs to mean anything (a larger value would warn at
   * creation); clamped rather than rejected.
   */
  idleWarningMs: number;
  /**
   * ABSOLUTE cap on a room's age, measured from creation and reset by
   * nothing. This is the anti-keepalive-bot limit: the idle timeout can be
   * held off indefinitely by a script that submits one edit every nine
   * minutes, and this is the wall that script hits.
   *
   * Ending the session does not end the DOCUMENT: every participant holds a
   * bundle, and re-seeding produces a NEW EPOCH (fresh genesisId, fresh
   * keys in E2EE rooms) that the doc-15 lineage machinery already handles —
   * a rejoining holder whose head appears in the new seed's lineage
   * fast-forwards silently. So the cap costs a re-seed, never data.
   */
  roomMaxLifetimeMs: number;
  /** Warning lead time for {@link roomMaxLifetimeMs}. Unlike the idle
   * warning this one can never be cancelled — the deadline is fixed at
   * creation and no activity moves it. */
  lifetimeWarningMs: number;
}

/**
 * Per-IP abuse limits. These are BLUNT INSTRUMENTS and are meant to be: a
 * zero-custody server has no accounts, so the IP is the only actor identity
 * available, and it is a poor one (see the NAT caveat on
 * {@link maxDocsPerIp}). They exist to stop a script from consuming the whole
 * host, not to enforce fair use between humans.
 */
export interface IpLimits {
  /**
   * Token bucket on room CREATION (POST/PUT /docs), refilled continuously at
   * this rate per minute with a burst of the same size. Seeding is the
   * expensive unauthenticated operation — it parses a docx and allocates a
   * room — so it is the one that needs a rate, not just a cap.
   */
  seedPerMin: number;
  /**
   * Maximum LIVE rooms created by one IP. Released as rooms evict, so it is
   * a concurrency cap and not a quota.
   *
   * NAT CAVEAT, and the reason this number is deliberately generous: an
   * office, a school, a conference wifi, and a mobile carrier's CGNAT pool
   * all present as ONE address. 25 is chosen to be a number no plausible
   * shared-egress group of humans reaches while still bounding a script, and
   * it is configurable precisely because that guess is wrong for somebody.
   * If you are seeing `ip-doc-limit` refusals from real users, this is the
   * knob — raising it costs RAM proportional to live rooms, nothing more.
   */
  maxDocsPerIp: number;
  /**
   * Maximum concurrent WebSocket connections from one IP. Enforced when the
   * socket OPENS, before any protocol message: the resource being spent (a
   * socket, an fd, a slot in the event loop) is consumed at open, so a check
   * at hello would leave a pre-hello flood entirely unbounded.
   *
   * The same NAT caveat applies, one order of magnitude looser: shared
   * egress multiplied by several tabs each.
   */
  maxConnsPerIp: number;
  /**
   * NUMBER OF TRUSTED PROXY HOPS in front of this server — not a boolean.
   * 0 (the default) means trust nothing: the peer address of the socket is
   * the client, and `X-Forwarded-For` is never read. 1 means one reverse
   * proxy we control (the deploy's Caddy). 2 would be Caddy behind a CDN.
   *
   * WHY A COUNT: X-Forwarded-For is `client, proxy1, proxy2, …` and each
   * hop APPENDS the address it observed. A client that sends its own
   * `X-Forwarded-For: 9.9.9.9` therefore reaches us as `9.9.9.9, <real
   * address Caddy saw>` — the LEFT-most entry is attacker-controlled and the
   * RIGHT-most is the only one our own infrastructure vouched for. Trusting
   * N hops means taking the entry N from the right. Setting this HIGHER than
   * the real hop count is what re-opens spoofing: you start reading entries
   * no trusted proxy wrote.
   *
   * See {@link clientIp} for the extraction and both failure modes.
   */
  trustProxyHops: number;
}

/**
 * SURGE VALVES: the limits that bound what one hostile (or merely broken)
 * participant can make the server spend, in the dimensions the per-IP caps do
 * not cover.
 *
 * The per-IP caps answer "how many things may one address start". These
 * answer "how much may an already-admitted session consume" — a different
 * question, and the one a MIXED legitimate/hostile load actually turns on.
 * Each of these closes a hole where every existing limit is satisfied and the
 * server still dies.
 */
export interface SurgeLimits {
  /**
   * Largest inbound WebSocket frame accepted, in bytes.
   *
   * THE FIRST WALL, and it has to be first: `ws` defaults maxPayload to
   * 100 MiB, so without this a hostile client makes the server BUFFER AND
   * PARSE a hundred-megabyte frame before any hub-level cap — including the
   * 256 KB envelope cap — ever sees it. A cap enforced after parsing is a cap
   * enforced after the damage.
   *
   * The default sits comfortably above the 256 KB ciphertext cap plus JSON and
   * base64 overhead. `ws` closes a violating connection with status 1009
   * (message too big), which reaches our ordinary close path.
   */
  wsMaxPayloadBytes: number;
  /**
   * Ceiling on a single socket's OUTBOUND buffer, in bytes.
   *
   * The slow-consumer hole: a client that opens a socket and never reads
   * still receives every broadcast, and Node buffers all of it in the
   * server's memory indefinitely. Nothing else here notices — the client is
   * within its connection limit, its rate limit, and its frame size; it is
   * simply not draining. Past this ceiling the socket is terminated
   * (`backpressure`), which costs that participant a reconnect and costs the
   * room nothing: their bundle is intact and rejoining is cheap.
   */
  wsMaxBufferedBytes: number;
  /**
   * Ceiling on ONE room's retained envelope log, in bytes.
   *
   * THE MULTIPLICATIVE HOLE, and the least obvious one: 256 KB envelopes at
   * 400/s from 10 connections is inside every limit above and still grows a
   * single room's RAM at roughly a gigabyte per second. The log only prunes
   * at a quiescent checkpoint, and a room under sustained hostile submission
   * never goes quiescent — so the one mechanism that would bound it is
   * exactly the one the attack prevents from running.
   *
   * NOBODY LEGITIMATE COMES NEAR THIS. A sealed keystroke envelope is roughly
   * 500 bytes, so the default is on the order of a hundred thousand real
   * edits in a single epoch — far past where a checkpoint would have pruned.
   * Crossing it means the room is not being used as a document.
   */
  roomLogMaxBytes: number;
  /**
   * Global ceiling on live rooms; the seed path refuses `server-full` (503).
   *
   * Per-IP caps bound one address. They do not bound a thousand addresses,
   * and a botnet is exactly a thousand addresses each behaving impeccably.
   * This is the fail-fast valve: briefly refusing NEW sessions is a much
   * better failure than an OOM that takes every EXISTING session with it.
   */
  maxRooms: number;
  /**
   * Global ceiling on open sockets, enforced at accept for the same reason
   * the per-IP connection cap is: the resource is spent when the socket
   * opens. Existing sessions are untouched.
   */
  maxConns: number;
}

/**
 * Media relay limits (doc 16 §8, RAM tier). Hardcoded until the surge arc;
 * now configurable because the right blob size is a DEPLOYMENT decision — a
 * public demo wants a few megabytes, local testing wants room to paste a
 * camera image without thinking about it.
 */
export interface MediaLimits {
  /** Largest single blob the relay accepts (HTTP 413 above it, and the
   * response carries this number so a UI can say "images must be under 5 MB"
   * rather than "upload failed"). */
  maxBlobBytes: number;
  /** Total bytes one room's relay may hold (HTTP 507 above it). */
  maxRoomBytes: number;
  /** T_STAGE: how long an uploaded-but-unreferenced blob survives. Uploads
   * precede the intent that references them, so a blob that never gets
   * claimed is a cancelled paste and should not linger. */
  stageTtlMs: number;
  /** T_MEDIA: how long a promoted blob survives after its last download. The
   * locker is a CACHE, not a store — every blob is recoverable from
   * participants by re-supply, so expiry costs latency, never data. */
  ttlMs: number;
  /** How long a chosen holder has to answer a re-supply request before the
   * round is offered to someone else. */
  uploadDeadlineMs: number;
  /** PER-ROOM upload admission: sustained uploads per minute, and the burst
   * a room may spend at once.
   *
   * `maxRoomBytes` caps how much a room STORES; it does not cap how fast a
   * room is asked to store it. Those are different exposures: once a room is
   * full every further upload is refused with 507, and a refusal is still a
   * request the server accepted a body for. This is the rate half.
   *
   * BURST IS NOT DECORATION — it is sized for re-supply. When a joiner needs
   * a document's images, the hub asks a holder for each missing sha, so an
   * image-heavy document legitimately produces a burst of uploads with no
   * human pacing them at all. A limiter with no burst allowance would throttle
   * exactly that, and the symptom would be images that never arrive for the
   * person who just joined — the failure this relay exists to prevent. */
  roomUploadsPerMin: number;
  roomUploadBurst: number;
}

export interface HubLimits {
  lifecycle: LifecycleLimits;
  ip: IpLimits;
  surge: SurgeLimits;
  media: MediaLimits;
}

/**
 * Notes that belong with the config rather than buried at a call site.
 * Exported as documentation, referenced from the code that implements them.
 */
export const LIFECYCLE_NOTES = {
  /**
   * WHAT COUNTS AS ACTIVITY for the idle timeout: accepted submits (both
   * plaintext and encrypted), owner admin actions, media transfers, and an
   * accepted hello (someone joined).
   *
   * PRESENCE DOES NOT COUNT. This is the decision the whole timeout hangs
   * on: a forgotten open tab emits presence whenever the mouse moves or the
   * window regains focus, so counting it would mean any abandoned laptop
   * keeps a room alive forever and the timeout is dead letter.
   *
   * A HELLO DOES count, because the alternative is worse: a room idle for
   * nine minutes that someone new joins would otherwise hand the newcomer a
   * 60-second countdown and evict the session under them before they typed a
   * character. Keepalive-by-rejoin is bounded by the absolute lifetime cap
   * and the per-IP connection limit, so admitting joins here does not
   * reopen the hole the timeout exists to close.
   */
  activity: "submits + admin + media + hello; never presence",
} as const;

export const DEFAULT_LIFECYCLE: LifecycleLimits = {
  emptyRoomTtlMs: 60_000,
  idleTimeoutMs: 10 * 60_000,
  idleWarningMs: 60_000,
  roomMaxLifetimeMs: 12 * 60 * 60_000,
  lifetimeWarningMs: 5 * 60_000,
};

export const DEFAULT_IP_LIMITS: IpLimits = {
  seedPerMin: 10,
  maxDocsPerIp: 25,
  maxConnsPerIp: 50,
  trustProxyHops: 0,
};

export const DEFAULT_SURGE_LIMITS: SurgeLimits = {
  wsMaxPayloadBytes: 512 * 1024,
  wsMaxBufferedBytes: 4 * 1024 * 1024,
  roomLogMaxBytes: 64 * 1024 * 1024,
  maxRooms: 2000,
  maxConns: 5000,
};

export const DEFAULT_MEDIA_LIMITS: MediaLimits = {
  // 5 MB for a public deployment (the operator-facing number); local stacks
  // raise it via WW_MEDIA_MAX_BLOB_BYTES.
  maxBlobBytes: 5 * 1024 * 1024,
  maxRoomBytes: 100 * 1024 * 1024,
  stageTtlMs: 60_000,
  ttlMs: 5 * 60_000,
  uploadDeadlineMs: 15_000,
  // 60/min sustained, burst 64. Sized against RE-SUPPLY, not against typing:
  // a joiner needing a document's images makes the hub ask holders for each
  // missing sha at once, so the burst has to cover a whole document's image
  // count or those uploads are refused, the hub rotates holders on its
  // uploadDeadlineMs, and the joiner sees "image unavailable" for pictures
  // that exist. 64 covers realistic documents; note the room byte cap
  // (100 MB) is the binding limit well before then for large images.
  //
  // A document with MORE images than the burst still resolves — the client
  // re-requests holes it can still see, so the tail arrives as tokens refill
  // rather than being lost. It is slower, which is the intended shape of a
  // limit. Raise the burst before the rate if that tail is being felt.
  roomUploadsPerMin: 60,
  roomUploadBurst: 64,
};

export const DEFAULT_LIMITS: HubLimits = {
  lifecycle: DEFAULT_LIFECYCLE,
  ip: DEFAULT_IP_LIMITS,
  surge: DEFAULT_SURGE_LIMITS,
  media: DEFAULT_MEDIA_LIMITS,
};

/**
 * Read a non-negative integer from the environment, falling back to `def` for
 * anything unusable: unset, empty, NaN, negative, infinite, or a value with
 * units someone hoped we would parse (`10min`). Never throws.
 */
export function envInt(name: string, def: number): number {
  const raw = typeof process !== "undefined" ? process.env[name] : undefined;
  if (raw === undefined || raw.trim() === "") return def;
  const n = Number(raw);
  if (!Number.isFinite(n) || n < 0) return def;
  return Math.floor(n);
}

/**
 * Resolve the whole limit set from the environment. Called ONCE at startup
 * (cli.ts) — everything downstream takes the resolved object, so no code path
 * re-reads `process.env` and a test can hand over any configuration it likes.
 */
export function envLimits(): HubLimits {
  return normalizeLimits({
    lifecycle: {
      emptyRoomTtlMs: envInt("WW_EMPTY_ROOM_TTL_MS", DEFAULT_LIFECYCLE.emptyRoomTtlMs),
      idleTimeoutMs: envInt("WW_IDLE_TIMEOUT_MS", DEFAULT_LIFECYCLE.idleTimeoutMs),
      idleWarningMs: envInt("WW_IDLE_WARNING_MS", DEFAULT_LIFECYCLE.idleWarningMs),
      roomMaxLifetimeMs: envInt("WW_ROOM_MAX_LIFETIME_MS", DEFAULT_LIFECYCLE.roomMaxLifetimeMs),
      lifetimeWarningMs: envInt("WW_LIFETIME_WARNING_MS", DEFAULT_LIFECYCLE.lifetimeWarningMs),
    },
    ip: {
      seedPerMin: envInt("WW_IP_SEED_PER_MIN", DEFAULT_IP_LIMITS.seedPerMin),
      maxDocsPerIp: envInt("WW_IP_MAX_DOCS", DEFAULT_IP_LIMITS.maxDocsPerIp),
      maxConnsPerIp: envInt("WW_IP_MAX_CONNS", DEFAULT_IP_LIMITS.maxConnsPerIp),
      trustProxyHops: envInt("WW_TRUST_PROXY", DEFAULT_IP_LIMITS.trustProxyHops),
    },
    surge: {
      wsMaxPayloadBytes: envInt("WW_WS_MAX_PAYLOAD", DEFAULT_SURGE_LIMITS.wsMaxPayloadBytes),
      wsMaxBufferedBytes: envInt("WW_WS_MAX_BUFFERED", DEFAULT_SURGE_LIMITS.wsMaxBufferedBytes),
      roomLogMaxBytes: envInt("WW_ROOM_LOG_MAX_BYTES", DEFAULT_SURGE_LIMITS.roomLogMaxBytes),
      // NAMED `_GLOBAL` DELIBERATELY. `WW_MAX_CONNS` sits one character away
      // from the per-IP `WW_IP_MAX_CONNS` and means something entirely
      // different — an operator who confuses them either throttles the whole
      // server to one address's budget or opens the per-IP cap to the global
      // one. The same trap already sprang once in the other direction:
      // `WW_ROOM_CAP` (participants in one room) was documented as capping
      // the NUMBER OF ROOMS. Global ceilings therefore carry the suffix, and
      // the two families read differently at a glance.
      maxRooms: envInt("WW_MAX_ROOMS_GLOBAL", DEFAULT_SURGE_LIMITS.maxRooms),
      maxConns: envInt("WW_MAX_CONNS_GLOBAL", DEFAULT_SURGE_LIMITS.maxConns),
    },
    media: {
      // No clamp here on purpose: normalizeLimits() is the single funnel every
      // path passes through (env parsing, embedders, tests), and the ceiling
      // relationship is enforced there so it cannot be bypassed by anyone
      // constructing limits directly. Clamping in this function as well was
      // the first version of the fix and left exactly that hole open.
      maxBlobBytes: envInt("WW_MEDIA_MAX_BLOB_BYTES", DEFAULT_MEDIA_LIMITS.maxBlobBytes),
      maxRoomBytes: envInt("WW_MEDIA_MAX_ROOM_BYTES", DEFAULT_MEDIA_LIMITS.maxRoomBytes),
      stageTtlMs: envInt("WW_MEDIA_STAGE_TTL_MS", DEFAULT_MEDIA_LIMITS.stageTtlMs),
      ttlMs: envInt("WW_MEDIA_TTL_MS", DEFAULT_MEDIA_LIMITS.ttlMs),
      uploadDeadlineMs: envInt("WW_MEDIA_UPLOAD_DEADLINE_MS", DEFAULT_MEDIA_LIMITS.uploadDeadlineMs),
      roomUploadsPerMin: envInt("WW_MEDIA_ROOM_UPLOADS_PER_MIN", DEFAULT_MEDIA_LIMITS.roomUploadsPerMin),
      roomUploadBurst: envInt("WW_MEDIA_ROOM_UPLOAD_BURST", DEFAULT_MEDIA_LIMITS.roomUploadBurst),
    },
  });
}

/**
 * Headroom between the configured media blob cap and the SOCKET-LEVEL guard
 * that destroys an oversized request body mid-flight.
 *
 * The socket guard must sit strictly ABOVE the blob cap. If they cross, a
 * legal upload is killed at the transport before the hub can answer 413, and
 * the client sees a dropped connection instead of "that image is too big" —
 * a failure that is invisible until someone raises the cap and cannot be
 * diagnosed from either side. Deriving one from the other makes the ordering
 * structural instead of a pair of constants that agree today.
 */
export const MEDIA_WIRE_HEADROOM_BYTES = 2 * 1024 * 1024;

/** The socket-level destroy threshold for a media upload body. */
export function mediaWireCap(media: MediaLimits): number {
  return media.maxBlobBytes + MEDIA_WIRE_HEADROOM_BYTES;
}

/**
 * Fill in a partial configuration and repair the one combination that is
 * incoherent rather than merely unusual: a warning lead time at or beyond its
 * own deadline, which would fire the warning at creation and leave no
 * countdown. Clamped to the deadline rather than rejected — the operator
 * asked for "warn as early as possible", and refusing to boot over it would
 * be the worse answer.
 */
export function normalizeLimits(partial: {
  lifecycle?: Partial<LifecycleLimits>;
  ip?: Partial<IpLimits>;
  surge?: Partial<SurgeLimits>;
  media?: Partial<MediaLimits>;
} = {}): HubLimits {
  const lifecycle: LifecycleLimits = { ...DEFAULT_LIFECYCLE, ...partial.lifecycle };
  const ip: IpLimits = { ...DEFAULT_IP_LIMITS, ...partial.ip };
  const surge: SurgeLimits = { ...DEFAULT_SURGE_LIMITS, ...partial.surge };
  const media: MediaLimits = { ...DEFAULT_MEDIA_LIMITS, ...partial.media };
  if (lifecycle.idleTimeoutMs > 0 && lifecycle.idleWarningMs >= lifecycle.idleTimeoutMs) {
    lifecycle.idleWarningMs = lifecycle.idleTimeoutMs;
  }
  if (lifecycle.roomMaxLifetimeMs > 0 && lifecycle.lifetimeWarningMs >= lifecycle.roomMaxLifetimeMs) {
    lifecycle.lifetimeWarningMs = lifecycle.roomMaxLifetimeMs;
  }
  // THE MEDIA CAP CLAMP BELONGS HERE, not only in envLimits.
  //
  // A blob cap above MAX_IMAGE_BYTES opens the window this whole fix exists to
  // close: the relay accepts an upload and the resulting intent is then
  // rejected as "bad size" on every replica — no image, no error, nothing to
  // diagnose. Clamping only while parsing the ENVIRONMENT leaves the hole open
  // for every other way a configuration is built: a test, an embedder wiring
  // its own hub, anything calling normalizeLimits directly. This is the one
  // funnel both paths pass through (envLimits builds its object and hands it
  // straight here), so enforcing the relationship at this point makes it
  // structural rather than a property of one entry point.
  if (media.maxBlobBytes > MAX_IMAGE_BYTES) media.maxBlobBytes = MAX_IMAGE_BYTES;
  // A rate with no burst admits NOTHING: the bucket starts at `burst` tokens
  // and a request needs one, so burst 0 refuses the first upload of every
  // room and every image in the deployment silently stops arriving. Zero rate
  // means "disabled" (the convention the per-IP limits already use); a rate
  // that is on therefore needs at least one token to spend.
  if (media.roomUploadsPerMin > 0 && media.roomUploadBurst < 1) media.roomUploadBurst = 1;
  return { lifecycle, ip, surge, media };
}

/** The minimal request shape IP extraction needs, so this module depends on
 * no transport type: both `node:http`'s IncomingMessage and `ws`'s upgrade
 * request satisfy it structurally. */
export interface IpRequestLike {
  socket?: { remoteAddress?: string | null } | null;
  headers?: Record<string, string | string[] | undefined>;
}

/**
 * Determine the client address to attribute a request to.
 *
 * THE SECURITY BOUNDARY OF THIS ARC. Every per-IP limit is exactly as strong
 * as this function, and the intuitive implementation ("read X-Forwarded-For,
 * take the first entry — that's the original client") produces a limiter that
 * any client bypasses with one header while looking perfectly correct in
 * every test where nobody tries.
 *
 * The rule: `X-Forwarded-For` is a comma list that each hop APPENDS to. With
 * `hops` proxies we control, the entry `hops` from the RIGHT is the address
 * our own infrastructure observed; everything to the left of it was supplied
 * by someone we do not trust. With `hops === 0` the header is not read at
 * all and the socket's peer address is used.
 *
 * BOTH FAILURE MODES, because they are asymmetric and that asymmetry is why
 * the default is 0:
 *
 *  - **hops=0 while actually behind a proxy** (the safe direction): every
 *    request appears to come from the proxy's address, so the whole internet
 *    shares one bucket and the limits fire against legitimate users. Loud and
 *    self-diagnosing — `ip-conn-limit` refusals climbing while the deployment
 *    obviously has more than one user.
 *  - **hops≥1 while NOT behind a proxy** (the dangerous direction): the
 *    client supplies the header itself and nothing appends a real address, so
 *    the "trusted" entry is attacker-chosen and per-IP protection is silently
 *    OFF while appearing configured. Nothing looks wrong until someone
 *    abuses it.
 *
 * A misconfiguration in between (hops set higher than the entries present)
 * degrades toward OVER-limiting — we fall back to the left-most entry, then
 * to the peer address — never toward no limiting.
 *
 * The returned value is a normalized BUCKET KEY, not necessarily a literal
 * address: IPv6 is truncated to its /64 prefix, because a residential IPv6
 * allocation is 2^64 addresses and per-exact-address caps would otherwise be
 * bypassable by picking a fresh address per request, making every limit here
 * IPv4-only. It is never logged (see observability.ts's zero-custody rule).
 */
export function clientIp(req: IpRequestLike | undefined | null, hops: number): string {
  const peer = normalizeIp(req?.socket?.remoteAddress ?? "");
  if (hops <= 0) return peer;
  const raw = req?.headers?.["x-forwarded-for"];
  const header = Array.isArray(raw) ? raw.join(",") : raw;
  if (!header) return peer; // proxy not actually forwarding: over-limit, don't under-limit
  const parts = header
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  if (parts.length === 0) return peer;
  // `hops` from the right; if the chain is shorter than configured, the
  // left-most entry is the oldest one present.
  const idx = Math.max(0, parts.length - hops);
  return normalizeIp(parts[idx]) || peer;
}

/**
 * Normalize an address into a stable bucket key so one client cannot occupy
 * two buckets (and two clients cannot share one by accident):
 * lowercase, brackets and IPv6 zone ids stripped, IPv4-mapped IPv6
 * (`::ffff:1.2.3.4`) reduced to its IPv4 form, a trailing `:port` removed
 * from IPv4 literals, and IPv6 truncated to /64.
 */
export function normalizeIp(addr: string): string {
  let s = (addr ?? "").trim().toLowerCase();
  if (!s) return "";
  if (s.startsWith("[")) {
    const close = s.indexOf("]");
    if (close > 0) s = s.slice(1, close);
  }
  const zone = s.indexOf("%"); // link-local scope id
  if (zone >= 0) s = s.slice(0, zone);
  if (s.startsWith("::ffff:") && s.includes(".")) s = s.slice("::ffff:".length);
  if (!s.includes(":")) {
    return s; // IPv4 (or a hostname in a test fake) — used whole
  }
  // A bare `1.2.3.4:5678` would have been caught above; anything with more
  // than one colon is IPv6. Truncate to the routing prefix a single customer
  // is allocated, so churning through addresses inside it does not multiply
  // an attacker's budget.
  return ipv6Prefix64(s);
}

/** First four hextets of an IPv6 address, `::`-expanded — the /64 a single
 * subscriber is typically delegated. Malformed input returns the input, which
 * simply buckets it by itself. */
function ipv6Prefix64(addr: string): string {
  const dbl = addr.indexOf("::");
  let groups: string[];
  if (dbl >= 0) {
    const head = addr.slice(0, dbl).split(":").filter(Boolean);
    const tail = addr.slice(dbl + 2).split(":").filter(Boolean);
    const fill = Math.max(0, 8 - head.length - tail.length);
    groups = [...head, ...Array(fill).fill("0"), ...tail];
  } else {
    groups = addr.split(":");
  }
  if (groups.length < 4) return addr;
  return groups
    .slice(0, 4)
    .map((g) => (g === "" ? "0" : g.replace(/^0+(?=.)/, "")))
    .join(":") + "::/64";
}
