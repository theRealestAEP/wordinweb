/**
 * Server-side observability for the zero-custody collab host.
 *
 * ZERO-CUSTODY INVARIANT (non-negotiable): this module records SHAPES and
 * COUNTS, never content. Nothing here ever touches plaintext, ciphertext
 * bodies, keys, tokens, share-codes, owner tokens, docIds, or clientIds — the
 * only strings it stores or logs are the fixed protocol vocabulary of refusal/
 * kick reasons (`taken-over`, `already-open`, `unauthorized`, …), which carry
 * no session content. Everything is in-memory and ephemeral: a process restart
 * resets all metrics, which is correct — the server persists nothing, ever.
 *
 * The sink is injected into the hub (constructor, no-op default) so the hub's
 * control flow, timing, and wire output stay byte-identical — observability is
 * strictly passive. One-line JSON events go to STDERR (stdout carries the
 * startup lines scripts parse — this module must never write there).
 *
 * DEFAULT VERBOSITY FLIPPED (2026-07-27, the log-shipping arc). This module
 * used to document "OFF by default so normal runs stay quiet" as a rule, and
 * that rule was load-bearing while the only consumer was a human watching a
 * terminal. It no longer holds: a server whose default is silence cannot be
 * shipped to a collector, which is the entire point of structured logs. The
 * default is now `info` — lifecycle only, not the per-op firehose — and the
 * old behavior is one env var away (`WW_LOG_LEVEL=silent`). See envLogLevel
 * for the full migration, including how the pre-existing `WW_OBS` flag maps.
 */

/** Fields safe to log/expose: aggregate primitives only, never identifiers or
 * content. The type is intentionally narrow so a stray object can't smuggle a
 * payload into a log line. */
export type SafeFields = Record<string, string | number | boolean | undefined>;

/**
 * Log levels, ordered. A line is emitted when its level is at or above the
 * configured threshold; `silent` emits nothing at all.
 *
 * The split is by OPERATIONAL MEANING, not by how interesting the code
 * thought it was:
 *  - debug  per-op traffic (submit-ok, broadcast) — 400/s under load, useful
 *           when reconstructing an incident, ruinous as a default
 *  - info   lifecycle: sockets, rooms, joins, checkpoints, media
 *  - warn   refusals and kicks — someone was told no, and the reason matters
 *  - error  something threw; carries a serialized stack
 */
export type LogLevel = "silent" | "debug" | "info" | "warn" | "error";

const LEVEL_ORDER: Record<LogLevel, number> = { silent: 99, debug: 10, info: 20, warn: 30, error: 40 };

/**
 * An error flattened to safe primitives. Stacks are the point of this arc —
 * three bugs in one session survived because a `catch {}` ate the only
 * evidence — but they are also the one field that can carry arbitrary text
 * into a log, so it is depth-capped and the CAUSE CHAIN is walked explicitly
 * rather than by serializing an unknown object graph.
 */
export interface SerializedError {
  name: string;
  message: string;
  stack?: string;
  /** `cause` chain, flattened outward-in, capped (see serializeError). */
  causes?: string[];
}

/** Maximum characters of a stack retained. Long enough for the frames that
 * matter, short enough that a pathological stack cannot flood a collector. */
const STACK_CAP = 4000;
/** Maximum links of a `cause` chain followed. */
const CAUSE_CAP = 3;

/**
 * Flatten an unknown throw into safe primitives. Accepts `unknown` because
 * that is what `catch` gives you and non-Error throws are real (a string, a
 * DOMException, `undefined` from a rejected promise with no reason).
 */
export function serializeError(err: unknown): SerializedError {
  if (!(err instanceof Error)) {
    return { name: typeof err, message: String(err).slice(0, 500) };
  }
  const out: SerializedError = {
    name: err.name,
    message: err.message.slice(0, 500),
    stack: err.stack?.slice(0, STACK_CAP),
  };
  const causes: string[] = [];
  let cause: unknown = (err as { cause?: unknown }).cause;
  for (let i = 0; i < CAUSE_CAP && cause !== undefined && cause !== null; i++) {
    causes.push(cause instanceof Error ? `${cause.name}: ${cause.message}`.slice(0, 300) : String(cause).slice(0, 300));
    cause = cause instanceof Error ? (cause as { cause?: unknown }).cause : undefined;
  }
  if (causes.length) out.causes = causes;
  return out;
}

/**
 * The passive metrics/event sink the hub (and WS adapter) call at lifecycle
 * points. Every method is fire-and-forget and side-effect-free w.r.t. the
 * caller's state. Implemented by {@link MetricsObservability}; the no-op
 * {@link NO_OP_OBSERVABILITY} is the default so an un-wired hub costs nothing.
 */
export interface Observability {
  /** A transport socket opened (WS `connection`). Bumps the live-connection
   * gauge; pairs with {@link connectionClosed}. */
  connectionOpened(): void;
  /** A transport socket closed (WS `close`). Drops the live-connection gauge.
   * Counted at the transport layer (exactly once per socket) rather than in
   * the hub's `disconnect`, which a takeover kick can invoke twice. */
  connectionClosed(): void;
  /** A hello passed every gate and joined a room. */
  helloAccepted(): void;
  /** A `{t:"refused"}` frame was sent, tallied by its (fixed-vocabulary)
   * reason. A spike in `taken-over`/`already-open` is the exact single-tab
   * bug class this instrumentation exists to surface. Kicks that send a
   * refusal count here too (see {@link kicked}). */
  refused(reason: string): void;
  /** A connection was force-disconnected (`kick`), tallied by reason. A kick
   * also sends a refusal, so it additionally lands in {@link refused}. */
  kicked(reason: string): void;
  /** A submit (plaintext or encrypted) was sequenced and broadcast. */
  submitAccepted(): void;
  /** A submit (plaintext or encrypted) was refused before sequencing. */
  submitRejected(): void;
  /** A sequenced-content fan-out happened (one per submit/submit-enc). */
  broadcast(): void;
  /** A room was created (seed, encrypted seed, or provider auto-create). */
  roomCreated(): void;
  /**
   * Rooms were evicted (defaults to one), tallied by cause. "Rooms are
   * evicting" is a useless signal on its own — every room evicts eventually;
   * "rooms are evicting by `session-expired`" says a keepalive script is
   * hitting the absolute-lifetime wall, and `idle-timeout` climbing says
   * people are leaving tabs open. `reason` is the fixed EvictionReason
   * vocabulary. */
  roomEvicted(count?: number, reason?: string): void;
  /**
   * A session-ending deadline was ANNOUNCED to a room (idle or lifetime).
   * Counted separately from the eviction it precedes, because the ratio is
   * the interesting number: warnings that mostly do NOT become evictions mean
   * the grace period is doing its job and people are coming back.
   */
  sessionWarned(reason: string): void;
  /**
   * A per-IP limit refused something, tallied by which limit
   * (`ip-seed-limit` | `ip-doc-limit` | `ip-conn-limit`).
   *
   * THE ADDRESS IS NEVER PASSED HERE, and this signature is the enforcement:
   * an IP is personal data, the zero-custody rule is shapes and counts, and a
   * method that cannot accept an address cannot leak one. The rate and the
   * reason are what an operator needs; identifying WHO is a job for the
   * proxy's own logs, where that decision is made deliberately.
   *
   * Also note these are counted for BOTH transports — the WebSocket path,
   * which additionally sends a refusal frame and so also lands in
   * {@link refused}, and the HTTP seed path, which has no frame to tally.
   * That makes this map the complete picture of per-IP enforcement.
   */
  ipLimited(reason: string): void;
  /**
   * A GLOBAL ceiling refused something (`server-full` for the room cap at
   * seed, `server-busy` for the socket cap at accept).
   *
   * Deliberately NOT folded into {@link ipLimited}: those mean "this caller is
   * doing too much" and these mean "the server is full", which is not the
   * caller's fault and needs the opposite operational response — the first
   * says look at an address, the second says add capacity. Conflating them
   * would hide a saturated host inside an abuse counter.
   */
  capacityRefused(reason: string): void;
  /**
   * An intent was REJECTED by canonical validation (plaintext rooms; a blind
   * server cannot see this in encrypted ones). Tallied by the validator's
   * fixed reason vocabulary — never content.
   *
   * Worth its own counter because of how it fails: a rejection is sequenced
   * and agreed by every replica, so it produces no divergence, no error and
   * no refusal frame. The edit is simply gone. A wrong bound is invisible
   * without this number.
   */
  intentRejected(reason: string): void;
  /** A media blob was accepted by the relay (HTTP PUT, 201). */
  mediaUp(): void;
  /** A media blob was served by the relay (HTTP GET hit). */
  mediaDown(): void;
  /** A RAM blob was demoted to the spill tier (the encrypted disk write
   * completed). A climbing rate means the RAM budget is under pressure. */
  spillWrite(): void;
  /** A spilled blob was served (disk-tier GET hit). */
  spillRead(): void;
  /** A spilled blob was evicted ENTIRELY under disk pressure — the ladder's
   * last rung (RAM → disk → gone; re-supply recovers it). A climbing rate
   * means the disk budget is thrashing. */
  spillEviction(): void;
  /**
   * A spill rm FAILED and the path was parked on the hub's retry list.
   *
   * THE signal that the retry list is not draining: the spill's byte
   * accounting was already released, so a stuck file is invisible to the
   * gauges — a monotonically climbing count here is a disk-space leak (never
   * a retention leak; the per-boot key and the boot wipe still cover the
   * bytes). Counted only — the call site's {@link error} already carries the
   * line and the stack, so a second log line would double-report.
   */
  spillUnlinkFailure(): void;
  /** Media tier gauges observed after a tier mutation: bytes resident in
   * RAM, bytes on disk (IV overhead included), and spilled file count.
   * Aggregate numbers only — never a path, sha, or identifier. */
  mediaTierObserved(ramBytes: number, diskBytes: number, spillFiles: number): void;
  /** A roster fan-out happened with `size` entries; feeds the roster-size
   * peak gauge. */
  rosterObserved(size: number): void;
  /**
   * An inbound frame could not be parsed as JSON. Counted, never logged with
   * a body: the sender is untrusted, so the interesting signal is the RATE,
   * not any one malformed frame.
   */
  badFrame(): void;
  /**
   * Something threw where it should not have. The ONE event that carries a
   * stack, and the reason this interface exists beyond counting: a server
   * that cannot say what failed is a server nobody can operate.
   *
   * `where` is a fixed code-site label from the caller's own vocabulary
   * (`ws.handle`, `cli.uncaught`, …) — never interpolated with session data.
   * `fields` follows the same shapes-and-counts rule as everything else.
   */
  error(where: string, err: unknown, fields?: SafeFields): void;
  /**
   * The STARTUP BANNER (doc 16 §4 mandates it; doc 12 §2 names its wording):
   * one `server-start` line stating the storage contract and the media
   * budgets, emitted once after the port binds. Structured stderr rather
   * than stdout, which run-scripts parse and which this server keeps clean.
   *
   * `fields` is CONFIG ECHOED BACK — values the operator set (or defaults),
   * never anything derived from session data. That is why the spill dir path
   * may appear here and never in {@link snapshot}: the path came from the
   * environment, not from a room.
   */
  serverStarted(fields: SafeFields): void;
  /** Current aggregate state — safe numbers only, for the dev `/stats` route
   * and tests. */
  snapshot(): ObsSnapshot;
}

/** The shape returned by {@link Observability.snapshot} and the `/stats`
 * route — aggregate numbers only. */
export interface ObsSnapshot {
  /** Monotonic count of structured events emitted (0 when logging is off). */
  seq: number;
  /** Wall-clock ms since this sink was constructed. */
  uptimeMs: number;
  counters: {
    connectionsOpened: number;
    connectionsClosed: number;
    hellosAccepted: number;
    submitsAccepted: number;
    submitsRejected: number;
    broadcasts: number;
    roomsCreated: number;
    roomsEvicted: number;
    mediaUp: number;
    mediaDown: number;
    /** RAM → disk demotions (completed spill writes). */
    spillWrites: number;
    /** Disk-tier GET hits. */
    spillReads: number;
    /** Disk-pressure evictions (spilled blobs gone entirely). */
    spillEvictions: number;
    /** Failed spill rms parked for retry — climbing = disk-space leak. */
    spillUnlinkFailures: number;
    /** Throws caught and reported via {@link Observability.error}. The first
     * number to look at after an incident. */
    errors: number;
    /** Unparseable inbound frames (untrusted senders; rate is the signal). */
    badFrames: number;
  };
  /** Refusals tallied by protocol reason (the primary signal). */
  refusedByReason: Record<string, number>;
  /** Kicks tallied by protocol reason. */
  kicksByReason: Record<string, number>;
  /** Room evictions tallied by cause (`empty`, `idle-timeout`,
   * `session-expired`). */
  evictionsByReason: Record<string, number>;
  /** Session-ending warnings announced, by deadline (`idle`, `lifetime`). */
  warningsByReason: Record<string, number>;
  /** Per-IP limit refusals, by which limit fired. Never contains an
   * address — only the fixed limit vocabulary. */
  ipLimitsByReason: Record<string, number>;
  /** Global-ceiling refusals (`server-full`, `server-busy`). A non-zero
   * count here means ADD CAPACITY, not "find the abuser". */
  capacityByReason: Record<string, number>;
  /** Canonical validation rejections by reason — edits that died silently. */
  rejectedByReason: Record<string, number>;
  gauges: {
    /** Sockets currently open (opened − closed). */
    liveConnections: number;
    /** Rooms currently held (created − evicted). */
    rooms: number;
    /** Largest roster size observed on any fan-out since start. */
    rosterMax: number;
    /** Media bytes resident in RAM (last observed). */
    mediaRamBytes: number;
    /** Media bytes on disk in the spill tier (last observed, IV included). */
    mediaDiskBytes: number;
    /** Spilled files on disk (last observed). */
    mediaSpillFiles: number;
  };
}

/** A sink that records nothing — the hub's default, so observability is
 * strictly opt-in and an un-wired hub pays zero cost. */
export const NO_OP_OBSERVABILITY: Observability = {
  error() {},
  badFrame() {},
  connectionOpened() {},
  connectionClosed() {},
  helloAccepted() {},
  refused() {},
  kicked() {},
  submitAccepted() {},
  submitRejected() {},
  broadcast() {},
  roomCreated() {},
  roomEvicted() {},
  sessionWarned() {},
  ipLimited() {},
  capacityRefused() {},
  intentRejected() {},
  mediaUp() {},
  mediaDown() {},
  spillWrite() {},
  spillRead() {},
  spillEviction() {},
  spillUnlinkFailure() {},
  mediaTierObserved() {},
  rosterObserved() {},
  serverStarted() {},
  snapshot: () => emptySnapshot(),
};

function emptySnapshot(): ObsSnapshot {
  return {
    seq: 0,
    uptimeMs: 0,
    counters: {
      connectionsOpened: 0,
      connectionsClosed: 0,
      hellosAccepted: 0,
      submitsAccepted: 0,
      submitsRejected: 0,
      broadcasts: 0,
      roomsCreated: 0,
      roomsEvicted: 0,
      mediaUp: 0,
      mediaDown: 0,
      spillWrites: 0,
      spillReads: 0,
      spillEvictions: 0,
      spillUnlinkFailures: 0,
      errors: 0,
      badFrames: 0,
    },
    refusedByReason: {},
    kicksByReason: {},
    evictionsByReason: {},
    warningsByReason: {},
    ipLimitsByReason: {},
    capacityByReason: {},
    rejectedByReason: {},
    gauges: { liveConnections: 0, rooms: 0, rosterMax: 0, mediaRamBytes: 0, mediaDiskBytes: 0, mediaSpillFiles: 0 },
  };
}

/** Construction knobs — injected in tests for determinism, defaulted from the
 * environment in production. */
export interface ObservabilityOptions {
  /** Threshold for emitted lines. Default: {@link envLogLevel}. Counters are
   * always collected regardless of level — only the WRITING is gated, so
   * `/stats` stays accurate even at `silent`. */
  level?: LogLevel;
  /** Where a log line goes. Default: STDERR (stdout is reserved for the
   * startup lines run-scripts parse). */
  out?: (line: string) => void;
  /** Wall-clock source (epoch ms). Default `Date.now`. The zero-custody
   * server is normal Node — plain `Date.now()` for event timestamps is fine
   * here (the intent-path determinism rule governs apply/serialize values,
   * not observability wall-clock). */
  now?: () => number;
}

/**
 * The real, in-memory metrics collector + structured logger.
 *
 * Counters and gauges are plain numbers/maps mutated in place — no allocation
 * on the hot path beyond first-seen reason keys. Structured logging (when
 * enabled) writes exactly one JSON object per line to `out`, each carrying a
 * monotonic `seq` and a wall-clock `t`, plus a handful of safe fields.
 */
export class MetricsObservability implements Observability {
  private readonly level: LogLevel;
  private readonly out: (line: string) => void;
  private readonly now: () => number;
  private readonly startedAt: number;
  private seq = 0;

  private c = {
    connectionsOpened: 0,
    connectionsClosed: 0,
    hellosAccepted: 0,
    submitsAccepted: 0,
    submitsRejected: 0,
    broadcasts: 0,
    roomsCreated: 0,
    roomsEvicted: 0,
    mediaUp: 0,
    mediaDown: 0,
    spillWrites: 0,
    spillReads: 0,
    spillEvictions: 0,
    spillUnlinkFailures: 0,
    errors: 0,
    badFrames: 0,
  };
  private refusedReasons = new Map<string, number>();
  private kickReasons = new Map<string, number>();
  private evictionReasons = new Map<string, number>();
  private warningReasons = new Map<string, number>();
  private ipLimitReasons = new Map<string, number>();
  private capacityReasons = new Map<string, number>();
  private rejectedReasons = new Map<string, number>();
  private g = { liveConnections: 0, rooms: 0, rosterMax: 0, mediaRamBytes: 0, mediaDiskBytes: 0, mediaSpillFiles: 0 };

  constructor(opts: ObservabilityOptions = {}) {
    this.level = opts.level ?? envLogLevel();
    this.out = opts.out ?? ((line) => process.stderr.write(line + "\n"));
    this.now = opts.now ?? Date.now;
    this.startedAt = this.now();
  }

  /**
   * Emit one structured JSON line, gated by level. `ev` is the short event
   * name; `fields` must be safe aggregates — this method never receives
   * content, which is the invariant presence-blindness-style greps pin.
   *
   * Field order is deliberate: ts, level, ev first, so a human tailing the
   * stream reads left to right and a collector's cheap prefix match works.
   */
  private emit(level: Exclude<LogLevel, "silent">, ev: string, fields?: SafeFields): void {
    if (LEVEL_ORDER[level] < LEVEL_ORDER[this.level]) return;
    const record: SafeFields = {
      ts: new Date(this.now()).toISOString(),
      level,
      ev,
      // Monotonic per process: how a collector notices it DROPPED lines,
      // which matters for a server whose logs are the only account of what
      // happened (it persists nothing else).
      seq: ++this.seq,
    };
    if (fields) for (const k in fields) if (fields[k] !== undefined) record[k] = fields[k];
    this.out(JSON.stringify(record));
  }

  badFrame(): void {
    this.c.badFrames++;
    this.emit("debug", "bad-frame", { total: this.c.badFrames });
  }

  error(where: string, err: unknown, fields?: SafeFields): void {
    this.c.errors++;
    const e = serializeError(err);
    this.emit("error", "error", { ...fields, where, name: e.name, message: e.message, stack: e.stack, causes: e.causes?.join(" <- ") });
  }

  connectionOpened(): void {
    this.c.connectionsOpened++;
    this.g.liveConnections++;
    this.emit("info", "conn-open", { live: this.g.liveConnections });
  }
  connectionClosed(): void {
    this.c.connectionsClosed++;
    if (this.g.liveConnections > 0) this.g.liveConnections--;
    this.emit("info", "conn-close", { live: this.g.liveConnections });
  }
  helloAccepted(): void {
    this.c.hellosAccepted++;
    this.emit("info", "hello", { rooms: this.g.rooms });
  }
  refused(reason: string): void {
    this.refusedReasons.set(reason, (this.refusedReasons.get(reason) ?? 0) + 1);
    this.emit("warn", "refused", { reason });
  }
  kicked(reason: string): void {
    this.kickReasons.set(reason, (this.kickReasons.get(reason) ?? 0) + 1);
    this.emit("warn", "kick", { reason });
  }
  submitAccepted(): void {
    this.c.submitsAccepted++;
    this.emit("debug", "submit-ok");
  }
  submitRejected(): void {
    this.c.submitsRejected++;
    this.emit("warn", "submit-reject");
  }
  broadcast(): void {
    this.c.broadcasts++;
    // Deliberately unlogged per-event: broadcasts are the highest-frequency
    // signal; the counter is enough and a line per fan-out would drown the log.
  }
  roomCreated(): void {
    this.c.roomsCreated++;
    this.g.rooms++;
    this.emit("info", "room-open", { rooms: this.g.rooms });
  }
  roomEvicted(count = 1, reason = "empty"): void {
    if (count <= 0) return;
    this.c.roomsEvicted += count;
    this.evictionReasons.set(reason, (this.evictionReasons.get(reason) ?? 0) + count);
    this.g.rooms = Math.max(0, this.g.rooms - count);
    this.emit("info", "room-evict", { count, reason, rooms: this.g.rooms });
  }
  sessionWarned(reason: string): void {
    this.warningReasons.set(reason, (this.warningReasons.get(reason) ?? 0) + 1);
    // `warn` level: a session is about to end. Not an error — it is policy
    // working — but an operator scanning warnings should see it.
    this.emit("warn", "session-warning", { reason });
  }
  ipLimited(reason: string): void {
    this.ipLimitReasons.set(reason, (this.ipLimitReasons.get(reason) ?? 0) + 1);
    this.emit("warn", "ip-limit", { reason });
  }
  capacityRefused(reason: string): void {
    this.capacityReasons.set(reason, (this.capacityReasons.get(reason) ?? 0) + 1);
    this.emit("warn", "capacity", { reason });
  }
  intentRejected(reason: string): void {
    this.rejectedReasons.set(reason, (this.rejectedReasons.get(reason) ?? 0) + 1);
    this.emit("warn", "intent-rejected", { reason });
  }
  mediaUp(): void {
    this.c.mediaUp++;
    this.emit("info", "media-up");
  }
  mediaDown(): void {
    this.c.mediaDown++;
    this.emit("info", "media-down");
  }
  spillWrite(): void {
    this.c.spillWrites++;
    this.emit("info", "spill-write", { total: this.c.spillWrites });
  }
  spillRead(): void {
    this.c.spillReads++;
    this.emit("info", "spill-read", { total: this.c.spillReads });
  }
  spillEviction(): void {
    this.c.spillEvictions++;
    this.emit("info", "spill-evict", { total: this.c.spillEvictions });
  }
  spillUnlinkFailure(): void {
    // Counter only, no line: the call site's obs.error already emitted the
    // stack, and this counter exists so /stats shows the RATE without anyone
    // parsing logs. Two lines per failure would double-report one event.
    this.c.spillUnlinkFailures++;
  }
  mediaTierObserved(ramBytes: number, diskBytes: number, spillFiles: number): void {
    this.g.mediaRamBytes = ramBytes;
    this.g.mediaDiskBytes = diskBytes;
    this.g.mediaSpillFiles = spillFiles;
  }
  rosterObserved(size: number): void {
    if (size > this.g.rosterMax) this.g.rosterMax = size;
  }
  serverStarted(fields: SafeFields): void {
    this.emit("info", "server-start", fields);
  }

  snapshot(): ObsSnapshot {
    return {
      seq: this.seq,
      uptimeMs: this.now() - this.startedAt,
      counters: { ...this.c },
      refusedByReason: Object.fromEntries(this.refusedReasons),
      kicksByReason: Object.fromEntries(this.kickReasons),
      evictionsByReason: Object.fromEntries(this.evictionReasons),
      warningsByReason: Object.fromEntries(this.warningReasons),
      ipLimitsByReason: Object.fromEntries(this.ipLimitReasons),
      capacityByReason: Object.fromEntries(this.capacityReasons),
      rejectedByReason: Object.fromEntries(this.rejectedReasons),
      gauges: { ...this.g },
    };
  }
}

/** True when the `WW_OBS` env flag is set to a truthy value. */
export function envLogEnabled(): boolean {
  const v = (typeof process !== "undefined" ? process.env.WW_OBS : undefined)?.toLowerCase();
  return v === "1" || v === "true" || v === "yes" || v === "on";
}

/**
 * THE DEFAULT LEVEL when nothing is configured.
 *
 * Deliberately one constant, because it is the arc's one behavior change: the
 * server used to emit NOTHING unless WW_OBS was set, and a server whose
 * default is silence cannot be shipped to a collector — which is the whole
 * point of structured logs. Lifecycle at `info` is what an operator needs to
 * answer "is it up, who joined, what got refused" without being handed the
 * 400/s per-op firehose.
 *
 * Flip this to "silent" to restore the pre-arc behavior exactly.
 */
const DEFAULT_LEVEL: LogLevel = "info";

/**
 * Resolve the level from the environment.
 *
 * MIGRATION from the WW_OBS flag, which predates levels:
 *  - WW_LOG_LEVEL set  ⇒ wins outright (silent|debug|info|warn|error)
 *  - else WW_OBS truthy ⇒ `debug`, reproducing the old flag's firehose
 *    exactly (it emitted every event including per-op submit-ok)
 *  - else               ⇒ DEFAULT_LEVEL
 * An unrecognized WW_LOG_LEVEL falls back to the default rather than throwing:
 * a typo in an env var must not stop a server from booting.
 */
export function envLogLevel(): LogLevel {
  const raw = (typeof process !== "undefined" ? process.env.WW_LOG_LEVEL : undefined)?.toLowerCase();
  if (raw && raw in LEVEL_ORDER) return raw as LogLevel;
  if (envLogEnabled()) return "debug";
  return DEFAULT_LEVEL;
}

/**
 * Whether the dev `/stats` READ PATH is enabled. Still `WW_OBS`, deliberately
 * NOT the log level: `/stats` is an HTTP surface, and raising a log level to
 * see more lines must never silently open a network endpoint.
 */
export function observabilityEnabled(): boolean {
  return envLogEnabled();
}

/** Build the production sink from the environment. Logging follows `WW_OBS`;
 * a fresh {@link MetricsObservability} is returned either way so `/stats` can
 * read counters even when logging is quiet — the route itself is what the env
 * flag gates in `cli.ts`. */
export function createObservability(opts: ObservabilityOptions = {}): MetricsObservability {
  return new MetricsObservability(opts);
}
