import { DocxDocument } from "@wordinweb/core";
import { DocumentSession, type IdSidecar, type ParticipantProfile } from "@wordinweb/collab/server";
import type { EnvelopeEntry, SealedCheckpoint, IntentEnvelope, LineageHead } from "@wordinweb/collab/server";
import { ClientMessage, ServerMessage, PROTOCOL_VERSION, sanitizePresencePosition, type PresencePosition, type WriteStatus } from "@wordinweb/collab/server";
import type { Readable } from "node:stream";
import { StorageDriver } from "./storage.js";
import { type Observability, NO_OP_OBSERVABILITY } from "./observability.js";
import { type HubLimits, DEFAULT_LIMITS, maxSealedBytes } from "./limits.js";
import { IpGuard, NO_OP_IP_GUARD, type IpLimitReason } from "./ip-guard.js";
import { MediaSpill, newRoomSpillId, SPILL_FILE_OVERHEAD } from "./media-spill.js";

/**
 * A connection the hub can send to, abstracting the transport (WebSocket in
 * production, a fake in tests). `id` identifies the socket; `send` delivers a
 * server message.
 */
export interface Connection {
  id: string;
  send(msg: ServerMessage): void;
  /**
   * Close the underlying transport. OPTIONAL so in-process transports (tests,
   * the loopback harness) need not implement it.
   *
   * Without this the hub could refuse a client but not hang up on it. A kick —
   * an idle timeout, a lifetime expiry, a ban — sent the `refused` frame and
   * cleaned up the hub's own bookkeeping, and the socket then stayed open until
   * the client or TCP gave up. Every one of those leaked sockets still counted
   * toward the global and per-IP connection caps, so long-lived rooms ending
   * naturally could exhaust them.
   *
   * Called AFTER the refusal is sent, so the client still learns why.
   */
  close?(): void;
}

/**
 * How the hub obtains a document to start a session. Production wires this to
 * a storage driver (snapshot + log, plan doc 06); tests supply bytes inline.
 */
export interface DocProvider {
  /** Load the initial document bytes for a docId (a fresh blank doc, a stored
   * snapshot, etc.). */
  load(docId: string): Uint8Array;
}

/**
 * Verified identity/authorization for a connection (plan doc 06/07). The app's
 * backend mints a JWT; the hub only VERIFIES via an injected verifier and
 * enforces that the token authorizes the requested docId. Roles gate writes at
 * the sequencer. The hub never trusts a claimed identity — that is the
 * validate-then-install / cross-tenant-isolation rule (doc 11).
 */
export interface AuthResult {
  userId: string;
  role: "editor" | "viewer";
  /** Token expiry in epoch ms (from the JWT `exp`). When set, the hub enforces
   * it on the LIVE socket — not just at connect — so a short-lived token
   * actually limits the session (plan doc 06/11 gate 5). Omit for no expiry. */
  expiresAt?: number;
}
export interface TokenVerifier {
  /** Verify a token for a docId. Return the authorized identity, or null to
   * refuse the connection. Must reject a token whose docId does not match. */
  verify(token: string | undefined, docId: string): AuthResult | null;
}

/**
 * State of an ENCRYPTED room (doc 13 §2): the server is a blind sequencer —
 * it cannot parse, transform, validate, or apply anything here. It holds the
 * seeder's sealed checkpoint, assigns seqs to opaque envelopes in arrival
 * order, dedups by the plaintext (clientId, clientSeq) bookkeeping, and
 * retains the epoch's whole envelope log in RAM (which makes joiner tails
 * base-complete by construction — round-4 blocker 1; client-produced
 * checkpoints are a later RAM optimization, doc 13 item 6).
 */
interface EncryptedState {
  checkpoint: SealedCheckpoint;
  log: EnvelopeEntry[];
  seen: Set<string>;
  /** conn.id of the server-ASSIGNED checkpointer (doc 13 §3, blocker 2:
   * assignment replaces the riggable lowest-clientId election — rigging
   * now requires being picked). Rotated when that socket leaves. */
  checkpointerConnId?: string;
  /** The engine version the epoch's FIRST participant registered; everyone
   * after must match (client-derived canonical forms diverge across
   * transform-semantics versions with no arbiter — doc 13 §2). */
  engineVersion?: string;
  /**
   * Running total of retained ciphertext bytes in {@link log}, maintained
   * incrementally rather than recomputed: the whole point of the budget is to
   * bound a room under sustained submission, and summing the log on every
   * submit would itself become the quadratic cost the attack is looking for.
   * Reset to 0 when a checkpoint prunes the log.
   */
  logBytes: number;
}

/**
 * One room's resource usage, each measure beside the limit it runs against.
 * See {@link CollabHub.roomsSnapshot}. Contains no identifier of any kind —
 * `room` is the opaque, freshly-random `Room.obsLabel`.
 */
export interface RoomSnapshot {
  room: string;
  mode: "plaintext" | "encrypted";
  ageMs: number;
  ageLimitMs: number;
  idleMs: number;
  idleLimitMs: number;
  emptyMs: number | null;
  emptyLimitMs: number;
  conns: number;
  rosterTotal: number;
  rosterConnected: number;
  connLimit: number;
  /** Media bytes resident in RAM for this room. Together with
   * `mediaDiskBytes` these sum to the figure `mediaLimitBytes` caps — the
   * per-room quota is tier-agnostic (spilling never buys a room capacity). */
  mediaRamBytes: number;
  /** Media bytes this room holds in the disk spill tier (blob lengths, not
   * on-disk sizes — the same basis the room quota counts). */
  mediaDiskBytes: number;
  mediaLimitBytes: number;
  mediaBlobs: number;
  mediaWaiting: number;
  uploadTokens: number;
  uploadBurst: number;
  logBytes: number;
  logLimitBytes: number;
  readOnly: boolean;
  idleWarned: boolean;
  lifetimeWarned: boolean;
}

/**
 * One locker entry, discriminated by WHERE the bytes live (doc 16 §4
 * two-tier index). A disk entry keeps only the blob length (for byte
 * accounting); its bytes live in the spill, keyed by the same sha.
 *
 * `demoting` marks a RAM entry whose spill write is in flight. The RAM
 * buffer is HELD until the tmp-then-rename write completes and the index
 * entry swaps — so a GET racing a demotion always finds either the buffer
 * or a complete file, never a partial one. The flag only stops a second
 * concurrent demotion from picking the same victim.
 */
type MediaBlobEntry =
  | { loc: "ram"; bytes: Uint8Array; lastDownloadAt: number; staged: boolean; demoting?: boolean }
  | { loc: "disk"; len: number; lastDownloadAt: number; staged: boolean };

/**
 * What a media GET serves: RAM blobs come back as the bytes themselves
 * (today's shape, unchanged); disk blobs come back as a length plus a lazy
 * decrypting stream so the transport pipes them without ever buffering the
 * file — the spill tier is exactly the tier whose blobs should not re-enter
 * RAM whole (the one-way ladder, doc 16 §4).
 */
export type MediaDownloadHit = Uint8Array | { len: number; open: () => Promise<Readable> };

/** Blob length regardless of tier — the number the room byte cap counts. */
function mediaBlobLen(b: MediaBlobEntry): number {
  return b.loc === "ram" ? b.bytes.length : b.len;
}

interface Room {
  /**
   * Opaque label for the OPS VIEW, and the reason per-room metrics can exist
   * at all.
   *
   * NOT derived from the docId, and not a hash of it — freshly random. The
   * docId is the capability that opens the document, so a metrics row keyed by
   * it turns an ops dashboard into a way into other people's documents, and a
   * screenshot of that dashboard leaks them. A derived key would be safe in
   * practice and still invites "we can reverse this if we need to"; unrelated
   * randomness removes the question.
   *
   * It identifies a room only for as long as the room exists, which is the
   * whole useful lifetime — rooms do not survive a restart by design.
   */
  obsLabel: string;
  /** The authoritative session — PLAINTEXT rooms only (null when `enc`). */
  session: DocumentSession | null;
  enc: EncryptedState | null;
  conns: Set<Connection>;
  /** The session EPOCH id (plan doc 12): minted fresh at every seed/re-seed
   * (and at provider-based creation). Rejoining clients compare it against
   * their bundle's stored epoch to distinguish seamless resume from
   * someone-re-seeded-while-I-was-away — epochs never merge. */
  genesisId: string;
  /** Session roster (doc 14 §2): clientId → sanitized profile + liveness.
   * Ephemeral like presence — never persisted, dies with the room.
   * Disconnected entries stay (greyed in UI) so attribution keeps a name
   * for everyone who touched the session; reconnects resume the entry. */
  roster: Map<string, { profile: ParticipantProfile; connected: boolean; participantType: "human" | "agent" }>;
  /** Short-lived bearer capabilities created by participants for AI agents. */
  agentInvites: Map<string, {
    token: string;
    inviterClientId: string;
    expiresAt: number;
    agentClientId?: string;
  }>;
  /** Connected or reconnectable AI identities and their private chat owner. */
  agentClients: Map<string, { inviteId: string; inviterClientId: string }>;
  /** Share-code verifier registered at seed (doc 13 §7) — a PBKDF2 output,
   * NOT the code. Optional; rotation happens naturally at re-seed. */
  codeVerifier?: string;
  /** Owner capability token (doc 14 §2.5): minted at seed, returned only
   * to the seeder. 256-bit — direct compare, no stretching needed. */
  ownerToken?: string;
  /** Owner-imposed write controls (integrity, not confidentiality). */
  readOnly?: boolean;
  demoted?: Set<string>;
  banned?: Set<string>;
  /** The seeder's lineage chain (doc 15) — held OPAQUELY and echoed in
   * welcomes so rejoining holders can decide fast-forward vs fork.
   * Ephemeral like everything else in the room. */
  lineage?: LineageHead[];
  /** Online-guess budget: consecutive failures + lockout deadline. */
  codeGuard?: { failures: number; lockedUntil: number };
  /** Media relay state (doc 16 §4): the RAM hot tier of the locker. The
   * locker is a CACHE, not a store — every promoted blob is recoverable
   * from participants via re-supply, so eviction costs latency, never
   * data. (The disk spill tier of doc 16 §4 is capacity engineering on
   * top of this same state machine; RAM-only is correct by the same
   * cache-not-store argument.) */
  media: {
    blobs: Map<string, MediaBlobEntry>;
    waiters: Map<string, Set<string>>; // sha -> conn.ids awaiting media-ready
    resupply: Map<string, { chosenConnId: string; deadline: number }>;
    totalBytes: number;
    /** Per-room upload admission bucket (see MediaLimits.roomUploadsPerMin).
     * Lives ON THE ROOM so it is freed when the room is, rather than in a
     * hub-level map keyed by docId — that map would accumulate an entry per
     * docId anyone ever probed, which is a slow leak an attacker controls the
     * size of, in the one component whose whole promise is holding nothing. */
    uploadBucket: { tokens: number; last: number };
  };
  /** This room's directory name in the spill tier — FRESH RANDOM, minted on
   * the room's first demotion (see newRoomSpillId for why it is never
   * derived from the docId). Unset until the room spills something. */
  spillId?: string;
  /** Epoch-ms timestamp of when the room last became empty, or undefined
   * while anyone is connected. Drives zero-custody eviction (plan doc 12
   * §2): after the last-disconnect grace elapses the room — doc, log,
   * everything — is deleted. Only the empty state starts THIS clock; a room
   * with people in it is governed by the idle/lifetime deadlines below
   * instead (round-4 F7 said "idle-with-people never evicts", which was the
   * right call while eviction had no warning and no countdown — the server
   * lifecycle arc supersedes it with a graceful, announced ending rather
   * than a silent disappearance). */
  emptySince?: number;
  /** When this room was created (seed, re-seed, or provider auto-create) —
   * the fixed origin of the ABSOLUTE lifetime cap. Nothing resets it: that
   * is the entire point, since an activity-based clock is exactly what a
   * keepalive script defeats. */
  createdAt: number;
  /** Last QUALIFYING activity (see LIFECYCLE_NOTES.activity in limits.ts):
   * accepted submits, admin actions, media transfers, and joins — never
   * presence. Drives the idle deadline. */
  lastActivityAt: number;
  /** True once the idle warning has been sent for the CURRENT idle stretch,
   * so it goes out once rather than on every sweep. Cleared by activity,
   * which also sends the cancellation. */
  idleWarned?: boolean;
  /** True once the lifetime warning has been sent. Never cleared — the
   * deadline it announces cannot move. */
  lifetimeWarned?: boolean;
  /** Normalized address of whoever seeded this room, for the per-IP live-doc
   * cap. Held ONLY as a guard-internal bucket key (see ip-guard.ts): never
   * logged, never sent, never in a snapshot. */
  creatorIp?: string;
}

/** Last-disconnect grace before an empty room is deleted (plan doc 12 §1):
 * long enough to survive a page refresh or a flaky reconnect, short enough
 * that "everyone left" promptly means "the server holds nothing". */
export const EVICTION_GRACE_MS = 60_000;

/**
 * A room's death certificate. Fixed vocabulary (safe to log and tally by),
 * and distinct per cause because "rooms are evicting" is a useless signal
 * while "rooms are evicting by `session-expired`" says a keepalive script is
 * hitting the 12h wall.
 *
 *  - `empty`           everyone left and the grace elapsed (the zero-custody
 *                      promise's normal enforcement path)
 *  - `idle-timeout`    nobody did anything for the idle window
 *  - `session-expired` the absolute lifetime cap
 */
export type EvictionReason = "empty" | "idle-timeout" | "session-expired" | "log-overflow";

/**
 * Outcome of a seed attempt. Three distinct failures, kept distinct because
 * the caller maps them to different HTTP statuses and the client to different
 * user-facing stories: `exists` is first-wins (join the incumbent epoch — it
 * carries the genesisId to join), while the ip-limit reasons are throttles
 * that carry no epoch because no epoch was consulted.
 */
export type SeedResult =
  | { ok: true; genesisId: string; ownerToken: string }
  | { ok: false; reason: "exists"; genesisId: string }
  | { ok: false; reason: IpLimitReason }
  /** The GLOBAL room ceiling — the whole server is at capacity, not this
   * address. Distinct from the ip-limit reasons because it is not the
   * caller's fault and the honest client answer is "try again shortly"
   * rather than "you are doing too much". Maps to 503, not 429. */
  | { ok: false; reason: "server-full" };

/**
 * Routes clients to per-document sessions and fans out broadcasts. This is the
 * thin plumbing layer (plan doc 06): it owns no editing logic — the
 * DocumentSession is the authority — and no transport, so it is unit-testable
 * with in-memory connections.
 */
export class CollabHub {
  private rooms = new Map<string, Room>();
  private connDoc = new Map<string, string>();

  /** Optional persistence (plan doc 06). Without a driver the hub is purely
   * in-memory (a session lives only while a room is held). With one, rooms
   * rehydrate from snapshot+log on first access and every sequenced entry is
   * durably appended before broadcast. */
  private auth = new Map<string, AuthResult>();
  private connById = new Map<string, Connection>();
  /**
   * conn.id → the clientId bound at hello (plan doc 11 decision 8, round-4
   * F4). The SECURITY-load-bearing map: `(clientId, clientSeq)` is the
   * session's idempotency key, so an unbound claimed id would let any
   * participant submit under a victim's identity — the victim's own next
   * intent then dedups against the forgery and is silently swallowed, and
   * every attribution/roster/election feature inherits the forged identity.
   * Bound once per socket; every submit and presence message is checked
   * against it.
   */
  private connClient = new Map<string, string>();
  /** conn.id set of proven owners (doc 14 §2.5) — flagged at hello by
   * token match, consumed by the admin channel + read-only bypass. */
  private ownerConns = new Set<string>();

  constructor(
    /**
     * Template source for auto-created rooms (party/dev mode: a hello to an
     * unknown docId gets a fresh doc from here). Pass `null` for the
     * zero-custody magic-link posture (plan doc 12 §5 case 3): rooms then
     * exist ONLY via seed()/re-seed, and a hello to an unknown docId is
     * refused `no-session` — the client offers "Bring it back live" from its
     * bundle instead.
     */
    private provider: DocProvider | null,
    private storage?: StorageDriver,
    /** Optional token verifier. When set, every hello must present a token
     * that authorizes the requested docId; without it the hub is auth-off
     * (dev/demo only). */
    private verifier?: TokenVerifier,
    /** Injectable clock (epoch ms) for token-expiry checks; defaults to
     * Date.now. Injected in tests for determinism. */
    private now: () => number = () => Date.now(),
    /** Injectable epoch-id generator (fresh 128-bit hex per seed). Injected
     * in tests for determinism; production default uses crypto randomness. */
    private genGenesisId: () => string = defaultGenesisId,
    /** Passive observability sink (counters/events). Constructor-injected
     * with a no-op default so it never changes behavior and tests can assert
     * on it, mirroring the clock/provider pattern above. Records shapes and
     * counts only — never content (zero-custody). */
    private obs: Observability = NO_OP_OBSERVABILITY,
    /** Lifecycle + abuse limits (see limits.ts for the whole config surface
     * and its env overrides). Resolved ONCE at startup and injected, so no
     * code path below reads process.env and tests hand over exact numbers
     * to drive against the injected clock. */
    private limits: HubLimits = DEFAULT_LIMITS,
    /** Per-IP abuse guard. Defaults to the permissive no-op, so a hub built
     * without one behaves exactly as it did before this arc — same pattern
     * as the observability sink above. */
    private ipGuard: IpGuard = NO_OP_IP_GUARD,
  ) {
    // The disk spill tier (doc 16 §4). `diskBytes: 0` — the default — means
    // NO spill: no MediaSpill is constructed, nothing touches the filesystem,
    // and the RAM-pressure ladder below never runs, so the default path is
    // byte-for-byte the pre-spill behavior.
    this.spill = this.limits.media.diskBytes > 0 ? new MediaSpill(this.limits.media.spillDir) : null;
  }

  /** The spill tier, or null when `diskBytes` is 0 (spill disabled). */
  private readonly spill: MediaSpill | null;
  /** Global gauge of media bytes RESIDENT IN RAM across all rooms — the
   * number `limits.media.ramBytes` bounds. Maintained unconditionally (it is
   * one integer), enforced only when the spill tier is on. */
  private mediaRamBytes = 0;
  /** Spill files/directories whose rm FAILED (EACCES, EIO, …): retried on
   * every sweepMedia. Accounting was already released, so a stuck entry is a
   * disk-space bug, never a retention bug — the boot wipe and the per-boot
   * key both still cover the bytes. `sha` unset means the room's whole
   * directory. */
  private spillRetry: { roomSpillId: string; sha?: string }[] = [];

  /** Send a refusal and tally it by reason (the one place `{t:"refused"}` is
   * emitted outside `kick`), so the refused-by-reason counter always tracks
   * the wire exactly. Behavior is byte-identical to sending the frame inline. */
  private refuse(conn: Connection, reason: string): void {
    conn.send({ t: "refused", reason });
    this.obs.refused(reason);
  }

  /** Refuse an intent in a submit/submit-enc handler: same as {@link refuse}
   * plus the submits-rejected tally. (Token-expiry cuts go through `kick` and
   * count as session cuts, not submit rejections.) */
  private refuseSubmit(conn: Connection, reason: string): void {
    this.refuse(conn, reason);
    this.obs.submitRejected();
  }

  /** Per-connection submit token bucket (flood guard, perf B10): generous
   * for humans (fast typing ≈ 10 ops/s) AND for well-behaved ack-paced
   * bulk clients, but caps a pathological flooder. Tuned an order of
   * magnitude above humans and ABOVE the fastest legitimate ack-paced
   * client (post-scoped-resync seeding sustains ~210 ops/s locally) —
   * re-measure after every perf arc; every legit-client speedup so far
   * has re-tripped a cap tuned to the previous generation (25 → 80 →
   * 400). A rate-limited submit is NOT sequenced and the refusal is
   * NON-FATAL by contract: honest clients keep the op pending and
   * re-drive the whole queue after backoff — no lost work. */
  private buckets = new Map<string, { tokens: number; last: number }>();
  private static RATE_PER_SEC = 400;
  private static BURST = 800;
  /** Room participant cap (live sockets, distinct clientIds). Refusal
   * reason "room-full"; reconnect/takeover of an in-room identity is
   * exempt. 10 for now per owner decision — raise deliberately, with the
   * swarm suite re-run at the new number first. WW_ROOM_CAP env override
   * exists for the stress harnesses (scripts/swarm.mjs drives up to 30
   * clients into one room), NOT for production tuning. */
  private static MAX_ROOM_CLIENTS = Number(process.env.WW_ROOM_CAP ?? 10) || 10;
  private static MAX_AGENT_INVITES = 20;
  private static MAX_AGENT_INVITE_LIFETIME_MS = 15 * 60_000;
  private static MAX_AGENT_CHAT_CIPHERTEXT = 48_000;

  private refuseAgentInvite(conn: Connection, inviteId: string, reason: string): void {
    conn.send({ t: "agent-invite-refused", inviteId, reason });
  }

  private pruneAgentInvites(room: Room): void {
    const now = this.now();
    for (const [inviteId, invite] of room.agentInvites) {
      if (invite.expiresAt <= now && !invite.agentClientId) room.agentInvites.delete(inviteId);
    }
  }
  private rateLimited(connId: string): boolean {
    const now = this.now();
    let b = this.buckets.get(connId);
    if (!b) {
      b = { tokens: CollabHub.BURST, last: now };
      this.buckets.set(connId, b);
    }
    b.tokens = Math.min(CollabHub.BURST, b.tokens + ((now - b.last) / 1000) * CollabHub.RATE_PER_SEC);
    b.last = now;
    if (b.tokens < 1) return true;
    b.tokens -= 1;
    return false;
  }

  /** True if the connection's token has expired (live-socket enforcement). */
  private expired(connId: string): boolean {
    const a = this.auth.get(connId);
    return a?.expiresAt !== undefined && a.expiresAt <= this.now();
  }

  /** Force-disconnect a connection with a refusal (expiry/revocation). A kick
   * sends a refusal frame, so it tallies BOTH the kick-by-reason and the
   * refused-by-reason counter — a `taken-over` kick shows up in each, which
   * is the single-tab regression signal. */
  private kick(conn: Connection, reason: string): void {
    conn.send({ t: "refused", reason });
    this.obs.kicked(reason);
    this.obs.refused(reason);
    this.disconnect(conn);
    // HANG UP, not just refuse. `disconnect` clears the hub's bookkeeping but
    // cannot touch the transport, so before this the socket outlived the
    // session it belonged to — a room reaching its idle timeout left one open
    // socket per participant, counting against the global and per-IP caps
    // until the client noticed. Ordered after the refusal so the reason is
    // delivered first, and after `disconnect` so the ws close handler finds
    // the bookkeeping already clean and its own teardown is a no-op.
    try {
      conn.close?.();
    } catch {
      // Already gone. The refusal and the bookkeeping are what matter.
    }
  }

  /**
   * Revoke all live sessions for a user (the app's backend calls this when it
   * deprovisions/downgrades someone — plan doc 06 gate 5). They are refused
   * and disconnected immediately, not at next reconnect.
   */
  revoke(userId: string): void {
    for (const [connId, a] of [...this.auth]) {
      if (a.userId === userId) {
        const conn = this.connById.get(connId);
        if (conn) this.kick(conn, "revoked");
      }
    }
  }

  /** Sweep expired sessions (call periodically from the transport). */
  sweepExpired(): void {
    for (const [connId] of [...this.auth]) {
      if (this.expired(connId)) {
        const conn = this.connById.get(connId);
        if (conn) this.kick(conn, "token-expired");
      }
    }
  }

  /** Handle an inbound message from a connection. Async because a storage-
   * backed hello rehydrates and a submit persists before broadcast. */
  async handle(conn: Connection, msg: ClientMessage): Promise<void> {
    switch (msg.t) {
      case "hello": {
        if (msg.protocolVersion !== PROTOCOL_VERSION) {
          this.refuse(conn, "version-mismatch");
          return;
        }
        if (!msg.clientId) {
          // The bound identity is not optional — every security property
          // downstream (dedup, attribution, single-tab) hangs off it.
          this.refuse(conn, "client-id-required");
          return;
        }
        if (this.verifier) {
          const authed = this.verifier.verify(msg.token, msg.docId);
          if (!authed) {
            this.refuse(conn, "unauthorized");
            return;
          }
          this.auth.set(conn.id, authed);
        }
        // Single-live-connection rule (plan doc 12 §7): one socket per
        // (docId, clientId). Same-profile tabs share the persistent clientId
        // and would collide clientSeq counters — self-inflicted dedup
        // poisoning — so the duplicate is refused. `takeover` inverts the
        // outcome (the old socket may be a zombie tab): the incumbent is
        // kicked and the new connection claims the identity. A different
        // browser/profile/device mints a different clientId and is untouched.
        const holder = this.findClientConn(msg.docId, msg.clientId);
        if (holder && holder.id !== conn.id) {
          if (msg.takeover) this.kick(holder, "taken-over");
          else {
            this.refuse(conn, "already-open");
            return;
          }
        }
        // Room capacity (owner decision 2026-07-27: "max 10 editors per
        // document for now"). Counted as LIVE SOCKETS holding a DIFFERENT
        // clientId — a reconnect/takeover of an identity already in the
        // room must never be blocked by the cap (the zombie it replaces
        // may still linger in conns until its close fires).
        {
          const existing = this.rooms.get(msg.docId);
          if (existing) {
            let others = 0;
            for (const c of existing.conns) {
              if (this.connClient.get(c.id) !== msg.clientId) others++;
            }
            if (others >= CollabHub.MAX_ROOM_CLIENTS) {
              this.refuse(conn, "room-full");
              return;
            }
          }
        }
        // Share-code gate (doc 13 §7): server-side attempt budget makes a
        // 6-digit space adequate (5 tries/lockout vs 10^6). Checked before
        // any state is bound for this socket. In E2EE mode this is only the
        // OUTER wall — the code also mixes into key derivation, so a
        // hypothetically bypassed gate still yields ciphertext.
        {
          const guarded = this.rooms.get(msg.docId);
          if (guarded?.codeVerifier) {
            const g = (guarded.codeGuard ??= { failures: 0, lockedUntil: 0 });
            if (g.lockedUntil > this.now()) {
              this.refuse(conn, "code-locked");
              return;
            }
            if (!msg.codeProof) {
              this.refuse(conn, "code-required");
              return;
            }
            if (msg.codeProof !== guarded.codeVerifier) {
              g.failures++;
              if (g.failures >= 5) {
                g.failures = 0;
                g.lockedUntil = this.now() + 60_000; // 1-min lockout window
              }
              this.refuse(conn, "code-invalid");
              return;
            }
            g.failures = 0; // success resets the budget
          }
        }
        // Owner controls (doc 14 §2.5): banned clientIds stay out; a hello
        // proving the owner token flags this socket as the epoch's owner.
        {
          const r0 = this.rooms.get(msg.docId);
          if (r0?.banned?.has(msg.clientId)) {
            this.refuse(conn, "kicked");
            return;
          }
          if (r0?.ownerToken && msg.ownerToken === r0.ownerToken) this.ownerConns.add(conn.id);
        }
        // Zero-custody posture (no provider): rooms exist only while seeded
        // and live. An unknown docId is not an error the server can fix —
        // the DOCUMENT lives in participants' bundles, so the answer is
        // `no-session`: the client offers "Bring it back live" (doc 12 §5.3)
        // and re-seeds from its own copy.
        if (!this.rooms.has(msg.docId) && !this.provider) {
          this.refuse(conn, "no-session");
          return;
        }
        let agentPair: { inviteId: string; inviterClientId: string } | null = null;
        {
          const inviteRoom = this.rooms.get(msg.docId);
          const previous = inviteRoom?.agentClients.get(msg.clientId);
          if (msg.agentInvite) {
            const invite = inviteRoom?.agentInvites.get(msg.agentInvite.inviteId);
            const usable = invite &&
              invite.token === msg.agentInvite.token &&
              (!invite.agentClientId || invite.agentClientId === msg.clientId) &&
              (invite.expiresAt > this.now() || invite.agentClientId === msg.clientId);
            if (!usable) {
              this.refuse(conn, "agent-invite-invalid");
              return;
            }
            invite.agentClientId = msg.clientId;
            agentPair = { inviteId: msg.agentInvite.inviteId, inviterClientId: invite.inviterClientId };
          } else if (previous) {
            this.refuse(conn, "agent-invite-required");
            return;
          }
        }
        // Encrypted rooms enforce the ENGINE fence before admission
        // (doc 13 §2): canonical forms are client-derived there, so mixed
        // transform semantics would diverge with no arbiter. First joiner
        // registers the epoch's version; later mismatches are refused with
        // the download/draft path client-side.
        const encRoom = this.rooms.get(msg.docId)?.enc ?? null;
        if (encRoom) {
          if (!msg.engineVersion) {
            this.refuse(conn, "engine-version-required");
            return;
          }
          if (encRoom.engineVersion && encRoom.engineVersion !== msg.engineVersion) {
            this.refuse(conn, "engine-version-mismatch");
            return;
          }
          encRoom.engineVersion = msg.engineVersion;
        }
        this.connClient.set(conn.id, msg.clientId);
        const room = await this.room(msg.docId);
        if (agentPair) room.agentClients.set(msg.clientId, agentPair);
        room.conns.add(conn);
        room.emptySince = undefined; // occupied: the eviction clock stops.
        this.connDoc.set(conn.id, msg.docId);
        this.connById.set(conn.id, conn);
        this.obs.helloAccepted();
        // A JOIN counts as activity (LIFECYCLE_NOTES.activity). Without this,
        // someone opening a link into a room that has been quiet for nine
        // minutes gets a 60-second countdown and watches the session evict
        // under them before they type a character. Keepalive-by-rejoin is
        // bounded by the absolute lifetime cap and the per-IP connection
        // limit, so admitting joins here does not reopen the hole the idle
        // timeout closes.
        this.noteActivity(room);
        if (room.enc) {
          // Blind welcome (doc 13 §3): sealed seed checkpoint + the WHOLE
          // epoch envelope log after it — base-complete by construction
          // because nothing is ever pruned within an epoch (blocker 1).
          conn.send({
            t: "welcome-enc",
            docId: msg.docId,
            genesisId: room.genesisId,
            checkpoint: room.enc.checkpoint,
            // BASE-COMPLETE tail (blocker 1): quiescent checkpoints prune
            // to empty, and the stale-base guard keeps every later entry's
            // base ≥ the checkpoint — so the retained log is exactly what a
            // joiner needs, always.
            tail: room.enc.log,
            mode: "encrypted",
            // The CONFIGURED value, never the default constant: a client that
            // pre-checks against a number this server does not enforce is
            // worse than one that does not pre-check at all.
            mediaMaxBlobBytes: this.limits.media.maxBlobBytes,
          });
          const joinedProfile = sanitizeProfile(msg.profile, msg.clientId);
          room.roster.set(msg.clientId, {
            profile: joinedProfile,
            connected: true,
            participantType: agentPair ? "agent" : "human",
          });
          this.broadcastRoster(room);
          if (agentPair) {
            this.findClientConn(msg.docId, agentPair.inviterClientId)?.send({
              t: "agent-connected",
              inviteId: agentPair.inviteId,
              agentClientId: msg.clientId,
              profile: joinedProfile,
            });
          }
          this.assignCheckpointer(room);
          this.sendPendingWarning(room, conn);
          return;
        }
        // The welcome is a checkpoint bundle (plan doc 12 §2, round-4 F10):
        // snapshot + the id sidecar, never the bytes alone — a joiner cannot
        // re-derive the id table from parse order once history contains
        // split-created carried ids. The snapshot is the CURRENT document
        // (at room.session.seq), so the tail after it is empty — sending
        // entries the snapshot already contains would double-apply them.
        const cp = room.session!.checkpoint();
        conn.send({
          t: "welcome",
          docId: msg.docId,
          seq: cp.seq,
          snapshot: bytesToBase64(cp.docx),
          sidecar: cp.sidecar,
          tail: room.session!.entriesSince(cp.seq),
          genesisId: room.genesisId,
          mode: "plaintext", // E2EE mode lands with doc 13 items 1-2.
          lineage: room.lineage,
          mediaNeeded: [...room.media.waiters.keys()],
          // Late-join media addresses (doc 16 §6): the snapshot registers
          // the parts but the sha/iv live only in already-folded intents —
          // without this map a joiner can reserve the box yet never fetch.
          // Parse-derived holes with unknown sha (sha === "") are omitted.
          media: [...room.session!.doc.mediaMeta.entries()]
            .filter(([, m]) => m.sha !== "")
            .map(([part, m]) => ({ part, sha: m.sha, ...(m.iv ? { iv: m.iv } : {}) })),
          mediaMaxBlobBytes: this.limits.media.maxBlobBytes,
        });
        // Roster upsert + fan-out (doc 14 §2): keyed by the BOUND clientId,
        // so a reconnect resumes the same entry. Sanitization is server-side
        // here (plaintext mode); clients still render defensively (doc 11
        // vector 7), which is also what E2EE mode will rely on entirely.
        const joinedProfile = sanitizeProfile(msg.profile, msg.clientId);
        room.roster.set(msg.clientId, {
          profile: joinedProfile,
          connected: true,
          participantType: agentPair ? "agent" : "human",
        });
        this.broadcastRoster(room);
        if (agentPair) {
          this.findClientConn(msg.docId, agentPair.inviterClientId)?.send({
            t: "agent-connected",
            inviteId: agentPair.inviteId,
            agentClientId: msg.clientId,
            profile: joinedProfile,
          });
        }
        this.sendPendingWarning(room, conn);
        return;
      }
      case "agent-invite": {
        const docId = this.connDoc.get(conn.id);
        const inviterClientId = this.connClient.get(conn.id);
        if (!docId || !inviterClientId) return;
        const room = this.rooms.get(docId)!;
        this.pruneAgentInvites(room);
        const validId = /^[A-Za-z0-9_-]{16,128}$/.test(msg.inviteId);
        const validToken = /^[A-Za-z0-9_-]{32,256}$/.test(msg.token);
        const validExpiry = Number.isFinite(msg.expiresAt) &&
          msg.expiresAt > this.now() &&
          msg.expiresAt <= this.now() + CollabHub.MAX_AGENT_INVITE_LIFETIME_MS;
        const existing = room.agentInvites.get(msg.inviteId);
        if (!validId || !validToken || !validExpiry) {
          this.refuseAgentInvite(conn, msg.inviteId, "invalid");
          return;
        }
        if (existing && existing.inviterClientId !== inviterClientId) {
          this.refuseAgentInvite(conn, msg.inviteId, "already-exists");
          return;
        }
        if (!existing && room.agentInvites.size >= CollabHub.MAX_AGENT_INVITES) {
          this.refuseAgentInvite(conn, msg.inviteId, "room-limit");
          return;
        }
        room.agentInvites.set(msg.inviteId, { token: msg.token, inviterClientId, expiresAt: msg.expiresAt });
        conn.send({ t: "agent-invite-registered", inviteId: msg.inviteId, expiresAt: msg.expiresAt });
        this.noteActivity(room);
        return;
      }
      case "agent-invite-revoke": {
        const docId = this.connDoc.get(conn.id);
        const inviterClientId = this.connClient.get(conn.id);
        if (!docId || !inviterClientId) return;
        const room = this.rooms.get(docId)!;
        const invite = room.agentInvites.get(msg.inviteId);
        if (!invite || invite.inviterClientId !== inviterClientId) return;
        room.agentInvites.delete(msg.inviteId);
        if (invite.agentClientId) {
          room.agentClients.delete(invite.agentClientId);
          const agentConn = this.findClientConn(docId, invite.agentClientId);
          if (agentConn) this.kick(agentConn, "agent-invite-revoked");
        }
        this.noteActivity(room);
        return;
      }
      case "agent-chat": {
        const docId = this.connDoc.get(conn.id);
        const senderClientId = this.connClient.get(conn.id);
        if (!docId || !senderClientId) return;
        const room = this.rooms.get(docId)!;
        if (!/^[A-Za-z0-9_-]{8,128}$/.test(msg.messageId) ||
            !/^[A-Za-z0-9_-]{8,128}$/.test(msg.iv) ||
            !/^[A-Za-z0-9_-]+$/.test(msg.ciphertext) ||
            msg.ciphertext.length > CollabHub.MAX_AGENT_CHAT_CIPHERTEXT) return;
        const pair = room.agentClients.get(msg.agentClientId);
        if (!pair) return;
        const sender = senderClientId === msg.agentClientId
          ? "agent" as const
          : senderClientId === pair.inviterClientId
            ? "inviter" as const
            : null;
        if (!sender) return;
        const out: ServerMessage = {
          t: "agent-chat",
          agentClientId: msg.agentClientId,
          messageId: msg.messageId,
          sender,
          iv: msg.iv,
          ciphertext: msg.ciphertext,
        };
        const agentConn = this.findClientConn(docId, msg.agentClientId);
        const inviterConn = this.findClientConn(docId, pair.inviterClientId);
        agentConn?.send(out);
        if (inviterConn && inviterConn.id !== agentConn?.id) inviterConn.send(out);
        this.noteActivity(room);
        return;
      }
      case "admin": {
        const docId = this.connDoc.get(conn.id);
        if (!docId || !this.ownerConns.has(conn.id)) {
          this.refuse(conn, "not-owner");
          return;
        }
        const room = this.rooms.get(docId)!;
        const a = msg.action;
        if (a.op === "readOnly") {
          room.readOnly = a.on;
        } else if (a.op === "setRole") {
          (room.demoted ??= new Set()).delete(a.clientId);
          if (a.role === "viewer") room.demoted!.add(a.clientId);
        } else if (a.op === "kick") {
          (room.banned ??= new Set()).add(a.clientId);
          const target = this.findClientConn(docId, a.clientId);
          if (target) this.kick(target, "kicked");
        }
        // RE-FAN THE ROSTER: every admin action changes somebody's write
        // status, and the transition has to reach clients WITHOUT them
        // attempting a write to discover it. That is the whole point of the
        // signal — a lift used to be undiscoverable (a blocked client stayed
        // in viewer mode until reload, because the only evidence of the lock
        // was a per-edit refusal and nothing announced its removal).
        this.broadcastRoster(room);
        // An owner administering the room is unmistakably a live human.
        this.noteActivity(room);
        return;
      }
      case "profile": {
        const docId2 = this.connDoc.get(conn.id);
        if (!docId2) return; // ignore before join, like presence
        const room2 = this.rooms.get(docId2)!;
        const clientId = this.connClient.get(conn.id)!;
        const entry = room2.roster.get(clientId);
        if (entry) {
          entry.profile = sanitizeProfile(msg.profile, clientId);
          this.broadcastRoster(room2);
        }
        return;
      }
      case "submit": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) {
          this.refuseSubmit(conn, "not-joined");
          return;
        }
        // Live-socket token-expiry enforcement (gate 5): a session outlives
        // its token only until the next action, then is cut.
        if (this.expired(conn.id)) {
          this.kick(conn, "token-expired");
          return;
        }
        if (this.rateLimited(conn.id)) {
          this.refuseSubmit(conn, "rate-limit");
          return;
        }
        // Identity binding (doc 11 decision 8): the intent must carry the
        // clientId this socket registered at hello. Anything else is a
        // forgery attempt (or a client bug) — refused BEFORE the session
        // sees it, so a forged (clientId, clientSeq) can never enter the
        // dedup set and swallow the victim's real intent.
        if (msg.intent.clientId !== this.connClient.get(conn.id)) {
          this.refuseSubmit(conn, "client-id-mismatch");
          return;
        }
        const room = this.rooms.get(docId)!;
        // Write controls, ALL THREE through the one predicate the roster also
        // advertises (see writeBlock): role, demotion, and the room-wide lock
        // for every participant. Two copies of this rule is how the advertised
        // status and the enforced one drift apart, and a client told "you may
        // write" while the server refuses is worse off than one told nothing.
        if (this.writeBlock(room, conn, msg.intent.clientId)) {
          this.refuseSubmit(conn, "read-only");
          return;
        }
        if (room.enc) {
          // No mixed-mode documents (doc 13 §6): a plaintext submit into an
          // encrypted room is refused outright — a blind server could not
          // validate it, and honest clients derive mode from their link.
          this.refuseSubmit(conn, "wrong-mode");
          return;
        }
        const before = room.session!.seq;
        const entry = room.session!.submit(msg.intent);
        // A canonical REJECTION is an edit dying, and it is the quietest way
        // that can happen: the entry is sequenced, every replica agrees, and
        // nothing diverges — which is precisely how a stale image-size bound
        // ate large photos without a single error anywhere. Counted by reason
        // so the next wrong bound shows up as a spike instead of a mystery.
        // (Encrypted rooms have no equivalent here: the server is blind, so
        // the client mirror reports it via onIntentRejected.)
        if (entry.kind === "rejected") this.obs.intentRejected(entry.reason);
        // A deduplicated re-send returns an already-sequenced entry (seq <=
        // before); persist only genuinely new entries so the log stays
        // append-once. Durability before ack (plan doc 06), then broadcast.
        if (this.storage && entry.seq > before) await this.storage.appendEntries(docId, [entry]);
        const out: ServerMessage = { t: "broadcast", entries: [entry] };
        for (const c of room.conns) c.send(out);
        this.obs.submitAccepted();
        this.obs.broadcast();
        this.noteActivity(room); // an ACCEPTED edit — the canonical activity
        return;
      }
      case "submit-enc": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) {
          this.refuseSubmit(conn, "not-joined");
          return;
        }
        if (this.expired(conn.id)) {
          this.kick(conn, "token-expired");
          return;
        }
        if (this.rateLimited(conn.id)) {
          this.refuseSubmit(conn, "rate-limit");
          return;
        }
        // Identity binding applies identically in encrypted mode — the
        // bookkeeping is plaintext precisely so the server can enforce it
        // (doc 13 §2 precondition #1).
        if (msg.envelope.clientId !== this.connClient.get(conn.id)) {
          this.refuseSubmit(conn, "client-id-mismatch");
          return;
        }
        const room = this.rooms.get(docId)!;
        // Same single predicate as the plaintext path and the roster.
        if (this.writeBlock(room, conn, msg.envelope.clientId)) {
          this.refuseSubmit(conn, "read-only");
          return;
        }
        if (!room.enc) {
          this.refuseSubmit(conn, "wrong-mode");
          return;
        }
        // Ciphertext size cap (doc 13 §2: one of the few things a blind
        // server CAN enforce). 256 KB covers any real intent incl. paste.
        if (msg.envelope.ciphertext.length > 256 * 1024) {
          this.refuseSubmit(conn, "too-large");
          return;
        }
        // Stale-base guard (pairs with quiescent checkpoints): an envelope
        // whose base precedes the adopted checkpoint would be untransformable
        // by future joiners (its context is baked-and-pruned). Refuse it —
        // the submitter has, by then, received broadcasts past the
        // checkpoint and resubmits rebased with the SAME (clientId,
        // clientSeq), so dedup semantics are unchanged.
        if (msg.envelope.base < room.enc.checkpoint.seq) {
          this.refuseSubmit(conn, "stale-base");
          return;
        }
        // Dedup by plaintext bookkeeping; re-sends return the prior entry.
        const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
        let entry: EnvelopeEntry | undefined;
        if (room.enc.seen.has(key)) {
          entry = room.enc.log.find((e) => `${e.clientId}:${e.clientSeq}` === key);
          if (!entry) {
            // PRUNED KEY: `seen` outlives the log (a checkpoint adoption bakes
            // entries into the sealed bytes and drops them), so a resubmit of an
            // already-sequenced (clientId, clientSeq) at a base >= checkpoint.seq
            // clears the stale-base guard and finds NOTHING to return. Minting a
            // fresh seq here would put the same key on the wire under TWO seqs,
            // and every honest mirror — which re-derives seqs in arrival order
            // and dedups by the same key — hard-refuses `sequence-desync`, i.e.
            // one client's malformed resend takes down the whole room. So treat
            // it as what it is: already sequenced. Idempotent no-op, no seq, no
            // fan-out. That is exactly the normal-duplicate outcome minus the
            // wire traffic: a re-broadcast of the prior entry is a dedup no-op at
            // every mirror anyway, and this entry is already baked into the
            // checkpoint every participant holds. No honest client is left
            // waiting on an echo either — claiming base >= checkpoint.seq means
            // it has confirmed everything up to and including its own pruned
            // entry, so it cannot still hold that intent pending.
            this.obs.submitRejected(); // not sequenced; no `refused` frame to tally
            return;
          }
        }
        if (!entry) {
          room.enc.seen.add(key);
          // Continue numbering from the log head, or — when the log has been
          // pruned by an adopted checkpoint (log empty, checkpoint.seq > 0) —
          // from the checkpoint seq. Resetting to 1 here was a crash bug: after
          // the first checkpoint every client's mirror expects checkpoint.seq+1,
          // so a seq of 1 fails the mirror's `entry.seq === env.seq` check and
          // hard-refuses ALL clients with `sequence-desync` (fires ~CHECKPOINT_
          // EVERY edits into any session, right after a quiescent pause).
          const seq =
            room.enc.log.length === 0
              ? room.enc.checkpoint.seq + 1
              : room.enc.log[room.enc.log.length - 1].seq + 1;
          entry = { ...msg.envelope, seq };
          room.enc.log.push(entry);
          room.enc.logBytes += msg.envelope.ciphertext.length;
        }
        const outEnc: ServerMessage = { t: "broadcast-enc", entries: [entry] };
        for (const c of room.conns) c.send(outEnc);
        this.obs.submitAccepted();
        this.obs.broadcast();
        // Encrypted rooms count activity identically: the server cannot read
        // the edit, but "an envelope was sequenced" is plaintext bookkeeping
        // and is exactly the signal the timeout needs. A blind server can
        // still tell the difference between working and abandoned.
        this.noteActivity(room);
        // PER-ROOM LOG BUDGET, checked after the broadcast so the entry that
        // crossed the line is still delivered — the room is ending either way,
        // and a half-sequenced tail helps nobody.
        //
        // This is the one hole every other limit leaves open: 256 KB envelopes
        // at the per-connection rate cap, from a room's worth of connections,
        // satisfies the frame cap, the rate cap, the IP caps and the room cap
        // simultaneously while growing this log at roughly a gigabyte per
        // second. Pruning happens only at a QUIESCENT checkpoint, which a room
        // under sustained submission never reaches — so the mechanism that
        // would bound it is precisely the one the traffic prevents.
        //
        // A sealed keystroke is ~500 bytes, so the default budget is on the
        // order of a hundred thousand real edits within a single epoch, far
        // past where a checkpoint would have pruned. Crossing it means the room
        // is not being used as a document.
        {
          const budget = this.limits.surge.roomLogMaxBytes;
          if (budget > 0 && room.enc.logBytes > budget) {
            this.endSession(docId, room, "log-overflow");
          }
        }
        return;
      }
      case "checkpoint": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) return;
        const room = this.rooms.get(docId)!;
        // Only the ASSIGNED connection's checkpoints are accepted — a
        // volunteer (or attacker) can't insist on the role (blocker 2).
        if (!room.enc || room.enc.checkpointerConnId !== conn.id) return;
        // Size cap, DERIVED from the one document limit (SurgeLimits.maxDocBytes)
        // rather than written here independently — this constant used to be its
        // own 16 MiB and could silently disagree with what the seed route let
        // in, so a document could start a session and then have every
        // checkpoint refused.
        //
        // AND IT MUST BE LOUD. This was a bare `return`: no reply, no log, no
        // counter. Checkpoints are what prune the log, so silently dropping
        // them makes the log grow until the room hits its budget and dies,
        // while every joiner replays the whole epoch — all from a size check
        // that said nothing to anybody.
        if (msg.checkpoint.ciphertext.length > maxSealedBytes(this.limits.surge.maxDocBytes)) {
          this.obs.error("hub.checkpointTooLarge", new Error("checkpoint over size cap"), {
            room: room.obsLabel,
            bytes: msg.checkpoint.ciphertext.length,
            cap: maxSealedBytes(this.limits.surge.maxDocBytes),
          });
          return;
        }
        if (msg.checkpoint.seq <= room.enc.checkpoint.seq) return; // stale
        // QUIESCENT acceptance (doc 13 / round-4 blocker 1, fix direction
        // (b)): a checkpoint is adopted only when NO retained entry sits
        // beyond it and none straddles it — i.e. it is exactly the log
        // head. Then everything ≤ its seq is baked into its bytes and can
        // prune, and future joiner tails are trivially base-complete when
        // combined with the stale-base guard on submits (below). A
        // non-quiescent checkpoint is simply ignored — the checkpointer
        // retries at the next cadence point.
        const head = room.enc.log.length === 0 ? 0 : room.enc.log[room.enc.log.length - 1].seq;
        if (msg.checkpoint.seq !== head) return;
        room.enc.checkpoint = msg.checkpoint;
        room.enc.log = [];
        room.enc.logBytes = 0; // pruned: the budget starts over with the epoch's new base
        return;
      }
      case "media-need": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) return;
        const room = this.rooms.get(docId)!;
        const blob = room.media.blobs.get(msg.sha);
        if (blob) {
          // Needing a present blob promotes it (an authenticated peer can
          // only know the sha from the referencing intent) and refreshes
          // the TTL — same semantics as an HTTP download.
          blob.staged = false;
          blob.lastDownloadAt = this.now();
          conn.send({ t: "media-ready", sha: msg.sha });
          return;
        }
        // Relay miss (doc 16 §4): register the waiter; the FIRST need for a
        // sha starts the one coalesced re-supply round (a single upload
        // serves all waiters; later needs just join the waiter set).
        const firstNeed = !room.media.waiters.has(msg.sha);
        const waiters = room.media.waiters.get(msg.sha) ?? new Set<string>();
        waiters.add(conn.id);
        room.media.waiters.set(msg.sha, waiters);
        if (firstNeed) {
          const out: ServerMessage = { t: "media-request", sha: msg.sha };
          for (const c of room.conns) if (c.id !== conn.id) c.send(out);
        }
        return;
      }
      case "media-have": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) return;
        const room = this.rooms.get(docId)!;
        for (const sha of msg.shas.slice(0, 64)) {
          // First volunteer per needed sha is chosen (rotation happens
          // naturally as sockets churn); the deadline lets sweepMedia try
          // the next volunteer if this one stalls.
          if (room.media.waiters.has(sha) && !room.media.resupply.has(sha) && !room.media.blobs.has(sha)) {
            room.media.resupply.set(sha, { chosenConnId: conn.id, deadline: this.now() + this.limits.media.uploadDeadlineMs });
            conn.send({ t: "media-upload", sha });
          }
        }
        return;
      }
      case "gossip": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) return; // like presence: ignored before join
        const room = this.rooms.get(docId)!;
        if (!room.enc) return; // gossip is an encrypted-mode channel
        // Opaque relay (doc 13 §2): the server moves the blob and learns
        // nothing — not even the hash (a content fingerprint is
        // confirmable-by-guess and therefore secret from a blind server).
        const out: ServerMessage = {
          t: "gossip",
          from: this.connClient.get(conn.id)!,
          iv: msg.iv,
          ciphertext: msg.ciphertext,
        };
        for (const c of room.conns) if (c.id !== conn.id) c.send(out);
        return;
      }
      case "ping": {
        // LIVENESS ONLY. Answered before any room lookup, auth check or rate
        // limit, and deliberately WITHOUT noteActivity.
        //
        // No room lookup / no auth: a ping asks "is this socket alive", not
        // "may I do something", and every gate added here becomes a way for a
        // pong to go missing on a connection that is in fact perfectly healthy
        // — which the client is required to read as death. A liveness reply
        // that can be refused cannot distinguish itself from the failure it
        // exists to detect, so it is unconditional. It answers a socket that
        // has not said hello too; that socket learns it is connected, which is
        // true, and it is still refused everything else.
        //
        // No noteActivity: this is the load-bearing half. The idle timeout
        // reaps rooms nobody is using, and its notes state plainly that an
        // activity-based clock is exactly what a keepalive defeats. A
        // heartbeat that counted as activity would make every abandoned tab
        // immortal and turn the idle sweep into dead letter — so a ping proves
        // the socket lives while leaving the room's idle clock running.
        //
        // No rate limit: consistent with `presence`, which is likewise
        // unmetered. The reply is O(1) and strictly smaller than the broadcast
        // fan-out this socket already provokes; a flood is bounded by the
        // per-IP connection cap and the outbound backpressure guard in ws.ts,
        // which are where socket-level abuse is actually handled.
        conn.send({ t: "pong", nonce: msg.nonce });
        return;
      }
      case "presence": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) return; // presence before join is ignored, not refused
        const room = this.rooms.get(docId)!;
        // Ephemeral: fan out to every OTHER participant, never logged/persisted.
        // `participant` is the sender's BOUND clientId (round-4 F14), not the
        // socket id: it joins presence to the identity intents carry (roster,
        // attribution) and survives the sender reconnecting on a new socket.
        // Structural clamp at the relay: presence now carries selection
        // ranges, and a hostile sender could otherwise spend every peer's
        // bandwidth on an unbounded payload (clients sanitize too, but the
        // fan-out multiplier lives HERE).
        const out: ServerMessage = {
          t: "presence",
          participant: this.connClient.get(conn.id)!,
          position:
            msg.position === null
              ? null
              // Encrypted rooms (doc 13/#20): the payload is a SEALED blob the
              // server cannot parse — and must not touch. Do not rely on the
              // sanitizer's early-return happening to pass blobs through: any
              // future hardening of it would silently corrupt sealed presence
              // ("remote carets vanish in encrypted rooms only"). The
              // client-side sanitizer runs after decryption, where the
              // coordinates actually exist.
              : room.enc
                ? msg.position
                : sanitizePresencePosition(msg.position as PresencePosition),
        };
        for (const c of room.conns) if (c.id !== conn.id) c.send(out);
        return;
      }
    }
  }

  /** Drop a connection (socket closed). Starts the room's eviction clock
   * when this was the last participant (plan doc 12 §2). */
  disconnect(conn: Connection): void {
    const docId = this.connDoc.get(conn.id);
    const clientId = this.connClient.get(conn.id); // read BEFORE the deletes below
    this.buckets.delete(conn.id);
    this.connDoc.delete(conn.id);
    this.auth.delete(conn.id);
    this.connById.delete(conn.id);
    this.connClient.delete(conn.id);
    this.ownerConns.delete(conn.id);
    if (docId) {
      const room = this.rooms.get(docId);
      if (room) {
        room.conns.delete(conn);
        // A participant's last caret is ephemeral connection state. Clear it
        // for every remaining client when that socket leaves, or the name flag
        // stays painted on the document until the participant reconnects and
        // moves it.
        if (clientId && !this.findClientConn(docId, clientId)) {
          const out: ServerMessage = { t: "presence", participant: clientId, position: null };
          for (const c of room.conns) c.send(out);
        }
        // Roster: mark disconnected (entry survives for session-lifetime
        // attribution; a reconnect under the same clientId resumes it) —
        // unless another live socket still holds this clientId (takeover
        // races: the NEW socket's binding must not be marked dead by the
        // OLD socket's teardown).
        if (clientId && !this.findClientConn(docId, clientId)) {
          const entry = room.roster.get(clientId);
          if (entry && entry.connected) {
            entry.connected = false;
            this.broadcastRoster(room);
          }
        }
        if (room.conns.size === 0) room.emptySince = this.now();
        // Rotate the checkpointer if the assigned socket just left.
        if (room.enc?.checkpointerConnId === conn.id) {
          room.enc.checkpointerConnId = undefined;
          this.assignCheckpointer(room);
        }
      }
    }
  }

  /** (Re)assign the encrypted room's checkpointer (doc 13 §3): keep the
   * incumbent while its socket lives; otherwise pick the first connection
   * (arrival order — effectively round-robin as sockets churn). */
  private assignCheckpointer(room: Room): void {
    if (!room.enc) return;
    const current = room.enc.checkpointerConnId;
    if (current && [...room.conns].some((c) => c.id === current)) return;
    const next = [...room.conns][0];
    room.enc.checkpointerConnId = next?.id;
    if (next) next.send({ t: "checkpointer", active: true });
  }

  /** Fan the full roster snapshot to everyone in the room (doc 14 §2). */
  /**
   * THE ONE PREDICATE. Returns why this connection's writes would be refused,
   * or null if they will be sequenced.
   *
   * Used by the submit paths (both modes) AND by the roster that advertises
   * write status, which is the entire point: a client told "you may write"
   * while the sequencer refuses is worse off than a client told nothing, and
   * two copies of a three-condition rule is precisely how the image-size
   * ceiling silently rotted. Enforcement and advertisement cannot disagree
   * because there is only one of them.
   *
   * ORDER IS THE ENFORCEMENT ORDER, not preference:
   *  - role first — a viewer token stays a viewer
   *    token even for the epoch's owner (the token is the capability; the
   *    owner crown does not mint write access the token withheld);
   *  - demotion next;
   *  - the room-wide lock last, for every participant. The owner keeps the
   *    admin control that can lift the lock, but the document stays read-only.
   */
  private writeBlock(room: Room, conn: Connection, clientId: string): Exclude<WriteStatus, "allowed"> | null {
    if (this.verifier && this.auth.get(conn.id)?.role !== "editor") return "viewer-role";
    if (room.demoted?.has(clientId)) return "demoted";
    if (room.readOnly) return "owner-lock";
    return null;
  }

  private broadcastRoster(room: Room): void {
    // Map each roster identity back to its live socket. Token roles belong to
    // the connection, while demotion and the room lock belong to the room.
    const connOf = new Map<string, Connection>();
    for (const c of room.conns) {
      const cid = this.connClient.get(c.id);
      if (cid) connOf.set(cid, c);
    }
    const roster = [...room.roster].map(([clientId, e]) => {
      const conn = connOf.get(clientId);
      // A disconnected participant has no token role to evaluate. The room
      // lock and an identity demotion still have exact answers.
      const write: WriteStatus = conn
        ? (this.writeBlock(room, conn, clientId) ?? "allowed")
        : room.demoted?.has(clientId)
          ? "demoted"
          : room.readOnly ? "owner-lock" : "allowed";
      return { clientId, profile: e.profile, connected: e.connected, write, participantType: e.participantType };
    });
    const out: ServerMessage = { t: "roster", roster };
    for (const c of room.conns) c.send(out);
    this.obs.rosterObserved(room.roster.size);
  }

  /** The live connection currently bound to (docId, clientId), if any —
   * the single-tab rule's lookup (doc 12 §7). Scanned, not indexed: rooms
   * are small and hellos are rare. */
  private findClientConn(docId: string, clientId: string): Connection | undefined {
    const room = this.rooms.get(docId);
    if (!room) return undefined;
    for (const c of room.conns) {
      if (this.connClient.get(c.id) === clientId) return c;
    }
    return undefined;
  }

  /**
   * UPLOAD ADMISSION — answered BEFORE the request body is read.
   *
   * This is the half of upload control that `mediaUpload` structurally cannot
   * do. That method judges bytes, so calling it requires having received the
   * bytes: an upload to a room that does not exist still costs a full body in
   * memory before it earns its 404, and N concurrent such requests cost N
   * bodies. Unauthenticated, unpaced, and on the one route that accepts
   * megabytes, that is a memory amplifier rather than a rate limit.
   *
   * So admission asks the two questions that need no body at all:
   *
   *   1. Is there an OPEN ROOM to upload to? Rooms are deleted when they end
   *      (see endRoom), so presence in `this.rooms` IS openness — there is no
   *      separate closed state to miss.
   *   2. Has this room spent its upload budget? Per room, because a room is
   *      the unit a document and its participants share; a per-IP bucket
   *      would punish an office behind one address for collaborating, and a
   *      global one would let any single room starve every other.
   *
   * THE TOKEN IS SPENT HERE, not on completion, and that ordering is the
   * whole point. Charging on completion lets an attacker open a thousand
   * concurrent uploads that every one pass admission before any of them
   * finishes — the limit would be perfectly enforced against a caller
   * polite enough to wait its turn, and no defence at all against the
   * traffic it exists to stop.
   *
   * The cost of spending early is that a refused upload (413/507/400) still
   * consumed a token and is not refunded. That is deliberate: a client whose
   * uploads keep failing is either broken or hostile, and neither should get
   * an unmetered retry loop. Legitimate clients pre-check `mediaMaxBlobBytes`
   * from the welcome and do not send bodies they know are oversized.
   *
   * CONSIDERED AND REJECTED: exempting SOLICITED uploads — a sha the room is
   * actively waiting for or has an open re-supply round on — since the server
   * asked for those itself. It reads well and reopens the hole: `media-need`
   * is what creates a waiter, so anyone in the room could name arbitrary shas
   * to mint exemptions and then PUT unmetered bodies against them. The sha
   * mismatch that would reject those bodies is checked in `mediaUpload`,
   * AFTER buffering, so the exemption would buy back the exact amplifier this
   * function exists to close. The burst is sized for re-supply instead, which
   * gets the same outcome without a caller-controlled bypass.
   */
  mediaUploadAdmission(docId: string): number {
    const room = this.rooms.get(docId);
    if (!room) return 404;
    const perMin = this.limits.media.roomUploadsPerMin;
    if (perMin <= 0) return 200; // 0 = disabled, matching the per-IP convention
    const burst = this.limits.media.roomUploadBurst;
    const b = room.media.uploadBucket;
    const now = this.now();
    b.tokens = Math.min(burst, b.tokens + ((now - b.last) / 60_000) * perMin);
    b.last = now;
    if (b.tokens < 1) return 429;
    b.tokens -= 1;
    return 200;
  }

  /**
   * Media upload (doc 16 §4, the PUT handler's core). The server's ONLY
   * content-touching operation is sha256: the address IS the verification
   * (doc 16 §1.1) — a body that doesn't hash to `sha` is rejected no
   * matter who sent it, which is the entire swap-proofing story, and it
   * works identically for E2EE ciphertext blobs. New blobs enter STAGED
   * (upload-then-intent, doc 05); promotion happens when the plaintext
   * session applies the referencing intent, or on first peer download in
   * encrypted rooms (the download IS the observable claim).
   */
  async mediaUpload(docId: string, sha: string, bytes: Uint8Array): Promise<number> {
    const room = this.rooms.get(docId);
    if (!room) return 404;
    if (bytes.length > this.limits.media.maxBlobBytes) return 413;
    if (room.media.blobs.has(sha)) return 200; // content-addressed dedup
    if (room.media.totalBytes + bytes.length > this.limits.media.maxRoomBytes) return 507;
    const digest = await crypto.subtle.digest("SHA-256", bytes as unknown as ArrayBuffer);
    let hex = "";
    for (const b of new Uint8Array(digest)) hex += b.toString(16).padStart(2, "0");
    if (hex !== sha) return 400; // the sha check — never trust the label
    // RAM-PRESSURE LADDER (doc 16 §4), only when the spill tier is on: if
    // admitting these bytes would push global media RAM past the budget,
    // demote LRU victims to disk (and beyond the disk budget, evict
    // disk-LRU entirely) until it fits. 507 only when RAM cannot be freed
    // at all — every remaining RAM blob has waiters.
    if (this.spill && this.limits.media.ramBytes > 0 && this.mediaRamBytes + bytes.length > this.limits.media.ramBytes) {
      const fits = await this.demoteForRam(bytes.length);
      if (!fits) {
        this.noteMediaTier(); // the failed ladder may still have demoted some victims
        return 507;
      }
    }
    // RE-VALIDATE after the awaits above (digest + ladder): the room may have
    // evicted underneath us, and a concurrent upload of the same sha may have
    // landed — inserting anyway would strand bytes in a dead room's map or
    // double-count the blob, which is exactly the accounting drift the
    // gauges must never have.
    if (this.rooms.get(docId) !== room) return 404;
    if (room.media.blobs.has(sha)) return 200;
    room.media.blobs.set(sha, { loc: "ram", bytes, lastDownloadAt: this.now(), staged: true });
    room.media.totalBytes += bytes.length;
    this.mediaRamBytes += bytes.length;
    // A re-supply round for this sha completes here: serve every waiter.
    const waiters = room.media.waiters.get(sha);
    if (waiters) {
      for (const cid of waiters) this.connById.get(cid)?.send({ t: "media-ready", sha });
      room.media.waiters.delete(sha);
      room.media.resupply.delete(sha);
      room.media.blobs.get(sha)!.staged = false; // re-supplied ⇒ referenced
    }
    this.obs.mediaUp();
    this.noteMediaTier();
    // Media transfers are work in the room even when no intent is sequenced
    // (a paste of a large image is one submit and many seconds of upload) —
    // a session must not idle out mid-transfer.
    this.noteActivity(room);
    return 201;
  }

  /** Media download (the GET handler's core). A peer download PROMOTES a
   * staged blob (doc 16 §4: in encrypted rooms the server can't read the
   * referencing intent, but only someone who saw it can know the sha) and
   * refreshes the TTL — in BOTH tiers, identically. A RAM hit returns the
   * bytes (unchanged); a disk hit returns a lazy decrypting stream the
   * transport pipes, so a spilled blob is never pulled back into RAM
   * (one-way ladder). If the file vanishes between this index hit and
   * `open()` (a sweep unlink racing the request), `open()` rejects and the
   * transport answers absent — the same outcome as arriving one tick later. */
  mediaDownload(docId: string, sha: string): MediaDownloadHit | null {
    const room = this.rooms.get(docId);
    const blob = room?.media.blobs.get(sha);
    if (!room || !blob) return null;
    blob.staged = false;
    blob.lastDownloadAt = this.now();
    this.obs.mediaDown();
    this.noteActivity(room);
    if (blob.loc === "ram") return blob.bytes;
    const spill = this.spill;
    const spillId = room.spillId;
    if (!spill || spillId === undefined) return null; // disk entry with no spill: cannot happen while wired
    // Counted at the index hit, not the open(): a sweep unlink racing the
    // request makes the open reject, but the serve was attempted — the rate
    // is the signal either way.
    this.obs.spillRead();
    return { len: blob.len, open: () => spill.openRead(spillId, sha) };
  }

  /** Current tier gauges — RAM-resident media bytes and spill disk bytes
   * (on-disk sizes, IV included). For tests (the accounting-drift invariant)
   * and ops; contains no content and no identifiers. */
  mediaTierBytes(): { ram: number; disk: number } {
    return { ram: this.mediaRamBytes, disk: this.spill?.totalBytes() ?? 0 };
  }

  /** Spilled files on disk right now — the `mediaSpillFiles` gauge. */
  mediaSpillFiles(): number {
    return this.spill?.fileCount() ?? 0;
  }

  /** Push the tier gauges to the observability sink. Called after every
   * operation that can move bytes between tiers (upload/ladder, sweep, room
   * eviction), so the sink's last-observed gauges track the ladder without
   * the sink ever holding a reference to the hub. */
  private noteMediaTier(): void {
    this.obs.mediaTierObserved(this.mediaRamBytes, this.spill?.totalBytes() ?? 0, this.spill?.fileCount() ?? 0);
  }

  /**
   * Demote RAM-LRU victims to disk until `incomingLen` more bytes fit under
   * the RAM budget. Returns false when no demotable victim remains — every
   * RAM blob has waiters or is already mid-demotion — which is the ONLY
   * situation an upload answers 507 for (doc 16 §4: full-ladder eviction
   * couldn't make room).
   */
  private async demoteForRam(incomingLen: number): Promise<boolean> {
    const budget = this.limits.media.ramBytes;
    while (this.mediaRamBytes + incomingLen > budget) {
      const victim = this.pickRamVictim();
      if (!victim) return false;
      await this.demote(victim.docId, victim.room, victim.sha, victim.entry);
    }
    return true;
  }

  /**
   * Choose the next demotion victim, globally across rooms (the budget is a
   * process budget). Order per the ladder spec: promoted-LRU by
   * lastDownloadAt first, then staged-LRU. Never a blob with waiters, never
   * one already mid-demotion. A blob whose body is still being received or
   * hashed is structurally exempt: it is not in any room's map until
   * `mediaUpload` finishes hashing it.
   */
  private pickRamVictim(): { docId: string; room: Room; sha: string; entry: Extract<MediaBlobEntry, { loc: "ram" }> } | null {
    type Victim = { docId: string; room: Room; sha: string; entry: Extract<MediaBlobEntry, { loc: "ram" }> } | null;
    let promoted: Victim = null;
    let staged: Victim = null;
    for (const [docId, room] of this.rooms) {
      for (const [sha, entry] of room.media.blobs) {
        if (entry.loc !== "ram" || entry.demoting || room.media.waiters.has(sha)) continue;
        const slot = entry.staged ? staged : promoted;
        if (!slot || entry.lastDownloadAt < slot.entry.lastDownloadAt) {
          if (entry.staged) staged = { docId, room, sha, entry };
          else promoted = { docId, room, sha, entry };
        }
      }
    }
    return promoted ?? staged;
  }

  /**
   * Demote one RAM blob to the spill tier. The ONE-WAY LADDER's middle rung:
   *
   *  1. Make disk room first — beyond `diskBytes`, evict disk-LRU entries
   *     ENTIRELY (gone; re-supply recovers them — the locker is a cache).
   *  2. Stream the bytes to disk (tmp-then-rename inside the spill). The RAM
   *     entry stays in the index UNTOUCHED throughout, so a GET racing the
   *     write serves from the buffer; only after the rename does the index
   *     swap to a disk entry and the buffer get released.
   *  3. If disk cannot take it — budget even after full disk eviction, or a
   *     real ENOSPC/EIO — the candidate is DROPPED ENTIRELY and counted
   *     (obs.error). Degrading loses a cache entry re-supply recovers; it
   *     must never surface as a failed upload.
   *
   * The swap copies `lastDownloadAt`/`staged` from the LIVE entry, so a GET
   * that promoted or refreshed the blob mid-write carries into the disk
   * entry. If the entry vanished mid-write (TTL sweep, room eviction), the
   * freshly written file is an orphan and is removed — whoever removed the
   * entry already released its RAM bytes.
   */
  private async demote(docId: string, room: Room, sha: string, entry: Extract<MediaBlobEntry, { loc: "ram" }>): Promise<void> {
    entry.demoting = true;
    const len = entry.bytes.length;
    const spill = this.spill!;
    const spillId = (room.spillId ??= newRoomSpillId());
    if (this.evictDiskFor(len + SPILL_FILE_OVERHEAD)) {
      try {
        await spill.write(spillId, sha, entry.bytes);
        this.obs.spillWrite();
        const live = this.rooms.get(docId) === room && room.media.blobs.get(sha) === entry;
        if (live) {
          room.media.blobs.set(sha, { loc: "disk", len, lastDownloadAt: entry.lastDownloadAt, staged: entry.staged });
          this.mediaRamBytes -= len;
        } else if (this.rooms.get(docId) !== room) {
          this.spillRm(spillId); // room died mid-write: the write may have recreated its directory
        } else {
          this.spillRm(spillId, sha); // entry swept mid-write: the file alone is the orphan
        }
        return;
      } catch (err) {
        // ENOSPC/EIO on the spill write: fall through to the drop below.
        this.obs.error("hub.spillWrite", err);
      }
    }
    // Disk cannot take it: RAM → gone, skipping the middle rung. Counted via
    // the error above when a write failed; a budget-impossible victim (bigger
    // than the whole disk tier) needs no line — the ladder simply ends.
    if (this.rooms.get(docId) === room && room.media.blobs.get(sha) === entry) {
      room.media.blobs.delete(sha);
      room.media.totalBytes -= len;
      this.mediaRamBytes -= len;
    }
  }

  /** Evict disk-LRU entries (never with waiters) until `bytesNeeded` more
   * fits under the disk budget. Eviction is total: index entry gone, room
   * budget released, file removed. Returns false when it cannot fit even
   * with the tier emptied of every evictable entry. */
  private evictDiskFor(bytesNeeded: number): boolean {
    const budget = this.limits.media.diskBytes;
    const spill = this.spill!;
    while (spill.totalBytes() + bytesNeeded > budget) {
      let lru: { room: Room; sha: string; entry: Extract<MediaBlobEntry, { loc: "disk" }> } | null = null;
      for (const room of this.rooms.values()) {
        for (const [sha, entry] of room.media.blobs) {
          if (entry.loc !== "disk" || room.media.waiters.has(sha)) continue;
          if (!lru || entry.lastDownloadAt < lru.entry.lastDownloadAt) lru = { room, sha, entry };
        }
      }
      if (!lru) return false;
      lru.room.media.blobs.delete(lru.sha);
      lru.room.media.totalBytes -= lru.entry.len;
      this.obs.spillEviction();
      // The spill releases its byte accounting SYNCHRONOUSLY (see
      // MediaSpill.unlink), so the loop's totalBytes() re-read makes progress
      // even though the rm itself is still in flight.
      this.spillRm(lru.room.spillId!, lru.sha);
    }
    return true;
  }

  /**
   * Remove a spilled file — or a room's whole spill directory when `sha` is
   * omitted — without making the caller wait on the filesystem. On failure:
   * log + count (obs.error), then park the path on the retry list swept
   * alongside sweepMedia. The spill's byte accounting was already released
   * in the call's synchronous prefix, so a stuck file is a DISK-SPACE bug,
   * never a retention bug (per-boot key + boot wipe still cover it).
   */
  private spillRm(roomSpillId: string, sha?: string): void {
    if (!this.spill) return;
    const op = sha ? this.spill.unlink(roomSpillId, sha) : this.spill.removeRoom(roomSpillId);
    op.catch((err) => {
      this.obs.error(sha ? "hub.spillUnlink" : "hub.spillRemoveRoom", err);
      // The disk-space-leak signal: the retry list grew. A re-parked retry
      // counts again, so a NON-DRAINING list shows as a climbing rate.
      this.obs.spillUnlinkFailure();
      this.spillRetry.push({ roomSpillId, sha });
    });
  }

  /** Media sweep (doc 16 §4): staged-unreferenced past T_STAGE and
   * promoted-idle past T_MEDIA evict (never with waiters — a waiter
   * blocks eviction); stalled re-supply rounds re-broadcast the request
   * so another holder can volunteer. Call on the same timer as
   * sweepRooms. */
  sweepMedia(): void {
    const now = this.now();
    // Retry rms that failed earlier (see spillRm). Re-attempted here — the
    // same cadence as everything else time-based — and a still-failing rm
    // re-parks itself, so nothing is dropped and nothing spins.
    if (this.spillRetry.length > 0) {
      const retry = this.spillRetry;
      this.spillRetry = [];
      for (const r of retry) this.spillRm(r.roomSpillId, r.sha);
    }
    for (const room of this.rooms.values()) {
      for (const [sha, blob] of [...room.media.blobs]) {
        if (room.media.waiters.has(sha)) continue;
        const ttl = blob.staged ? this.limits.media.stageTtlMs : this.limits.media.ttlMs;
        if (now - blob.lastDownloadAt >= ttl) {
          room.media.blobs.delete(sha);
          room.media.totalBytes -= mediaBlobLen(blob);
          // Tier bookkeeping: a RAM eviction releases the global gauge; a
          // disk eviction also removes its file (TTL'd bytes must not sit on
          // disk until room death).
          if (blob.loc === "ram") this.mediaRamBytes -= blob.bytes.length;
          else if (room.spillId !== undefined) this.spillRm(room.spillId, sha);
        }
      }
      for (const [sha, r] of [...room.media.resupply]) {
        if (now >= r.deadline) {
          room.media.resupply.delete(sha);
          if (room.media.waiters.has(sha)) {
            let anyPeer = false;
            for (const c of room.conns) {
              if (!room.media.waiters.get(sha)!.has(c.id)) {
                anyPeer = true;
                c.send({ t: "media-request", sha });
              }
            }
            if (!anyPeer) {
              // Nobody left to ask (doc 16 §7 honest failure mode): tell the
              // waiters; the registration stays so a joining holder revives
              // it via welcome.mediaNeeded.
              for (const cid of room.media.waiters.get(sha)!) {
                this.connById.get(cid)?.send({ t: "media-unavailable", sha });
              }
            }
          }
        }
      }
    }
    this.noteMediaTier();
  }

  /**
   * Zero-custody eviction (plan doc 12 §2): delete every room that has been
   * empty for at least the grace period. Deletion is total — document, log,
   * dedup state, undo stacks all go with the Room object; with no storage
   * driver the server then holds NOTHING for that docId (the browsers are
   * the recovery machinery). Deliberately a method the transport calls on a
   * timer (like sweepExpired) rather than a self-armed setTimeout: the hub
   * stays timer-free and eviction is deterministic in tests via the
   * injected clock. Only EMPTY rooms are eligible — an idle room with
   * people connected is never deleted out from under them (round-4 F7).
   * Returns the evicted docIds (for logging/metrics).
   */
  /**
   * PER-ROOM RESOURCE SNAPSHOT for the ops view: which rooms are consuming
   * what, and how close each is to the limit that will end it.
   *
   * This is shapes and counts, exactly like the aggregate metrics — the extra
   * thing it does is ATTRIBUTE them, so an outlier room is visible instead of
   * being averaged into a global gauge that looks fine. "Media is at 60% of
   * the server" is not actionable; "one room is at 98% of its own cap" is.
   *
   * The zero-custody rule is unchanged and is why `obsLabel` exists: no docId,
   * no clientId, no share code, no owner token, no participant name, no IP,
   * and nothing derived from any of them. Every value below is a number, a
   * boolean, or the fixed word "plaintext"/"encrypted".
   *
   * Each measure is paired with the limit it runs against, so a consumer never
   * has to keep a second copy of the configuration in step with this one — the
   * bug that shipped once already in the compose file.
   */
  roomsSnapshot(): RoomSnapshot[] {
    const now = this.now();
    const out: RoomSnapshot[] = [];
    for (const room of this.rooms.values()) {
      let connected = 0;
      for (const entry of room.roster.values()) if (entry.connected) connected++;
      // Split the room's media figure by tier. Blob lengths in both tiers
      // (the quota's basis), summing to the old single `mediaBytes`.
      let mediaRam = 0;
      let mediaDisk = 0;
      for (const b of room.media.blobs.values()) {
        if (b.loc === "ram") mediaRam += b.bytes.length;
        else mediaDisk += b.len;
      }
      out.push({
        room: room.obsLabel,
        mode: room.enc ? "encrypted" : "plaintext",
        ageMs: now - room.createdAt,
        ageLimitMs: this.limits.lifecycle.roomMaxLifetimeMs,
        idleMs: now - room.lastActivityAt,
        idleLimitMs: this.limits.lifecycle.idleTimeoutMs,
        // null while anyone is connected — the eviction clock is not running.
        emptyMs: room.emptySince === undefined ? null : now - room.emptySince,
        emptyLimitMs: this.limits.lifecycle.emptyRoomTtlMs,
        conns: room.conns.size,
        rosterTotal: room.roster.size,
        rosterConnected: connected,
        connLimit: CollabHub.MAX_ROOM_CLIENTS,
        mediaRamBytes: mediaRam,
        mediaDiskBytes: mediaDisk,
        mediaLimitBytes: this.limits.media.maxRoomBytes,
        mediaBlobs: room.media.blobs.size,
        mediaWaiting: room.media.waiters.size,
        // Budget REMAINING, floored: a room being throttled reads 0 here,
        // which is the signal worth seeing.
        uploadTokens: Math.max(0, Math.floor(room.media.uploadBucket.tokens)),
        uploadBurst: this.limits.media.roomUploadBurst,
        logBytes: room.enc?.logBytes ?? 0,
        logLimitBytes: this.limits.surge.roomLogMaxBytes,
        readOnly: room.readOnly === true,
        idleWarned: room.idleWarned === true,
        lifetimeWarned: room.lifetimeWarned === true,
      });
    }
    return out;
  }

  sweepRooms(graceMs: number = this.limits.lifecycle.emptyRoomTtlMs): string[] {
    const evicted: string[] = [];
    if (graceMs <= 0) return evicted; // 0 disables (limits.ts convention)
    for (const [docId, room] of [...this.rooms]) {
      if (room.conns.size === 0 && room.emptySince !== undefined && this.now() - room.emptySince >= graceMs) {
        this.evictRoom(docId, "empty");
        evicted.push(docId);
      }
    }
    return evicted;
  }

  /**
   * Mark QUALIFYING ACTIVITY in a room and cancel any idle countdown.
   *
   * The definition of "qualifying" is the load-bearing decision (see
   * LIFECYCLE_NOTES.activity in limits.ts): accepted submits, owner admin
   * actions, media transfers, and joins — deliberately NOT presence. A
   * forgotten tab emits presence whenever the mouse moves or the window
   * regains focus, so counting it would let any abandoned laptop hold a room
   * open forever and the idle timeout would be dead letter.
   *
   * Cancelling is announced: a client that was shown a countdown must be told
   * it is over, or the UI sits at "ending in 12s" while the session is
   * perfectly healthy.
   */
  private noteActivity(room: Room): void {
    room.lastActivityAt = this.now();
    if (!room.idleWarned) return;
    room.idleWarned = false;
    const out: ServerMessage = { t: "session-warning-cleared", reason: "idle" };
    for (const c of room.conns) c.send(out);
  }

  /** Delete a room and everything in it, releasing the creator's per-IP doc
   * slot. The ONE place a room leaves `this.rooms`, so the eviction counter,
   * the reason tally, the IP release — and now the media tier bookkeeping —
   * can never drift apart from the deletion itself: the RAM gauge gives back
   * every resident blob, and the room's spill directory is removed (bytes
   * subtracted and index dropped synchronously inside spillRm; a failed rm
   * lands on the retry list — a disk-space bug, never a retention bug). */
  private evictRoom(docId: string, reason: EvictionReason): void {
    const room = this.rooms.get(docId);
    this.rooms.delete(docId);
    if (room) {
      for (const blob of room.media.blobs.values()) {
        if (blob.loc === "ram") this.mediaRamBytes -= blob.bytes.length;
      }
      if (room.spillId !== undefined) this.spillRm(room.spillId);
      this.noteMediaTier();
    }
    this.ipGuard.releaseRoom(docId);
    this.obs.roomEvicted(1, reason);
  }

  /**
   * Sweep the TIME-BASED session deadlines: the idle timeout and the absolute
   * lifetime cap, each with its warning. Called on the same timer as
   * {@link sweepRooms} (the hub stays timer-free — deadlines are deterministic
   * in tests via the injected clock).
   *
   * WARNING TIMING: a warning is sent when the remaining time first drops
   * below the configured lead, which — because this runs on a periodic sweep
   * — happens up to one sweep interval EARLY. The frame therefore carries the
   * measured remainder rather than the configured lead, so a client's
   * countdown matches when the session actually ends.
   *
   * ORDER MATTERS: lifetime is checked before idle, so a room that has hit
   * both deadlines in the same sweep dies of the one that cannot be argued
   * with, and the reason an operator sees is the true cause.
   *
   * Returns what it ended, for logging/metrics and for tests to assert on.
   */
  sweepLifecycle(): { docId: string; reason: EvictionReason }[] {
    const now = this.now();
    const { idleTimeoutMs, idleWarningMs, roomMaxLifetimeMs, lifetimeWarningMs } = this.limits.lifecycle;
    const ended: { docId: string; reason: EvictionReason }[] = [];
    for (const [docId, room] of [...this.rooms]) {
      // ── absolute lifetime cap ────────────────────────────────────────
      if (roomMaxLifetimeMs > 0) {
        const remaining = room.createdAt + roomMaxLifetimeMs - now;
        if (remaining <= 0) {
          this.endSession(docId, room, "session-expired");
          ended.push({ docId, reason: "session-expired" });
          continue;
        }
        if (!room.lifetimeWarned && lifetimeWarningMs > 0 && remaining <= lifetimeWarningMs) {
          room.lifetimeWarned = true;
          this.warn(room, "lifetime", remaining);
        }
      }
      // ── idle timeout ─────────────────────────────────────────────────
      if (idleTimeoutMs > 0) {
        const remaining = room.lastActivityAt + idleTimeoutMs - now;
        if (remaining <= 0) {
          this.endSession(docId, room, "idle-timeout");
          ended.push({ docId, reason: "idle-timeout" });
          continue;
        }
        if (!room.idleWarned && idleWarningMs > 0 && remaining <= idleWarningMs) {
          room.idleWarned = true;
          this.warn(room, "idle", remaining);
        }
      }
    }
    return ended;
  }

  /**
   * Bring a just-joined connection up to date on an ALREADY-ANNOUNCED
   * deadline, right after its welcome. Only the lifetime cap can be in this
   * state: a join is activity, so it has just cleared any idle warning.
   *
   * Without this, someone joining inside the final five minutes sees no
   * countdown at all and the session simply ends on them — the warning is a
   * room property, and a client that missed the broadcast still deserves it.
   */
  private sendPendingWarning(room: Room, conn: Connection): void {
    if (!room.lifetimeWarned || this.limits.lifecycle.roomMaxLifetimeMs <= 0) return;
    const remaining = room.createdAt + this.limits.lifecycle.roomMaxLifetimeMs - this.now();
    if (remaining > 0) conn.send({ t: "session-warning", reason: "lifetime", inMs: remaining });
  }

  /** Announce an approaching deadline to everyone in the room. */
  private warn(room: Room, reason: "idle" | "lifetime", remainingMs: number): void {
    const out: ServerMessage = { t: "session-warning", reason, inMs: Math.max(0, remainingMs) };
    for (const c of room.conns) c.send(out);
    this.obs.sessionWarned(reason);
  }

  /**
   * End a live session: every participant is kicked with a distinct reason,
   * then the room is deleted outright.
   *
   * The room is NOT left to the empty-room grace afterwards. That grace
   * exists so a refresh or a flaky reconnect can resume a session that is
   * still meant to exist; a session the server has just decided to END is a
   * different thing, and leaving a 60-second window in which a reconnecting
   * client rejoins a room already declared dead would make the timeout
   * advisory rather than enforced.
   *
   * Iterating a COPY of `conns` is required, not stylistic: `kick` calls
   * `disconnect`, which mutates the set being walked.
   */
  private endSession(
    docId: string,
    room: Room,
    reason: Extract<EvictionReason, "idle-timeout" | "session-expired" | "log-overflow">,
  ): void {
    for (const conn of [...room.conns]) this.kick(conn, reason);
    this.evictRoom(docId, reason);
  }

  /** Sessions currently held in memory (for eviction/metrics). */
  activeDocs(): string[] {
    return [...this.rooms.keys()];
  }

  /**
   * The GLOBAL room ceiling. Per-IP caps bound one address; they do not bound
   * a thousand cooperating addresses, and a botnet is exactly that — every
   * member individually impeccable.
   *
   * A FAIL-FAST VALVE, and the tradeoff is deliberate: briefly refusing NEW
   * sessions is a far better failure than an out-of-memory kill that takes
   * every EXISTING session down with it. Rooms already live are untouched, and
   * capacity returns on its own as they evict.
   */
  private atCapacity(): { ok: false; reason: "server-full" } | null {
    const max = this.limits.surge.maxRooms;
    if (max > 0 && this.rooms.size >= max) {
      this.obs.capacityRefused("server-full");
      return { ok: false, reason: "server-full" };
    }
    return null;
  }

  /**
   * Seed (or re-seed) a session from client-supplied bytes — the transport-
   * free core of `POST/PUT /docs` (plan doc 12 §3/§5.3). The bytes come from
   * a participant's bundle: under zero custody the browsers are the durable
   * copies, so "bring it back live" means any holder re-uploads theirs.
   *
   * First-wins atomically (single-threaded create-if-missing, per process —
   * doc 12 §5 notes a multi-node demo needs a real lock): if the room already
   * exists the seed is REFUSED and the caller gets the incumbent epoch's
   * genesisId — the HTTP layer maps that to 409 and the losing client lands
   * in resume case 2 (join the winner's session; keep its own copy as a
   * draft/lineage decision, never a silent merge).
   *
   * The sidecar is part of the seed bundle (round-4 F10): a bundle's docx
   * with split-created carried ids cannot re-derive its id table from parse
   * order. Callers seeding a FRESH doc (no history) may omit it.
   *
   * A freshly seeded room starts EMPTY with the eviction clock running — a
   * seed nobody joins evicts after the grace like any abandoned room, so an
   * unauthenticated seed endpoint cannot accrete unbounded RAM.
   */
  seed(
    docId: string,
    docx: Uint8Array,
    sidecar?: IdSidecar,
    codeVerifier?: string,
    lineage?: LineageHead[],
    /** Normalized creator address for the per-IP limits (see ip-guard.ts).
     * Omitted by callers that have no transport (tests, embedders), which
     * simply skips the per-IP accounting. */
    creatorIp?: string,
  ): SeedResult {
    const existing = this.rooms.get(docId);
    if (existing) return { ok: false, reason: "exists", genesisId: existing.genesisId };
    const full = this.atCapacity();
    if (full) return full;
    // Per-IP gate BEFORE the parse: seeding is the expensive unauthenticated
    // operation (it decodes a docx and allocates a room), so the limit has to
    // land in front of the work, not after it.
    const allowed = this.ipGuard.allowSeed(creatorIp ?? "");
    if (!allowed.ok) {
      this.obs.ipLimited(allowed.reason);
      return { ok: false, reason: allowed.reason };
    }
    const session = new DocumentSession(DocxDocument.load(docx));
    if (sidecar) session.installSidecar(sidecar);
    const room: Room = {
      session,
      enc: null,
      conns: new Set(),
      genesisId: this.genGenesisId(),
      roster: new Map(),
      agentInvites: new Map(),
      agentClients: new Map(),
      codeVerifier,
      lineage,
      ownerToken: defaultGenesisId().replace("g_", "o_"),
      emptySince: this.now(), // eviction clock runs until someone joins
      createdAt: this.now(),
      lastActivityAt: this.now(),
      creatorIp,
      obsLabel: newObsLabel(),
      media: emptyMedia(this.limits.media.roomUploadBurst, this.now()),
    };
    this.rooms.set(docId, room);
    if (creatorIp) this.ipGuard.noteRoom(creatorIp, docId);
    this.obs.roomCreated();
    return { ok: true, genesisId: room.genesisId, ownerToken: room.ownerToken! };
  }

  /**
   * Seed an ENCRYPTED session (doc 13): the body is the seeder's SEALED
   * checkpoint — the server stores it opaquely and can validate nothing but
   * size (content validation happens on every client at decrypt, doc 11
   * E2EE amendment). The seeder supplies the genesisId it sealed under:
   * epoch keys derive from (K_doc, genesisId), so the id must exist before
   * sealing — client-minted for encrypted epochs, server-minted for
   * plaintext ones. A hostile key-holder gaining anything from reusing an
   * old epoch id would need powers (content injection) the key already
   * grants legitimately; honest seeders mint fresh 128-bit ids.
   * First-wins semantics identical to plaintext seed().
   */
  seedEncrypted(
    docId: string,
    genesisId: string,
    checkpoint: SealedCheckpoint,
    codeVerifier?: string,
    creatorIp?: string,
  ): SeedResult {
    const existing = this.rooms.get(docId);
    if (existing) return { ok: false, reason: "exists", genesisId: existing.genesisId };
    const full = this.atCapacity();
    if (full) return full;
    const allowed = this.ipGuard.allowSeed(creatorIp ?? "");
    if (!allowed.ok) {
      this.obs.ipLimited(allowed.reason);
      return { ok: false, reason: allowed.reason };
    }
    const room: Room = {
      session: null,
      enc: { checkpoint, log: [], seen: new Set(), logBytes: 0 },
      conns: new Set(),
      genesisId,
      roster: new Map(),
      agentInvites: new Map(),
      agentClients: new Map(),
      codeVerifier,
      ownerToken: defaultGenesisId().replace("g_", "o_"),
      emptySince: this.now(),
      createdAt: this.now(),
      lastActivityAt: this.now(),
      creatorIp,
      obsLabel: newObsLabel(),
      media: emptyMedia(this.limits.media.roomUploadBurst, this.now()),
    };
    this.rooms.set(docId, room);
    if (creatorIp) this.ipGuard.noteRoom(creatorIp, docId);
    this.obs.roomCreated();
    return { ok: true, genesisId, ownerToken: room.ownerToken! };
  }

  private async room(docId: string): Promise<Room> {
    let room = this.rooms.get(docId);
    if (!room) {
      const session = await this.startSession(docId);
      room = {
        session,
        enc: null,
        conns: new Set(),
        genesisId: this.genGenesisId(),
        roster: new Map(),
        agentInvites: new Map(),
        agentClients: new Map(),
        createdAt: this.now(),
        lastActivityAt: this.now(),
        obsLabel: newObsLabel(),
      media: emptyMedia(this.limits.media.roomUploadBurst, this.now()),
      };
      this.rooms.set(docId, room);
      this.obs.roomCreated();
    }
    return room;
  }

  private async startSession(docId: string): Promise<DocumentSession> {
    // Reached only from room(), which the hello path gates: with no provider
    // an unknown docId was already refused `no-session` before this point.
    if (!this.provider) throw new Error("no provider (zero-custody hub) — rooms exist only via seed()");
    if (this.storage) {
      const snap = await this.storage.loadSnapshot(docId);
      if (snap) {
        const session = new DocumentSession(DocxDocument.load(snap.docx));
        // Restore the snapshot's exact id table BEFORE replaying the tail —
        // tail entries address nodes by the ids the session had when they
        // were sequenced, which parse order alone cannot reproduce once
        // splits happened pre-snapshot (round-2 F1 / round-4 F10).
        session.installSidecar(snap.sidecar);
        session.loadCanonical(await this.storage.readLog(docId, snap.seq));
        return session;
      }
      // No snapshot yet: start fresh but replay any persisted log tail.
      const session = new DocumentSession(DocxDocument.load(this.provider.load(docId)));
      session.loadCanonical(await this.storage.readLog(docId, 0));
      return session;
    }
    return new DocumentSession(DocxDocument.load(this.provider.load(docId)));
  }
}

/**
 * Media relay limits (doc 16 §8; RAM tier) — DEPRECATED as a source of truth.
 *
 * These are now configuration (`limits.media`, parsed in limits.ts with the
 * rest of the table) because the right blob size is a deployment decision: a
 * public demo wants a few megabytes, a local test wants room to paste a camera
 * image without thinking about it. The hub reads `this.limits.media`; this
 * object survives only as the legacy shape for existing importers and mirrors
 * the defaults. New code should take a `MediaLimits`.
 */
export const MEDIA_LIMITS = {
  maxBlobBytes: DEFAULT_LIMITS.media.maxBlobBytes,
  roomMediaBytes: DEFAULT_LIMITS.media.maxRoomBytes,
  tMediaMs: DEFAULT_LIMITS.media.ttlMs,
  tStageMs: DEFAULT_LIMITS.media.stageTtlMs,
  tUploadMs: DEFAULT_LIMITS.media.uploadDeadlineMs,
};

/** Fresh opaque room label for the ops view — see `Room.obsLabel`. */
function newObsLabel(): string {
  return `r_${crypto.randomUUID().replace(/-/g, "").slice(0, 10)}`;
}

function emptyMedia(uploadBurst: number, now: number): Room["media"] {
  // A new room starts with a FULL bucket: the first thing a real room does is
  // receive the images of the document that was just seeded into it, and
  // making that first burst wait for tokens to accrue would delay the one
  // upload pattern that is always legitimate.
  return { blobs: new Map(), waiters: new Map(), resupply: new Map(), totalBytes: 0, uploadBucket: { tokens: uploadBurst, last: now } };
}

/** Fixed presence palette (doc 11 XSS vector 7: colors are validated
 * against a palette or a strict hex shape — never arbitrary CSS). */
const PROFILE_PALETTE = ["#e05252", "#e0a952", "#7fbf5a", "#52b3e0", "#7d6ee0", "#d05fb8", "#4fc2a2", "#c2b04f"];

/**
 * Server-side profile sanitization (doc 14 §2, plaintext mode): name
 * trimmed, control characters stripped, 1–40 chars (empty → generated
 * default); color must be exactly `#rrggbb` or is replaced by a palette
 * color hashed from the clientId (stable across rejoins). Rendering clients
 * sanitize AGAIN (text-node-only, palette check) — this is defense in
 * depth, and in E2EE mode (opaque profiles) the client side is all there is.
 */
function sanitizeProfile(p: ParticipantProfile | undefined, clientId: string): ParticipantProfile {
  let h = 0;
  for (let i = 0; i < clientId.length; i++) h = (h * 31 + clientId.charCodeAt(i)) >>> 0;
  const fallbackColor = PROFILE_PALETTE[h % PROFILE_PALETTE.length];
  // eslint-disable-next-line no-control-regex
  const name = (p?.name ?? "").replace(/[\x00-\x1f\x7f]/g, "").trim().slice(0, 40);
  const color = /^#[0-9a-f]{6}$/i.test(p?.color ?? "") ? p!.color.toLowerCase() : fallbackColor;
  return { name: name || `Guest ${(h % 900) + 100}`, color };
}

/** Fresh 128-bit hex epoch id. Server-side identifier generation is exempt
 * from the intent-path determinism rule (doc 05 — that rule governs values
 * that feed serialized XML or apply behavior; epoch ids are bookkeeping). */
function defaultGenesisId(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  let hex = "";
  for (const b of bytes) hex += b.toString(16).padStart(2, "0");
  return `g_${hex}`;
}

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  // btoa exists in Node ≥16 globals and browsers.
  return typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
}
