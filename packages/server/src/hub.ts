import { DocxDocument } from "@wordinweb/core";
import { DocumentSession, type IdSidecar, type ParticipantProfile } from "@wordinweb/collab/server";
import type { EnvelopeEntry, SealedCheckpoint, IntentEnvelope } from "@wordinweb/collab/server";
import { ClientMessage, ServerMessage, PROTOCOL_VERSION } from "@wordinweb/collab/server";
import { StorageDriver } from "./storage.js";

/**
 * A connection the hub can send to, abstracting the transport (WebSocket in
 * production, a fake in tests). `id` identifies the socket; `send` delivers a
 * server message.
 */
export interface Connection {
  id: string;
  send(msg: ServerMessage): void;
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
  /** The engine version the epoch's FIRST participant registered; everyone
   * after must match (client-derived canonical forms diverge across
   * transform-semantics versions with no arbiter — doc 13 §2). */
  engineVersion?: string;
}

interface Room {
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
  roster: Map<string, { profile: ParticipantProfile; connected: boolean }>;
  /** Epoch-ms timestamp of when the room last became empty, or undefined
   * while anyone is connected. Drives zero-custody eviction (plan doc 12
   * §2): after the last-disconnect grace elapses the room — doc, log,
   * everything — is deleted. Idle-with-people never evicts (round-4 F7);
   * only the empty state starts this clock. */
  emptySince?: number;
}

/** Last-disconnect grace before an empty room is deleted (plan doc 12 §1):
 * long enough to survive a page refresh or a flaky reconnect, short enough
 * that "everyone left" promptly means "the server holds nothing". */
export const EVICTION_GRACE_MS = 60_000;

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
  ) {}

  /** True if the connection's token has expired (live-socket enforcement). */
  private expired(connId: string): boolean {
    const a = this.auth.get(connId);
    return a?.expiresAt !== undefined && a.expiresAt <= this.now();
  }

  /** Force-disconnect a connection with a refusal (expiry/revocation). */
  private kick(conn: Connection, reason: string): void {
    conn.send({ t: "refused", reason });
    this.disconnect(conn);
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
          conn.send({ t: "refused", reason: "version-mismatch" });
          return;
        }
        if (!msg.clientId) {
          // The bound identity is not optional — every security property
          // downstream (dedup, attribution, single-tab) hangs off it.
          conn.send({ t: "refused", reason: "client-id-required" });
          return;
        }
        if (this.verifier) {
          const authed = this.verifier.verify(msg.token, msg.docId);
          if (!authed) {
            conn.send({ t: "refused", reason: "unauthorized" });
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
            conn.send({ t: "refused", reason: "already-open" });
            return;
          }
        }
        // Zero-custody posture (no provider): rooms exist only while seeded
        // and live. An unknown docId is not an error the server can fix —
        // the DOCUMENT lives in participants' bundles, so the answer is
        // `no-session`: the client offers "Bring it back live" (doc 12 §5.3)
        // and re-seeds from its own copy.
        if (!this.rooms.has(msg.docId) && !this.provider) {
          conn.send({ t: "refused", reason: "no-session" });
          return;
        }
        // Encrypted rooms enforce the ENGINE fence before admission
        // (doc 13 §2): canonical forms are client-derived there, so mixed
        // transform semantics would diverge with no arbiter. First joiner
        // registers the epoch's version; later mismatches are refused with
        // the download/draft path client-side.
        const encRoom = this.rooms.get(msg.docId)?.enc ?? null;
        if (encRoom) {
          if (!msg.engineVersion) {
            conn.send({ t: "refused", reason: "engine-version-required" });
            return;
          }
          if (encRoom.engineVersion && encRoom.engineVersion !== msg.engineVersion) {
            conn.send({ t: "refused", reason: "engine-version-mismatch" });
            return;
          }
          encRoom.engineVersion = msg.engineVersion;
        }
        this.connClient.set(conn.id, msg.clientId);
        const room = await this.room(msg.docId);
        room.conns.add(conn);
        room.emptySince = undefined; // occupied: the eviction clock stops.
        this.connDoc.set(conn.id, msg.docId);
        this.connById.set(conn.id, conn);
        if (room.enc) {
          // Blind welcome (doc 13 §3): sealed seed checkpoint + the WHOLE
          // epoch envelope log after it — base-complete by construction
          // because nothing is ever pruned within an epoch (blocker 1).
          conn.send({
            t: "welcome-enc",
            docId: msg.docId,
            genesisId: room.genesisId,
            checkpoint: room.enc.checkpoint,
            tail: room.enc.log.filter((e) => e.seq > room.enc!.checkpoint.seq),
            mode: "encrypted",
          });
          room.roster.set(msg.clientId, {
            profile: sanitizeProfile(msg.profile, msg.clientId),
            connected: true,
          });
          this.broadcastRoster(room);
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
        });
        // Roster upsert + fan-out (doc 14 §2): keyed by the BOUND clientId,
        // so a reconnect resumes the same entry. Sanitization is server-side
        // here (plaintext mode); clients still render defensively (doc 11
        // vector 7), which is also what E2EE mode will rely on entirely.
        room.roster.set(msg.clientId, {
          profile: sanitizeProfile(msg.profile, msg.clientId),
          connected: true,
        });
        this.broadcastRoster(room);
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
          conn.send({ t: "refused", reason: "not-joined" });
          return;
        }
        // Live-socket token-expiry enforcement (gate 5): a session outlives
        // its token only until the next action, then is cut.
        if (this.expired(conn.id)) {
          this.kick(conn, "token-expired");
          return;
        }
        // Role enforcement at the sequencer (doc 06): a viewer's edits are
        // refused, not merely hidden in the UI.
        if (this.verifier && this.auth.get(conn.id)?.role !== "editor") {
          conn.send({ t: "refused", reason: "read-only" });
          return;
        }
        // Identity binding (doc 11 decision 8): the intent must carry the
        // clientId this socket registered at hello. Anything else is a
        // forgery attempt (or a client bug) — refused BEFORE the session
        // sees it, so a forged (clientId, clientSeq) can never enter the
        // dedup set and swallow the victim's real intent.
        if (msg.intent.clientId !== this.connClient.get(conn.id)) {
          conn.send({ t: "refused", reason: "client-id-mismatch" });
          return;
        }
        const room = this.rooms.get(docId)!;
        if (room.enc) {
          // No mixed-mode documents (doc 13 §6): a plaintext submit into an
          // encrypted room is refused outright — a blind server could not
          // validate it, and honest clients derive mode from their link.
          conn.send({ t: "refused", reason: "wrong-mode" });
          return;
        }
        const before = room.session!.seq;
        const entry = room.session!.submit(msg.intent);
        // A deduplicated re-send returns an already-sequenced entry (seq <=
        // before); persist only genuinely new entries so the log stays
        // append-once. Durability before ack (plan doc 06), then broadcast.
        if (this.storage && entry.seq > before) await this.storage.appendEntries(docId, [entry]);
        const out: ServerMessage = { t: "broadcast", entries: [entry] };
        for (const c of room.conns) c.send(out);
        return;
      }
      case "submit-enc": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) {
          conn.send({ t: "refused", reason: "not-joined" });
          return;
        }
        if (this.expired(conn.id)) {
          this.kick(conn, "token-expired");
          return;
        }
        if (this.verifier && this.auth.get(conn.id)?.role !== "editor") {
          conn.send({ t: "refused", reason: "read-only" });
          return;
        }
        // Identity binding applies identically in encrypted mode — the
        // bookkeeping is plaintext precisely so the server can enforce it
        // (doc 13 §2 precondition #1).
        if (msg.envelope.clientId !== this.connClient.get(conn.id)) {
          conn.send({ t: "refused", reason: "client-id-mismatch" });
          return;
        }
        const room = this.rooms.get(docId)!;
        if (!room.enc) {
          conn.send({ t: "refused", reason: "wrong-mode" });
          return;
        }
        // Ciphertext size cap (doc 13 §2: one of the few things a blind
        // server CAN enforce). 256 KB covers any real intent incl. paste.
        if (msg.envelope.ciphertext.length > 256 * 1024) {
          conn.send({ t: "refused", reason: "too-large" });
          return;
        }
        // Dedup by plaintext bookkeeping; re-sends return the prior entry.
        const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
        let entry = room.enc.seen.has(key)
          ? room.enc.log.find((e) => `${e.clientId}:${e.clientSeq}` === key)
          : undefined;
        if (!entry) {
          room.enc.seen.add(key);
          const seq = room.enc.log.length === 0 ? 1 : room.enc.log[room.enc.log.length - 1].seq + 1;
          entry = { ...msg.envelope, seq };
          room.enc.log.push(entry);
        }
        const outEnc: ServerMessage = { t: "broadcast-enc", entries: [entry] };
        for (const c of room.conns) c.send(outEnc);
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
        const out: ServerMessage = {
          t: "presence",
          participant: this.connClient.get(conn.id)!,
          position: msg.position,
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
    this.connDoc.delete(conn.id);
    this.auth.delete(conn.id);
    this.connById.delete(conn.id);
    this.connClient.delete(conn.id);
    if (docId) {
      const room = this.rooms.get(docId);
      if (room) {
        room.conns.delete(conn);
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
      }
    }
  }

  /** Fan the full roster snapshot to everyone in the room (doc 14 §2). */
  private broadcastRoster(room: Room): void {
    const roster = [...room.roster].map(([clientId, e]) => ({
      clientId,
      profile: e.profile,
      connected: e.connected,
    }));
    const out: ServerMessage = { t: "roster", roster };
    for (const c of room.conns) c.send(out);
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
  sweepRooms(graceMs: number = EVICTION_GRACE_MS): string[] {
    const evicted: string[] = [];
    for (const [docId, room] of [...this.rooms]) {
      if (room.conns.size === 0 && room.emptySince !== undefined && this.now() - room.emptySince >= graceMs) {
        this.rooms.delete(docId);
        evicted.push(docId);
      }
    }
    return evicted;
  }

  /** Sessions currently held in memory (for eviction/metrics). */
  activeDocs(): string[] {
    return [...this.rooms.keys()];
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
  ): { ok: true; genesisId: string } | { ok: false; reason: "exists"; genesisId: string } {
    const existing = this.rooms.get(docId);
    if (existing) return { ok: false, reason: "exists", genesisId: existing.genesisId };
    const session = new DocumentSession(DocxDocument.load(docx));
    if (sidecar) session.installSidecar(sidecar);
    const room: Room = {
      session,
      enc: null,
      conns: new Set(),
      genesisId: this.genGenesisId(),
      roster: new Map(),
      emptySince: this.now(), // eviction clock runs until someone joins
    };
    this.rooms.set(docId, room);
    return { ok: true, genesisId: room.genesisId };
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
  ): { ok: true; genesisId: string } | { ok: false; reason: "exists"; genesisId: string } {
    const existing = this.rooms.get(docId);
    if (existing) return { ok: false, reason: "exists", genesisId: existing.genesisId };
    const room: Room = {
      session: null,
      enc: { checkpoint, log: [], seen: new Set() },
      conns: new Set(),
      genesisId,
      roster: new Map(),
      emptySince: this.now(),
    };
    this.rooms.set(docId, room);
    return { ok: true, genesisId };
  }

  private async room(docId: string): Promise<Room> {
    let room = this.rooms.get(docId);
    if (!room) {
      const session = await this.startSession(docId);
      room = { session, enc: null, conns: new Set(), genesisId: this.genGenesisId(), roster: new Map() };
      this.rooms.set(docId, room);
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
  const name = (p?.name ?? "").replace(/[ -]/g, "").trim().slice(0, 40);
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
