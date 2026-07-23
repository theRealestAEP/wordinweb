import type { Intent, LogEntry } from "./intents.js";
import type { IdSidecar } from "./session.js";
import type { IntentEnvelope } from "./e2ee.js";

/** A sequenced envelope in an encrypted room's log (doc 13 §2): the
 * server-assigned seq plus the opaque envelope. The server can read only
 * the bookkeeping; clients derive the canonical form themselves. */
export type EnvelopeEntry = IntentEnvelope & { seq: number };

/**
 * One head in a document's LINEAGE chain (doc 15 §1): the identity of the
 * confirmed state a session epoch ended at. Ancestry is decided by hash
 * membership — never by dates (clocks are self-asserted). A rejoiner whose
 * own head appears in the seed's lineage is a strict ancestor and can
 * FAST-FORWARD silently; anything else is true divergence (draft/fork).
 */
export interface LineageHead {
  genesisId: string;
  seq: number;
  /** sha256 hex of the confirmed docx bytes at (genesisId, seq). */
  docHash: string;
}

/** A sealed checkpoint blob as the wire carries it (doc 13 §3). */
export interface SealedCheckpoint {
  seq: number;
  iv: string;
  ciphertext: string;
}

/**
 * The engine version fence for E2EE rooms (doc 13 §2): canonical forms are
 * CLIENT-derived in encrypted mode, so a room with mixed transform/apply
 * semantics can diverge with no arbiter. Bump whenever those semantics
 * change; the sequencer refuses a hello whose engineVersion mismatches the
 * room's (the stale client gets the download/draft path). Plaintext rooms
 * ignore it (the server's own apply is the arbiter there, and
 * PROTOCOL_VERSION already gates wire compatibility).
 */
export const ENGINE_VERSION = "e1";

/**
 * Wire protocol between a collab client and the server host. Transport-
 * agnostic: the hub speaks these messages; a WebSocket (or in-memory test)
 * adapter serializes them. Kept deliberately small (plan doc 06).
 */

/** An ephemeral cursor/selection position on the presence channel: stable-id
 * addresses (plan doc 03). Never logged, never persisted. */
export interface PresencePosition {
  anchor: { blockId: number; runId: number; offset: number };
  focus?: { blockId: number; runId: number; offset: number };
}

/**
 * A participant's self-asserted display profile (plan doc 14 §2). No
 * accounts: anyone with the link can claim any name — stated in the UI
 * ("names are chosen by participants") — but the F4 clientId binding
 * guarantees CONTINUITY: one name maps to one actor per session and nobody
 * can impersonate an existing participant mid-session. The server sanitizes
 * (length, control chars, color shape) and clients still render text-node-
 * only with palette-validated colors (doc 11 XSS vector 7 — defense in
 * depth, since a hostile server or E2EE mode moves sanitization client-side).
 */
export interface ParticipantProfile {
  /** 1–40 chars after server sanitization. */
  name: string;
  /** Hex color `#rrggbb`; invalid values are replaced server-side by a
   * palette color derived from the clientId hash. */
  color: string;
}

/** One roster row (doc 14 §2): keyed by the BOUND clientId — the same
 * identity intents carry, so attribution, presence, and roster share one
 * keyspace. Disconnected participants stay listed (greyed) for the session's
 * lifetime; a reconnect under the same clientId resumes the entry. */
export interface RosterEntry {
  clientId: string;
  profile: ParticipantProfile;
  connected: boolean;
}

/** Client → server. */
export type ClientMessage =
  | {
      t: "hello";
      protocolVersion: number;
      docId: string;
      /**
       * The identity every subsequent submit/presence message is bound to.
       * The hub registers it for this socket at hello and REFUSES any submit
       * whose intent.clientId differs (plan doc 11 decision 8 / round-4 F4):
       * `(clientId, clientSeq)` is the idempotency key, so an unbound claimed
       * id would let one client poison another's dedup space (silent edit
       * loss) and forge attribution. Registration-on-first-hello is
       * sufficient — the id needs continuity, not global meaning.
       */
      clientId: string;
      /**
       * Single-live-connection rule (plan doc 12 §7): a second hello for the
       * same (docId, clientId) is refused `already-open` — same-profile tabs
       * share localStorage identity and would collide clientSeq counters.
       * With `takeover: true` the NEW connection wins instead: the hub kicks
       * the old socket (it may be a zombie tab) and admits this one.
       */
      takeover?: boolean;
      token?: string;
      sinceSeq: number;
      /** The epoch id from the client's stored bundle, when resuming (plan
       * doc 12 §5). Informational to the server today — epoch comparison is
       * client-side (welcome.genesisId vs the bundle's) — but carried so the
       * server can log resume-vs-fresh joins and future-proof case handling. */
      genesisId?: string;
      /** Display profile (doc 14 §2) — optional; omitted = anonymous default
       * derived server-side. In E2EE mode this becomes an opaque encrypted
       * blob (doc 13/14) — the shape change rides that protocol bump. */
      profile?: ParticipantProfile;
      /** Engine (transform/apply semantics) version — REQUIRED for encrypted
       * rooms, where the fence prevents mixed-semantics divergence. */
      engineVersion?: string;
      /**
       * Share-code proof (doc 13 §7): the PBKDF2-stretched code, base64.
       * The server compares against the verifier registered at seed time and
       * enforces an attempt budget — 10^6 combinations is plenty against 5
       * online tries, and in E2EE mode the code ALSO mixes into key
       * derivation, so even a server bypass yields undecryptable content.
       */
      codeProof?: string;
    }
  /** Update this connection's profile mid-session (rename / recolor). */
  | { t: "profile"; profile: ParticipantProfile }
  | { t: "submit"; intent: Intent }
  /** Encrypted-mode submit (doc 13 §2): opaque body, plaintext bookkeeping.
   * The hub refuses this on plaintext rooms and refuses plaintext `submit`
   * on encrypted rooms — no mixed-mode documents, enforced both ways. */
  | { t: "submit-enc"; envelope: IntentEnvelope }
  /** Upload a sealed checkpoint (doc 13 §3) — accepted ONLY from the
   * connection the server currently designates as checkpointer. */
  | { t: "checkpoint"; checkpoint: SealedCheckpoint }
  /** Media re-supply (doc 16 §3): "I need blob <sha>" (relay miss). */
  | { t: "media-need"; sha: string }
  /** "I hold these blobs" — reply to media-request, or unsolicited right
   * after a welcome whose mediaNeeded intersects local holdings (§5.4). */
  | { t: "media-have"; shas: string[] }
  /** Encrypted hash gossip (doc 13 §2): an OPAQUE sealed {seq, hash} blob
   * the server relays without reading — divergence detection must not leak
   * even document hashes to a blind server (a hash is a stable content
   * fingerprint: confirmable by anyone holding a guess). */
  | { t: "gossip"; iv: string; ciphertext: string }
  | { t: "presence"; position: PresencePosition | null };

/** Server → client. */
export type ServerMessage =
  | {
      t: "welcome";
      docId: string;
      seq: number;
      snapshot: string;
      /**
       * The stable-id sidecar for the snapshot (plan doc 12 §2, round-4 F10).
       * A snapshot NEVER travels without it: the joiner cannot re-derive the
       * id table from parse order once history contains split-created carried
       * ids (round-2 F1) — it would silently mis-address every later intent.
       */
      sidecar: IdSidecar;
      tail: LogEntry[];
      /**
       * The session EPOCH id (plan doc 12): minted fresh every time a doc is
       * seeded/re-seeded. A resuming client compares it against its bundle's
       * stored genesisId — same ⇒ seamless rejoin (case 1); different ⇒
       * someone re-seeded while it was away (case 2: take server state, keep
       * the local copy per doc 15 lineage — NEVER silently merge epochs).
       */
      genesisId: string;
      /** Session encryption mode (plan doc 13 §6). Clients derive the truth
       * from their link (#k present ⇒ encrypted) and HARD-REFUSE a welcome
       * that contradicts it — the wire value must never downgrade a client. */
      mode: "plaintext" | "encrypted";
      /** The seed's lineage chain (doc 15): lets a rejoining holder decide
       * fast-forward vs fork CLIENT-side. Absent for provider-created and
       * legacy rooms (⇒ every epoch mismatch is a fork, the safe default). */
      lineage?: LineageHead[];
      /** Shas with waiters / outstanding unavailability (doc 16 §3): a
       * joining holder intersects with its local media and volunteers —
       * the mechanism behind "reappears when a holder rejoins". */
      mediaNeeded?: string[];
    }
  | { t: "broadcast"; entries: LogEntry[] }
  /**
   * Encrypted-mode welcome (doc 13 §3): the seed checkpoint (sealed by the
   * epoch's seeder) + the epoch's envelope log after it. BASE-COMPLETE by
   * construction (round-4 blocker 1): the demo server retains the WHOLE
   * epoch log in RAM (doc 12 — rooms are session-scoped), so every entry a
   * joiner needs to re-derive canonical forms is present; client-produced
   * mid-session checkpoints (doc 13 item 6) are a RAM optimization on top,
   * not a correctness requirement, and land with the retention rule.
   */
  | {
      t: "welcome-enc";
      docId: string;
      genesisId: string;
      checkpoint: SealedCheckpoint;
      tail: EnvelopeEntry[];
      mode: "encrypted";
      mediaNeeded?: string[];
    }
  /** Encrypted-mode broadcast: sequenced opaque envelopes. */
  | { t: "broadcast-enc"; entries: EnvelopeEntry[] }
  /** Relayed gossip blob; `from` is the sender's BOUND clientId. */
  | { t: "gossip"; from: string; iv: string; ciphertext: string }
  /** Media re-supply control (doc 16 §4). request = broadcast "who has
   * it"; upload = to ONE chosen holder; ready/unavailable = to waiters. */
  | { t: "media-request"; sha: string }
  | { t: "media-upload"; sha: string }
  | { t: "media-ready"; sha: string }
  | { t: "media-unavailable"; sha: string }
  /** Checkpointer designation (doc 13 §3): the SERVER assigns the role —
   * the v1 lowest-clientId election was riggable (round-4 blocker 2); an
   * assigned role can only be held by an authenticated connection the
   * server picked. Rotated on disconnect. */
  | { t: "checkpointer"; active: boolean }
  /** `participant` is the sender's bound clientId (round-4 F14) — the same
   * identifier intents carry — so presence joins the roster/attribution
   * keyspace and survives the sender reconnecting on a new socket. */
  | { t: "presence"; participant: string; position: PresencePosition | null }
  /** Full roster snapshot, fanned out on every join/leave/profile change
   * (rooms are small; a snapshot beats delta bookkeeping). Ephemeral like
   * presence: never logged, never persisted, dies with the room. */
  | { t: "roster"; roster: RosterEntry[] }
  | { t: "refused"; reason: string };

/** Bump when the intent apply/transform semantics change in a way that makes
 * mixed-version clients diverge (plan doc 03 version lockstep). A client whose
 * version differs is refused at hello with "please refresh".
 * v2: hello carries clientId (bound identity) + takeover; welcome carries the
 * id sidecar. */
export const PROTOCOL_VERSION = 2;
