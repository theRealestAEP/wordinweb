import type { Intent, LogEntry } from "./intents.js";
import type { IdSidecar } from "./session.js";

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
    }
  /** Update this connection's profile mid-session (rename / recolor). */
  | { t: "profile"; profile: ParticipantProfile }
  | { t: "submit"; intent: Intent }
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
    }
  | { t: "broadcast"; entries: LogEntry[] }
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
