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
    }
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
    }
  | { t: "broadcast"; entries: LogEntry[] }
  /** `participant` is the sender's bound clientId (round-4 F14) — the same
   * identifier intents carry — so presence joins the roster/attribution
   * keyspace and survives the sender reconnecting on a new socket. */
  | { t: "presence"; participant: string; position: PresencePosition | null }
  | { t: "refused"; reason: string };

/** Bump when the intent apply/transform semantics change in a way that makes
 * mixed-version clients diverge (plan doc 03 version lockstep). A client whose
 * version differs is refused at hello with "please refresh".
 * v2: hello carries clientId (bound identity) + takeover; welcome carries the
 * id sidecar. */
export const PROTOCOL_VERSION = 2;
