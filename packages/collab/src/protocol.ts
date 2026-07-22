import type { Intent, LogEntry } from "./intents.js";

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
  | { t: "hello"; protocolVersion: number; docId: string; token?: string; sinceSeq: number }
  | { t: "submit"; intent: Intent }
  | { t: "presence"; position: PresencePosition | null };

/** Server → client. */
export type ServerMessage =
  | { t: "welcome"; docId: string; seq: number; snapshot: string; tail: LogEntry[] }
  | { t: "broadcast"; entries: LogEntry[] }
  | { t: "presence"; participant: string; position: PresencePosition | null }
  | { t: "refused"; reason: string };

/** Bump when the intent apply/transform semantics change in a way that makes
 * mixed-version clients diverge (plan doc 03 version lockstep). A client whose
 * version differs is refused at hello with "please refresh". */
export const PROTOCOL_VERSION = 1;
