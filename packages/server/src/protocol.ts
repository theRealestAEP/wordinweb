import type { Intent, LogEntry } from "@wordinweb/collab/server";

/**
 * Wire protocol between a collab client and the server host. Transport-
 * agnostic: the hub speaks these messages; a WebSocket (or in-memory test)
 * adapter serializes them. Kept deliberately small (plan doc 06).
 */

/** Client → server. */
export type ClientMessage =
  | { t: "hello"; protocolVersion: number; docId: string; token?: string; sinceSeq: number }
  | { t: "submit"; intent: Intent };

/** Server → client. */
export type ServerMessage =
  | { t: "welcome"; docId: string; seq: number; snapshot: string; tail: LogEntry[] }
  | { t: "broadcast"; entries: LogEntry[] }
  | { t: "refused"; reason: string };

/** Bump when the intent apply/transform semantics change in a way that makes
 * mixed-version clients diverge (plan doc 03 version lockstep). A client whose
 * version differs is refused at hello with "please refresh". */
export const PROTOCOL_VERSION = 1;
