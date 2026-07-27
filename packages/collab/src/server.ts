/**
 * Server entry surface for @wordinweb/collab (transport-free).
 * The DocumentSession is the authoritative state machine a transport (a
 * WebSocket host in packages/server) drives; nothing here touches sockets,
 * storage, or timers.
 */
export { DocumentSession, type Broadcast, type IdSidecar } from "./session.js";
export { applyIntent } from "./apply.js";
export { validateIntent, DEFAULT_INTENT_LIMITS, type IntentLimits } from "./validate.js";
export * from "./intents.js";
export * from "./transform.js";
export {
  PROTOCOL_VERSION,
  type ClientMessage,
  type ServerMessage,
  type PresencePosition,
  type PresenceRange,
  PRESENCE_MAX_RANGES,
  sanitizePresencePosition,
  type ParticipantProfile,
  type RosterEntry,
  type EnvelopeEntry,
  type SealedCheckpoint,
  type LineageHead,
  ENGINE_VERSION,
} from "./protocol.js";
export { sealIntent, openIntent, sealCheckpoint, openCheckpoint, deriveEpochKeys, stretchShareCode, mintDocKey, docKeyFromFragment, type IntentEnvelope, type EpochKeys } from "./e2ee.js";
