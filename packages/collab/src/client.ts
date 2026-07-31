/**
 * Client entry surface for @wordinweb/collab: the optimistic replica, the
 * client-side connection that drives it over a transport, the wire types, and
 * the transform (used by reconciliation replay).
 */
export { ClientReplica } from "./replica.js";
export {
  InMemoryBundleStore,
  BundlePersister,
  parseBundleKey,
  versionKey,
  draftKey,
  supersededKey,
  type DocBundle,
  type BundleStore,
  type StoredDocSummary,
  type ParsedBundleKey,
  type StoredDocKind,
} from "./bundle.js";
export { VersionRing, InMemoryVersionStore, type DocVersion, type VersionStore } from "./versions.js";
export { toSuggestions, replayDrained, arrivalMode, OFFLINE_TAIL_CAP } from "./rebase.js";
// The canonical apply, exposed for OFFLINE editing (doc 12 §5): with no
// connection to route through, the react layer applies toolbar/API intents
// to the in-hand document through the SAME code every replica runs.
export { applyIntentScoped, resyncScope, unionScopes, type Scope } from "./apply.js";
export { CollabConnection, type ClientTransport, type ConnectionCallbacks } from "./connection.js";
export { EncryptedCollabConnection } from "./enc-connection.js";
export { CarriedIdAllocator } from "./id-allocator.js";
export { bindEditor, type EditorBridge, type EditorBinding } from "./binding.js";
export { createWebSocketTransport, type SocketLike } from "./ws-transport.js";
export { LivenessMonitor, monitorTransport, type ConnectionState, type LivenessOptions } from "./liveness.js";
export * from "./intents.js";
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
  type WriteStatus,
  type EnvelopeEntry,
  type SealedCheckpoint,
  type LineageHead,
  ENGINE_VERSION,
} from "./protocol.js";
export { transformIntent, transformPosition } from "./transform.js";
export { sealIntent, openIntent, sealCheckpoint, openCheckpoint, type CheckpointBody, deriveEpochKeys, stretchShareCode, mintDocKey, docKeyFromFragment, type IntentEnvelope, type EpochKeys } from "./e2ee.js";
export { docHash } from "./hash.js";
export type { UndoOutcome, UndoCandidate } from "./session.js";
export { bytesToB64, b64ToBytes, sealMediaBlob, openMediaBlob } from "./e2ee.js";
export { sealPresence, openPresence, isSealedPresence, type SealedPresence } from "./e2ee.js";
export {
  MediaClient,
  applyMediaAddresses,
  mediaAddressesOf,
  type MediaAddress,
  sha256Hex,
  putBlob,
  getBlob,
  type MediaTransportOptions,
  type MediaState,
  type MediaCrypto,
  type FetchLike,
} from "./media.js";
