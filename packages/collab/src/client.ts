/**
 * Client entry surface for @wordinweb/collab: the optimistic replica, the
 * client-side connection that drives it over a transport, the wire types, and
 * the transform (used by reconciliation replay).
 */
export { ClientReplica } from "./replica.js";
export { InMemoryBundleStore, BundlePersister, type DocBundle, type BundleStore } from "./bundle.js";
export { VersionRing, InMemoryVersionStore, type DocVersion, type VersionStore } from "./versions.js";
export { CollabConnection, type ClientTransport, type ConnectionCallbacks } from "./connection.js";
export { EncryptedCollabConnection } from "./enc-connection.js";
export { bindEditor, type EditorBridge, type EditorBinding } from "./binding.js";
export { createWebSocketTransport, type SocketLike } from "./ws-transport.js";
export * from "./intents.js";
export {
  PROTOCOL_VERSION,
  type ClientMessage,
  type ServerMessage,
  type PresencePosition,
  type ParticipantProfile,
  type RosterEntry,
  type EnvelopeEntry,
  type SealedCheckpoint,
  type LineageHead,
  ENGINE_VERSION,
} from "./protocol.js";
export { transformIntent, transformPosition } from "./transform.js";
export { sealIntent, openIntent, sealCheckpoint, openCheckpoint, deriveEpochKeys, stretchShareCode, mintDocKey, docKeyFromFragment, type IntentEnvelope, type EpochKeys } from "./e2ee.js";
export { docHash } from "./hash.js";
export { bytesToB64, b64ToBytes } from "./e2ee.js";
