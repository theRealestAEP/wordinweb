/**
 * Client entry surface for @wordinweb/collab: the optimistic replica, the
 * client-side connection that drives it over a transport, the wire types, and
 * the transform (used by reconciliation replay).
 */
export { ClientReplica } from "./replica.js";
export { CollabConnection, type ClientTransport, type ConnectionCallbacks } from "./connection.js";
export { bindEditor, type EditorBridge, type EditorBinding } from "./binding.js";
export { createWebSocketTransport, type SocketLike } from "./ws-transport.js";
export * from "./intents.js";
export {
  PROTOCOL_VERSION,
  type ClientMessage,
  type ServerMessage,
  type PresencePosition,
} from "./protocol.js";
export { transformIntent, transformPosition } from "./transform.js";
