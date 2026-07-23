/**
 * @wordinweb/server — the collaborative-editing host.
 *
 * The CollabHub is transport-agnostic thin plumbing over @wordinweb/collab's
 * authoritative DocumentSession. A WebSocket adapter (or any transport that
 * can deliver protocol messages) drives it; the hub owns no editing logic and
 * no sockets, so it is unit-testable with in-memory connections.
 */
export { CollabHub, type Connection, type DocProvider, type TokenVerifier, type AuthResult } from "./hub.js";
export { PROTOCOL_VERSION, type ClientMessage, type ServerMessage, type PresencePosition } from "@wordinweb/collab/server";
export { attachWebSocketServer, type WsServer, type WsSocket } from "./ws.js";
export { InMemoryStorage, type StorageDriver } from "./storage.js";
export { blankDocxBytes, blankProvider } from "./blank.js";
export { startDevServer } from "./cli.js";
export {
  makeDocId,
  PartyPool,
  RateLimiter,
  DEMO_INTENT_ALLOWLIST,
  intentAllowedInDemo,
  type DemoMode,
} from "./demo.js";
