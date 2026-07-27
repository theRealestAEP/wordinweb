/**
 * @wordinweb/server — the collaborative-editing host.
 *
 * The CollabHub is transport-agnostic thin plumbing over @wordinweb/collab's
 * authoritative DocumentSession. A WebSocket adapter (or any transport that
 * can deliver protocol messages) drives it; the hub owns no editing logic and
 * no sockets, so it is unit-testable with in-memory connections.
 */
export { CollabHub, EVICTION_GRACE_MS, type Connection, type DocProvider, type TokenVerifier, type AuthResult } from "./hub.js";
export { handleSeedRequest, type SeedHttpRequest, type SeedHttpResponse, type SeedHttpOptions } from "./seed-http.js";
export { PROTOCOL_VERSION, type ClientMessage, type ServerMessage, type PresencePosition } from "@wordinweb/collab/server";
export { attachWebSocketServer, type WsServer, type WsSocket } from "./ws.js";
export { InMemoryStorage, type StorageDriver } from "./storage.js";
export { blankDocxBytes, blankProvider } from "./blank.js";
export { startDevServer, startZeroCustodyServer } from "./cli.js";
export {
  MetricsObservability,
  NO_OP_OBSERVABILITY,
  createObservability,
  observabilityEnabled,
  type Observability,
  type ObsSnapshot,
  type ObservabilityOptions,
} from "./observability.js";
export {
  makeDocId,
  PartyPool,
  RateLimiter,
  DEMO_INTENT_ALLOWLIST,
  intentAllowedInDemo,
  type DemoMode,
} from "./demo.js";
