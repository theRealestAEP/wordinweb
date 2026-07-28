/**
 * @wordinweb/server — the collaborative-editing host.
 *
 * The CollabHub is transport-agnostic thin plumbing over @wordinweb/collab's
 * authoritative DocumentSession. A WebSocket adapter (or any transport that
 * can deliver protocol messages) drives it; the hub owns no editing logic and
 * no sockets, so it is unit-testable with in-memory connections.
 */
export {
  CollabHub,
  EVICTION_GRACE_MS,
  type Connection,
  type DocProvider,
  type TokenVerifier,
  type AuthResult,
  type EvictionReason,
  type SeedResult,
} from "./hub.js";
// Session lifecycle + per-IP abuse limits. Exported because an embedder that
// builds its own hub needs the same knobs the bundled server has — including
// `clientIp`, which is the security boundary the per-IP limits rest on and
// must not be reimplemented casually (see its doc comment).
export {
  envLimits,
  normalizeLimits,
  clientIp,
  normalizeIp,
  DEFAULT_LIMITS,
  LIFECYCLE_NOTES,
  type HubLimits,
  type LifecycleLimits,
  type IpLimits,
  type IpRequestLike,
} from "./limits.js";
export { IpGuard, NO_OP_IP_GUARD, type IpDecision, type IpLimitReason } from "./ip-guard.js";
export { handleSeedRequest, type SeedHttpRequest, type SeedHttpResponse, type SeedHttpOptions } from "./seed-http.js";
export { PROTOCOL_VERSION, type ClientMessage, type ServerMessage, type PresencePosition } from "@wordinweb/collab/server";
export { attachWebSocketServer, type WsServer, type WsSocket, type WsRequest, type WsGuardOptions } from "./ws.js";
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
