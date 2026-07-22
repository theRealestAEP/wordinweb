/**
 * Server entry surface for @wordinweb/collab (transport-free).
 * The DocumentSession is the authoritative state machine a transport (a
 * WebSocket host in packages/server) drives; nothing here touches sockets,
 * storage, or timers.
 */
export { DocumentSession, type Broadcast } from "./session.js";
export { applyIntent } from "./apply.js";
export * from "./intents.js";
export * from "./transform.js";
