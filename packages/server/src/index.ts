/**
 * @wordinweb/server — the collaborative-editing host.
 *
 * The CollabHub is transport-agnostic thin plumbing over @wordinweb/collab's
 * authoritative DocumentSession. A WebSocket adapter (or any transport that
 * can deliver protocol messages) drives it; the hub owns no editing logic and
 * no sockets, so it is unit-testable with in-memory connections.
 */
export { CollabHub, type Connection, type DocProvider } from "./hub.js";
export { PROTOCOL_VERSION, type ClientMessage, type ServerMessage } from "./protocol.js";
export { attachWebSocketServer, type WsServer, type WsSocket } from "./ws.js";
