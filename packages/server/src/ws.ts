import { CollabHub } from "./hub.js";
import { ClientMessage, ServerMessage } from "./protocol.js";

/**
 * Minimal structural shapes for a `ws`-style server and socket, so this
 * adapter needs no `ws` dependency at build time — the runnable entrypoint
 * passes a real `WebSocketServer`. This is the only transport-touching code
 * in the package; everything else is the transport-free hub (plan doc 07).
 */
export interface WsSocket {
  send(data: string): void;
  on(event: "message", cb: (data: unknown) => void): void;
  on(event: "close", cb: () => void): void;
}
export interface WsServer {
  on(event: "connection", cb: (socket: WsSocket) => void): void;
}

/**
 * Wire a ws-style server to a CollabHub. Each socket becomes a hub Connection;
 * inbound frames are parsed as ClientMessages and dispatched; hub sends are
 * serialized to JSON frames. Malformed frames are ignored (a hostile client
 * cannot crash the room).
 */
export function attachWebSocketServer(wss: WsServer, hub: CollabHub): void {
  let nextId = 0;
  wss.on("connection", (socket) => {
    const conn = {
      id: `ws${nextId++}`,
      send: (msg: ServerMessage) => socket.send(JSON.stringify(msg)),
    };
    socket.on("message", (data: unknown) => {
      let msg: ClientMessage;
      try {
        msg = JSON.parse(typeof data === "string" ? data : String(data)) as ClientMessage;
      } catch {
        return;
      }
      if (msg && typeof msg === "object" && "t" in msg) hub.handle(conn, msg);
    });
    socket.on("close", () => hub.disconnect(conn));
  });
}
