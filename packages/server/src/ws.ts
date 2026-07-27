import { CollabHub } from "./hub.js";
import { ClientMessage, ServerMessage } from "@wordinweb/collab/server";
import { type Observability, NO_OP_OBSERVABILITY } from "./observability.js";

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
 *
 * Socket open/close are counted HERE (not in the hub's `disconnect`, which a
 * takeover kick can invoke twice) so the live-connection gauge stays balanced
 * — the ws `close` event fires exactly once per socket. `obs` is passive and
 * defaults to a no-op.
 */
export function attachWebSocketServer(
  wss: WsServer,
  hub: CollabHub,
  obs: Observability = NO_OP_OBSERVABILITY,
): void {
  let nextId = 0;
  wss.on("connection", (socket) => {
    const conn = {
      id: `ws${nextId++}`,
      send: (msg: ServerMessage) => socket.send(JSON.stringify(msg)),
    };
    obs.connectionOpened();
    socket.on("message", (data: unknown) => {
      let msg: ClientMessage;
      try {
        msg = JSON.parse(typeof data === "string" ? data : String(data)) as ClientMessage;
      } catch {
        // SILENCE IS CORRECT HERE, and this comment is the decision rather
        // than an accident: the frame came from an untrusted peer, so a stack
        // trace of "unexpected token < in JSON" says nothing about OUR
        // health and a peer spraying garbage could otherwise spend our disk.
        // It is COUNTED (badFrame) so the volume is still visible.
        obs.badFrame();
        return;
      }
      if (msg && typeof msg === "object" && "t" in msg) {
        // THE SWALLOW THAT HID THREE BUGS. This used to be `.catch(() => {})`,
        // which meant any throw inside the hub's message handling — a
        // TypeError in a new code path, a rejected storage write, anything —
        // disappeared completely: the socket stayed open, the client waited
        // forever for a reply that would never come, and the server's logs
        // said the session was healthy. Now it is one error line with a
        // stack, plus the message TYPE (fixed protocol vocabulary, never the
        // payload) so the failing path is identifiable.
        void hub.handle(conn, msg).catch((err: unknown) => {
          obs.error("ws.handle", err, { msgType: String((msg as { t?: unknown }).t) });
        });
      }
    });
    socket.on("close", () => {
      hub.disconnect(conn);
      obs.connectionClosed();
    });
  });
}
