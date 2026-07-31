import { ClientTransport } from "./connection.js";
import { ClientMessage, ServerMessage } from "./protocol.js";

/** Minimal structural shape of a browser WebSocket (or a test double). */
export interface SocketLike {
  send(data: string): void;
  addEventListener(type: "message", cb: (ev: { data: unknown }) => void): void;
  addEventListener(type: "open", cb: () => void): void;
}

/**
 * Adapt a browser WebSocket to a ClientTransport: JSON-frame outbound client
 * messages, parse inbound server messages. Buffers sends until the socket is
 * open. The only browser-touching client code — everything else in
 * @wordinweb/collab is transport-free.
 */
export function createWebSocketTransport(socket: SocketLike): ClientTransport {
  let open = false;
  const backlog: string[] = [];
  socket.addEventListener("open", () => {
    open = true;
    for (const frame of backlog.splice(0)) socket.send(frame);
  });
  let handler: ((msg: ServerMessage) => void) | null = null;
  socket.addEventListener("message", (ev) => {
    let msg: ServerMessage;
    try {
      msg = JSON.parse(typeof ev.data === "string" ? ev.data : String(ev.data)) as ServerMessage;
    } catch {
      return;
    }
    handler?.(msg);
  });
  return {
    send(msg: ClientMessage): void {
      const frame = JSON.stringify(msg);
      if (open) socket.send(frame);
      else backlog.push(frame);
    },
    onMessage(cb): void {
      handler = cb;
    },
  };
}
