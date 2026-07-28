import { CollabHub } from "./hub.js";
import { ClientMessage, ServerMessage } from "@wordinweb/collab/server";
import { type Observability, NO_OP_OBSERVABILITY } from "./observability.js";
import { IpGuard, NO_OP_IP_GUARD } from "./ip-guard.js";
import { clientIp } from "./limits.js";

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
  /** Optional so in-memory test fakes need not implement it; the real `ws`
   * socket always has it. Used to hang up on a connection refused by the
   * per-IP cap — a refusal frame the peer can ignore is not a limit. */
  close?(): void;
  /** Bytes queued for this peer that Node has not managed to write yet. The
   * slow-consumer signal: a client that never reads makes this climb without
   * bound while every other limit stays satisfied. */
  readonly bufferedAmount?: number;
  /** Hard close, skipping the closing handshake — the right tool for a peer
   * that is already not draining, since a graceful close would just add more
   * to a buffer it is not reading. */
  terminate?(): void;
}
/**
 * The upgrade request behind a socket. `ws` passes it as the second argument
 * to the `connection` event (cli.ts already forwards it: `wss.emit(
 * "connection", client, req)`), and it is the ONLY place the peer address is
 * available — a WebSocket frame carries no addressing. Optional throughout so
 * every existing test fake, which emits a socket alone, still type-checks.
 */
export interface WsRequest {
  socket?: { remoteAddress?: string | null } | null;
  headers?: Record<string, string | string[] | undefined>;
}
export interface WsServer {
  on(event: "connection", cb: (socket: WsSocket, req?: WsRequest) => void): void;
}

/** Per-IP plumbing for the transport. Both fields optional: omitted ⇒ the
 * adapter behaves exactly as it did before this arc. */
export interface WsGuardOptions {
  guard?: IpGuard;
  /** Trusted proxy HOP COUNT — see clientIp in limits.ts, which documents
   * why this is a count and why reading the wrong end of X-Forwarded-For
   * produces a limiter anyone bypasses with one header. */
  trustProxyHops?: number;
  /** Per-socket outbound buffer ceiling; 0 disables. Over it, the socket is
   * terminated (`backpressure`). */
  maxBufferedBytes?: number;
  /** Global socket ceiling; 0 disables. Over it, new sockets are refused
   * (`server-busy`) — per-IP caps cannot bound a many-address botnet. */
  maxConns?: number;
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
  opts: WsGuardOptions = {},
): void {
  let nextId = 0;
  let liveSockets = 0;
  const guard = opts.guard ?? NO_OP_IP_GUARD;
  const maxBuffered = opts.maxBufferedBytes ?? 0;
  const maxConns = opts.maxConns ?? 0;
  wss.on("connection", (socket, req) => {
    let dead = false;
    const conn = {
      id: `ws${nextId++}`,
      send: (msg: ServerMessage) => {
        if (dead) return;
        // SLOW-CONSUMER BACKPRESSURE. A client can open a socket, join a busy
        // room, and simply never read: every broadcast still gets queued for
        // it, and Node holds all of it in THIS server's memory with no bound.
        // Nothing else notices — the peer is inside its connection limit, its
        // rate limit and its frame size; it is just not draining.
        //
        // Checked before the write rather than after, so the queue cannot grow
        // by one more fan-out past the ceiling, and enforced per socket at the
        // transport for the same reason the connection cap is: this is where
        // the memory is actually spent.
        //
        // `terminate` rather than a graceful close: a peer that is not reading
        // will not read a close handshake either. The cost is one reconnect —
        // their bundle is intact and rejoining is cheap.
        if (maxBuffered > 0 && (socket.bufferedAmount ?? 0) > maxBuffered) {
          dead = true;
          obs.kicked("backpressure");
          (socket.terminate ?? socket.close)?.call(socket);
          return;
        }
        socket.send(JSON.stringify(msg));
      },
    };
    // GLOBAL SOCKET CEILING. The per-IP cap below bounds one address; it does
    // nothing about a thousand addresses each opening a polite handful. Refuse
    // at accept — a new connection turned away beats an OOM that drops every
    // session already in progress.
    if (maxConns > 0 && liveSockets >= maxConns) {
      obs.capacityRefused("server-busy");
      try {
        socket.send(JSON.stringify({ t: "refused", reason: "server-busy" } satisfies ServerMessage));
      } catch {
        // Already gone; the close is what matters.
      }
      socket.close?.();
      return;
    }
    // PER-IP CONNECTION CAP, enforced HERE rather than at hello. The resource
    // a flood spends is the socket itself — an fd, a slot in the event loop,
    // a TLS session — and all of that is consumed before any protocol message
    // arrives. A check at hello would leave connect-and-say-nothing entirely
    // unbounded, which is the cheapest attack there is.
    //
    // Refused sockets are NOT counted as opened/closed: the gauge tracks
    // sockets this server is actually serving, and a connection hung up at
    // the door was never one. It is counted as an ip-limit instead, which is
    // where the signal belongs.
    const ip = clientIp(req, opts.trustProxyHops ?? 0);
    const admitted = guard.openConn(ip, conn.id);
    if (!admitted.ok) {
      obs.ipLimited(admitted.reason);
      // Tell the peer why before hanging up — an honest client can back off,
      // and a silent drop is indistinguishable from a network fault.
      try {
        socket.send(JSON.stringify({ t: "refused", reason: admitted.reason } satisfies ServerMessage));
      } catch {
        // The socket may already be gone; the close below is what matters.
      }
      socket.close?.();
      return;
    }
    liveSockets++;
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
      // Fires exactly once per admitted socket, including when `ws` closes it
      // itself for an oversized frame (status 1009, the maxPayload cap) or
      // when the backpressure guard terminated it — so every valve unwinds
      // through this one path and no counter can drift.
      dead = true;
      liveSockets = Math.max(0, liveSockets - 1);
      hub.disconnect(conn);
      guard.closeConn(conn.id);
      obs.connectionClosed();
    });
  });
}
