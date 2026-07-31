/**
 * Wire tap for the perf HUD — demo-layer only, metrics only.
 *
 * WHY A TAP AND NOT A PATCHED SESSION: the HUD must measure what the USER
 * does (typing → `session.submit`), not only what the E2E hook does
 * (`__ww.submitOp`). The session object the React binding hands to the app is
 * a fresh literal per render whose methods close over a private connection, so
 * there is nothing stable to observe from the demo layer. The SOCKET, however,
 * carries every intent both ways with its bookkeeping in the CLEAR — even in
 * E2EE mode (doc 13 §2: the blind sequencer needs `clientId`/`clientSeq`/`base`
 * to order envelopes, so only the intent BODY is sealed). That makes exact
 * pending counts and exact submit→own-echo round-trips readable without
 * touching a single line of the collab packages.
 *
 * ZERO CUSTODY: this file reads `t`, `clientId`, `clientSeq` and `seq` off the
 * envelopes and NOTHING else — never `ciphertext`, never a plaintext intent
 * body, never document text. Frame size is taken from the raw string length.
 *
 * ZERO COST WHEN HIDDEN: importing this module installs a WebSocket subclass
 * that does one thing per socket construction — add it to a registry. No
 * listener, no wrapper, no work per frame. Listeners and the `send` wrapper
 * are installed by `attachTap()` when the HUD opens and fully removed by
 * `detachTap()` when it closes.
 */

/** What the HUD needs from one outbound submit. */
export interface OutFrame {
  clientSeq: number;
  bytes: number;
  at: number;
}

/** What the HUD needs from one inbound sequenced entry. */
export interface InEntry {
  seq: number;
  clientId: string;
  clientSeq: number;
}

export interface TapHandlers {
  onSubmit(frame: OutFrame): void;
  /** One broadcast/welcome message: its entries plus the raw frame size. */
  onEntries(entries: InEntry[], bytes: number, at: number): void;
}

const sockets = new Set<WebSocket>();
let handlers: TapHandlers | null = null;
const wrapped = new WeakSet<WebSocket>();
const listeners = new WeakMap<WebSocket, (e: MessageEvent) => void>();

/** Install the registry (idempotent). Called at module import from main.tsx —
 * must run before the collab session constructs its socket. */
export function installWsRegistry(): void {
  const w = window as unknown as { __wwWsTapped?: boolean; WebSocket: typeof WebSocket };
  if (w.__wwWsTapped) return;
  w.__wwWsTapped = true;
  const Native = w.WebSocket;
  class TappedWebSocket extends Native {
    constructor(url: string | URL, protocols?: string | string[]) {
      super(url, protocols);
      sockets.add(this);
      this.addEventListener("close", () => {
        sockets.delete(this);
        wrapped.delete(this);
      });
      if (handlers) equip(this);
    }
  }
  w.WebSocket = TappedWebSocket as unknown as typeof WebSocket;
}

/** Read the bookkeeping the blind sequencer already sees. Returns null for
 * every frame that isn't an intent submit (presence, gossip, hello, …). */
function readSubmit(data: unknown): { clientSeq: number } | null {
  if (typeof data !== "string" || data.length === 0 || data[0] !== "{") return null;
  try {
    const msg = JSON.parse(data) as { t?: string; envelope?: { clientSeq?: number }; intent?: { clientSeq?: number } };
    if (msg.t === "submit-enc" && msg.envelope) return { clientSeq: msg.envelope.clientSeq ?? -1 };
    if (msg.t === "submit" && msg.intent) return { clientSeq: msg.intent.clientSeq ?? -1 };
    return null;
  } catch {
    return null;
  }
}

/** Sequenced entries out of a broadcast (encrypted or plaintext). */
function readEntries(data: unknown): InEntry[] | null {
  if (typeof data !== "string" || data.length === 0 || data[0] !== "{") return null;
  try {
    const msg = JSON.parse(data) as {
      t?: string;
      entries?: { seq?: number; clientId?: string; clientSeq?: number; intent?: { clientId?: string; clientSeq?: number } }[];
    };
    if (msg.t !== "broadcast-enc" && msg.t !== "broadcast") return null;
    return (msg.entries ?? []).map((e) => ({
      seq: e.seq ?? 0,
      // Encrypted: the envelope carries the bookkeeping. Plaintext: it rides
      // on the (readable) intent. We take only these two fields either way.
      clientId: e.clientId ?? e.intent?.clientId ?? "",
      clientSeq: e.clientSeq ?? e.intent?.clientSeq ?? -1,
    }));
  } catch {
    return null;
  }
}

function equip(socket: WebSocket): void {
  if (wrapped.has(socket)) return;
  wrapped.add(socket);
  const nativeSend = socket.send.bind(socket);
  // Own-property wrapper: `delete socket.send` restores the prototype method
  // exactly, so detaching leaves the socket byte-for-byte as it was.
  Object.defineProperty(socket, "send", {
    configurable: true,
    writable: true,
    value: (data: string | ArrayBufferLike | Blob | ArrayBufferView) => {
      const submit = readSubmit(data);
      if (submit && handlers) {
        handlers.onSubmit({
          clientSeq: submit.clientSeq,
          bytes: typeof data === "string" ? data.length : 0,
          at: performance.now(),
        });
      }
      nativeSend(data as never);
    },
  });
  const onMessage = (e: MessageEvent) => {
    if (!handlers) return;
    const at = performance.now();
    const entries = readEntries(e.data);
    if (entries && entries.length) {
      handlers.onEntries(entries, typeof e.data === "string" ? e.data.length : 0, at);
    }
  };
  listeners.set(socket, onMessage);
  // Capture phase so the sample is taken before the connection's own handler
  // spends time decrypting — the arrival timestamp is the honest one.
  socket.addEventListener("message", onMessage, true);
}

function unequip(socket: WebSocket): void {
  if (!wrapped.has(socket)) return;
  wrapped.delete(socket);
  delete (socket as unknown as Record<string, unknown>).send;
  const l = listeners.get(socket);
  if (l) {
    socket.removeEventListener("message", l, true);
    listeners.delete(socket);
  }
}

/** Start observing every live (and subsequently created) socket. */
export function attachTap(h: TapHandlers): void {
  handlers = h;
  for (const s of sockets) equip(s);
}

/** Stop observing — removes every listener and wrapper this module added. */
export function detachTap(): void {
  handlers = null;
  for (const s of sockets) unequip(s);
}
