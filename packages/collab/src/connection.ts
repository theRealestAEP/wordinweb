import { ClientReplica } from "./replica.js";
import { Intent } from "./intents.js";
import { ClientMessage, ServerMessage, PROTOCOL_VERSION, PresencePosition } from "./protocol.js";

/**
 * A transport the connection drives: send a client message, and register a
 * handler for inbound server messages. A WebSocket adapter (or a fake, in
 * tests) implements it. The connection contains no socket code — the client
 * counterpart to the transport-free server hub (plan doc 07).
 */
export interface ClientTransport {
  send(msg: ClientMessage): void;
  onMessage(cb: (msg: ServerMessage) => void): void;
}

export interface ConnectionCallbacks {
  /** The document changed (welcome loaded, a broadcast applied). */
  onChange?: () => void;
  /** A remote participant's cursor/selection moved (or cleared: null). */
  onPresence?: (participant: string, position: PresencePosition | null) => void;
  /** The server refused the connection (e.g. version mismatch). */
  onRefused?: (reason: string) => void;
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = typeof atob === "function" ? atob(b64) : Buffer.from(b64, "base64").toString("binary");
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/**
 * Client-side session: joins a document, applies canonical broadcasts to a
 * ClientReplica (optimistic apply + reconciliation), and relays presence.
 * The editor layer submits local intents and reads `replica.doc`.
 */
export class CollabConnection {
  private replica: ClientReplica | null = null;
  private clientSeq = 0;
  /** Bumps only when reconciliation RELOADED the document (a true conflict) —
   * the renderer re-mounts on this, but updates in place otherwise (no flash
   * for the common non-conflicting edits). */
  docEpoch = 0;
  /** Per-client id allocator for carried node ids (split/format-range/insert):
   * a client-specific base block keeps concurrently-allocated ids disjoint
   * across clients so they never collide. */
  private idCounter = 0;
  private idBase = -1;

  constructor(
    private transport: ClientTransport,
    private clientId: string,
    private cb: ConnectionCallbacks = {},
  ) {
    this.transport.onMessage((msg) => this.onServer(msg));
  }

  /** Replace the callbacks (used by bindEditor to attach after construction). */
  setCallbacks(cb: ConnectionCallbacks): void {
    this.cb = cb;
  }

  /** Allocate `n` fresh carried node ids in this client's disjoint block. */
  allocIds(n: number): number[] {
    if (this.idBase < 0) {
      // A large per-client base derived from the client id, so two clients'
      // allocations never overlap (blocks of 1e7 ids each).
      let h = 0;
      for (let i = 0; i < this.clientId.length; i++) h = (h * 31 + this.clientId.charCodeAt(i)) >>> 0;
      this.idBase = 1_000_000_000 + (h % 100_000) * 10_000_000;
    }
    const out: number[] = [];
    for (let i = 0; i < n; i++) out.push(this.idBase + this.idCounter++);
    return out;
  }

  /** Join a document. The server replies with a welcome (snapshot + tail). */
  join(docId: string, token?: string): void {
    this.transport.send({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId, token, sinceSeq: 0 });
  }

  /** The live document (null until welcome). The editor renders this. */
  get doc() {
    return this.replica?.doc ?? null;
  }

  get ready(): boolean {
    return this.replica !== null;
  }

  /**
   * Submit a local edit. The caller supplies the intent minus its wire
   * bookkeeping (clientId/clientSeq/base) — the connection fills those from
   * the current confirmed seq and applies it optimistically before sending.
   */
  submit(intent: Omit<Intent, "clientId" | "clientSeq" | "base">): void {
    this.submitFull(intent, /*preApplied*/ false);
  }

  /** Submit an intent whose mutation the caller ALREADY performed on this
   * connection's live doc (the editor-driven path: DocxEditor applies the
   * command to `conn.doc` and then emits the intent). Skips the optimistic
   * re-apply — applying twice doubled every keystroke — but tracks pending and
   * sends identically, so echoes and reconciliation work unchanged. */
  submitPreApplied(intent: Omit<Intent, "clientId" | "clientSeq" | "base">): void {
    this.submitFull(intent, /*preApplied*/ true);
  }

  private submitFull(intent: Omit<Intent, "clientId" | "clientSeq" | "base">, preApplied: boolean): void {
    if (!this.replica) return;
    const full = {
      ...intent,
      clientId: this.clientId,
      clientSeq: ++this.clientSeq,
      base: this.replica.confirmedSeq,
    } as Intent;
    if (preApplied) this.replica.trackLocal(full);
    else this.replica.submitLocal(full);
    this.transport.send({ t: "submit", intent: full });
    this.cb.onChange?.();
  }

  /** Broadcast this client's cursor/selection (ephemeral). */
  setPresence(position: PresencePosition | null): void {
    this.transport.send({ t: "presence", position });
  }

  private onServer(msg: ServerMessage): void {
    switch (msg.t) {
      case "welcome": {
        // The snapshot represents seq `msg.seq`; the replica's confirmed
        // baseline is that seq, and any tail entries are strictly after it.
        this.replica = new ClientReplica(base64ToBytes(msg.snapshot));
        this.replica.confirmedSeq = msg.seq;
        if (msg.tail.length) this.replica.receive(msg.tail);
        this.cb.onChange?.();
        return;
      }
      case "broadcast": {
        this.replica?.receive(msg.entries);
        if (this.replica?.reloaded) this.docEpoch++;
        this.cb.onChange?.();
        return;
      }
      case "presence": {
        this.cb.onPresence?.(msg.participant, msg.position);
        return;
      }
      case "refused": {
        this.cb.onRefused?.(msg.reason);
        return;
      }
    }
  }
}
