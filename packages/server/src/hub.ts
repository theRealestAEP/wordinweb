import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "@wordinweb/collab/server";
import { ClientMessage, ServerMessage, PROTOCOL_VERSION } from "@wordinweb/collab/server";
import { StorageDriver } from "./storage.js";

/**
 * A connection the hub can send to, abstracting the transport (WebSocket in
 * production, a fake in tests). `id` identifies the socket; `send` delivers a
 * server message.
 */
export interface Connection {
  id: string;
  send(msg: ServerMessage): void;
}

/**
 * How the hub obtains a document to start a session. Production wires this to
 * a storage driver (snapshot + log, plan doc 06); tests supply bytes inline.
 */
export interface DocProvider {
  /** Load the initial document bytes for a docId (a fresh blank doc, a stored
   * snapshot, etc.). */
  load(docId: string): Uint8Array;
}

/**
 * Verified identity/authorization for a connection (plan doc 06/07). The app's
 * backend mints a JWT; the hub only VERIFIES via an injected verifier and
 * enforces that the token authorizes the requested docId. Roles gate writes at
 * the sequencer. The hub never trusts a claimed identity — that is the
 * validate-then-install / cross-tenant-isolation rule (doc 11).
 */
export interface AuthResult {
  userId: string;
  role: "editor" | "viewer";
  /** Token expiry in epoch ms (from the JWT `exp`). When set, the hub enforces
   * it on the LIVE socket — not just at connect — so a short-lived token
   * actually limits the session (plan doc 06/11 gate 5). Omit for no expiry. */
  expiresAt?: number;
}
export interface TokenVerifier {
  /** Verify a token for a docId. Return the authorized identity, or null to
   * refuse the connection. Must reject a token whose docId does not match. */
  verify(token: string | undefined, docId: string): AuthResult | null;
}

interface Room {
  session: DocumentSession;
  conns: Set<Connection>;
}

/**
 * Routes clients to per-document sessions and fans out broadcasts. This is the
 * thin plumbing layer (plan doc 06): it owns no editing logic — the
 * DocumentSession is the authority — and no transport, so it is unit-testable
 * with in-memory connections.
 */
export class CollabHub {
  private rooms = new Map<string, Room>();
  private connDoc = new Map<string, string>();

  /** Optional persistence (plan doc 06). Without a driver the hub is purely
   * in-memory (a session lives only while a room is held). With one, rooms
   * rehydrate from snapshot+log on first access and every sequenced entry is
   * durably appended before broadcast. */
  private auth = new Map<string, AuthResult>();
  private connById = new Map<string, Connection>();

  constructor(
    private provider: DocProvider,
    private storage?: StorageDriver,
    /** Optional token verifier. When set, every hello must present a token
     * that authorizes the requested docId; without it the hub is auth-off
     * (dev/demo only). */
    private verifier?: TokenVerifier,
    /** Injectable clock (epoch ms) for token-expiry checks; defaults to
     * Date.now. Injected in tests for determinism. */
    private now: () => number = () => Date.now(),
  ) {}

  /** True if the connection's token has expired (live-socket enforcement). */
  private expired(connId: string): boolean {
    const a = this.auth.get(connId);
    return a?.expiresAt !== undefined && a.expiresAt <= this.now();
  }

  /** Force-disconnect a connection with a refusal (expiry/revocation). */
  private kick(conn: Connection, reason: string): void {
    conn.send({ t: "refused", reason });
    this.disconnect(conn);
  }

  /**
   * Revoke all live sessions for a user (the app's backend calls this when it
   * deprovisions/downgrades someone — plan doc 06 gate 5). They are refused
   * and disconnected immediately, not at next reconnect.
   */
  revoke(userId: string): void {
    for (const [connId, a] of [...this.auth]) {
      if (a.userId === userId) {
        const conn = this.connById.get(connId);
        if (conn) this.kick(conn, "revoked");
      }
    }
  }

  /** Sweep expired sessions (call periodically from the transport). */
  sweepExpired(): void {
    for (const [connId] of [...this.auth]) {
      if (this.expired(connId)) {
        const conn = this.connById.get(connId);
        if (conn) this.kick(conn, "token-expired");
      }
    }
  }

  /** Handle an inbound message from a connection. Async because a storage-
   * backed hello rehydrates and a submit persists before broadcast. */
  async handle(conn: Connection, msg: ClientMessage): Promise<void> {
    switch (msg.t) {
      case "hello": {
        if (msg.protocolVersion !== PROTOCOL_VERSION) {
          conn.send({ t: "refused", reason: "version-mismatch" });
          return;
        }
        if (this.verifier) {
          const authed = this.verifier.verify(msg.token, msg.docId);
          if (!authed) {
            conn.send({ t: "refused", reason: "unauthorized" });
            return;
          }
          this.auth.set(conn.id, authed);
        }
        const room = await this.room(msg.docId);
        room.conns.add(conn);
        this.connDoc.set(conn.id, msg.docId);
        this.connById.set(conn.id, conn);
        conn.send({
          t: "welcome",
          docId: msg.docId,
          seq: room.session.seq,
          snapshot: bytesToBase64(room.session.doc.save()),
          tail: room.session.entriesSince(Math.max(0, msg.sinceSeq)),
        });
        return;
      }
      case "submit": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) {
          conn.send({ t: "refused", reason: "not-joined" });
          return;
        }
        // Live-socket token-expiry enforcement (gate 5): a session outlives
        // its token only until the next action, then is cut.
        if (this.expired(conn.id)) {
          this.kick(conn, "token-expired");
          return;
        }
        // Role enforcement at the sequencer (doc 06): a viewer's edits are
        // refused, not merely hidden in the UI.
        if (this.verifier && this.auth.get(conn.id)?.role !== "editor") {
          conn.send({ t: "refused", reason: "read-only" });
          return;
        }
        const room = this.rooms.get(docId)!;
        const before = room.session.seq;
        const entry = room.session.submit(msg.intent);
        // A deduplicated re-send returns an already-sequenced entry (seq <=
        // before); persist only genuinely new entries so the log stays
        // append-once. Durability before ack (plan doc 06), then broadcast.
        if (this.storage && entry.seq > before) await this.storage.appendEntries(docId, [entry]);
        const out: ServerMessage = { t: "broadcast", entries: [entry] };
        for (const c of room.conns) c.send(out);
        return;
      }
      case "presence": {
        const docId = this.connDoc.get(conn.id);
        if (!docId) return; // presence before join is ignored, not refused
        const room = this.rooms.get(docId)!;
        // Ephemeral: fan out to every OTHER participant, never logged/persisted.
        const out: ServerMessage = { t: "presence", participant: conn.id, position: msg.position };
        for (const c of room.conns) if (c.id !== conn.id) c.send(out);
        return;
      }
    }
  }

  /** Drop a connection (socket closed). */
  disconnect(conn: Connection): void {
    const docId = this.connDoc.get(conn.id);
    this.connDoc.delete(conn.id);
    this.auth.delete(conn.id);
    this.connById.delete(conn.id);
    if (docId) this.rooms.get(docId)?.conns.delete(conn);
  }

  /** Sessions currently held in memory (for eviction/metrics). */
  activeDocs(): string[] {
    return [...this.rooms.keys()];
  }

  private async room(docId: string): Promise<Room> {
    let room = this.rooms.get(docId);
    if (!room) {
      const session = await this.startSession(docId);
      room = { session, conns: new Set() };
      this.rooms.set(docId, room);
    }
    return room;
  }

  private async startSession(docId: string): Promise<DocumentSession> {
    if (this.storage) {
      const snap = await this.storage.loadSnapshot(docId);
      if (snap) {
        const session = new DocumentSession(DocxDocument.load(snap.docx));
        session.loadCanonical(await this.storage.readLog(docId, snap.seq));
        return session;
      }
      // No snapshot yet: start fresh but replay any persisted log tail.
      const session = new DocumentSession(DocxDocument.load(this.provider.load(docId)));
      session.loadCanonical(await this.storage.readLog(docId, 0));
      return session;
    }
    return new DocumentSession(DocxDocument.load(this.provider.load(docId)));
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  // btoa exists in Node ≥16 globals and browsers.
  return typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
}
