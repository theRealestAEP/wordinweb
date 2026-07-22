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
  constructor(private provider: DocProvider, private storage?: StorageDriver) {}

  /** Handle an inbound message from a connection. Async because a storage-
   * backed hello rehydrates and a submit persists before broadcast. */
  async handle(conn: Connection, msg: ClientMessage): Promise<void> {
    switch (msg.t) {
      case "hello": {
        if (msg.protocolVersion !== PROTOCOL_VERSION) {
          conn.send({ t: "refused", reason: "version-mismatch" });
          return;
        }
        const room = await this.room(msg.docId);
        room.conns.add(conn);
        this.connDoc.set(conn.id, msg.docId);
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
