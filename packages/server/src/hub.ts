import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "@wordinweb/collab/server";
import { ClientMessage, ServerMessage, PROTOCOL_VERSION } from "./protocol.js";

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

  constructor(private provider: DocProvider) {}

  /** Handle an inbound message from a connection. */
  handle(conn: Connection, msg: ClientMessage): void {
    switch (msg.t) {
      case "hello": {
        if (msg.protocolVersion !== PROTOCOL_VERSION) {
          conn.send({ t: "refused", reason: "version-mismatch" });
          return;
        }
        const room = this.room(msg.docId);
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
        const entry = room.session.submit(msg.intent);
        // Broadcast the sequenced (canonical) entry to every participant.
        const out: ServerMessage = { t: "broadcast", entries: [entry] };
        for (const c of room.conns) c.send(out);
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

  private room(docId: string): Room {
    let room = this.rooms.get(docId);
    if (!room) {
      const doc = DocxDocument.load(this.provider.load(docId));
      room = { session: new DocumentSession(doc), conns: new Set() };
      this.rooms.set(docId, room);
    }
    return room;
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  // btoa exists in Node ≥16 globals and browsers.
  return typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
}
