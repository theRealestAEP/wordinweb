import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientMessage, ServerMessage, PROTOCOL_VERSION } from "../src/protocol.js";
import type { ClientTransport } from "../src/connection.js";

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return typeof btoa === "function" ? btoa(bin) : Buffer.from(bytes).toString("base64");
}

interface Peer {
  id: string;
  deliver: (msg: ServerMessage) => void;
  joined: boolean;
}

/**
 * A minimal in-process server for exercising CollabConnection end to end,
 * built from the collab primitives only (DocumentSession) — the collab tests
 * must not depend on @wordinweb/server (server → collab, not the reverse).
 * Synchronous: send() dispatches immediately, so tests read results inline.
 */
export class CollabHubLoopback {
  private session: DocumentSession;
  private peers: Peer[] = [];
  private nextId = 0;

  constructor(load: () => Uint8Array) {
    this.session = new DocumentSession(DocxDocument.load(load()));
  }

  connect(): ClientTransport {
    const peer: Peer = { id: `p${this.nextId++}`, deliver: () => {}, joined: false };
    this.peers.push(peer);
    return {
      send: (msg: ClientMessage) => this.handle(peer, msg),
      onMessage: (cb) => {
        peer.deliver = cb;
      },
    };
  }

  private handle(peer: Peer, msg: ClientMessage): void {
    switch (msg.t) {
      case "hello": {
        if (msg.protocolVersion !== PROTOCOL_VERSION) {
          peer.deliver({ t: "refused", reason: "version-mismatch" });
          return;
        }
        peer.joined = true;
        // Mirror the real hub (protocol v2): the welcome is a checkpoint
        // bundle — snapshot + id sidecar, tail strictly after the snapshot's
        // seq (a tail from sinceSeq with a current-doc snapshot double-applies
        // pre-join entries; the live-e2e suite caught the hub doing that).
        const cp = this.session.checkpoint();
        peer.deliver({
          t: "welcome",
          docId: msg.docId,
          seq: cp.seq,
          snapshot: bytesToBase64(cp.docx),
          sidecar: cp.sidecar,
          tail: this.session.entriesSince(cp.seq),
          genesisId: "g_loopback",
          mode: "plaintext",
        });
        return;
      }
      case "submit": {
        if (!peer.joined) return;
        const before = this.session.seq;
        const entry = this.session.submit(msg.intent);
        void before;
        for (const p of this.peers) if (p.joined) p.deliver({ t: "broadcast", entries: [entry] });
        return;
      }
      case "presence": {
        if (!peer.joined) return;
        for (const p of this.peers) if (p.joined && p !== peer) p.deliver({ t: "presence", participant: peer.id, position: msg.position });
        return;
      }
    }
  }
}
