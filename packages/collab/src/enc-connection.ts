import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "./session.js";
import { ClientReplica } from "./replica.js";
import { Intent } from "./intents.js";
import {
  ClientMessage,
  ServerMessage,
  EnvelopeEntry,
  PROTOCOL_VERSION,
  ENGINE_VERSION,
  PresencePosition,
  ParticipantProfile,
} from "./protocol.js";
import type { ClientTransport, ConnectionCallbacks } from "./connection.js";
import { deriveEpochKeys, openCheckpoint, openIntent, sealIntent, b64ToBytes, type EpochKeys } from "./e2ee.js";

/**
 * The E2EE client connection (plan doc 13 §2) — the counterpart of a BLIND
 * sequencer. The server orders opaque envelopes; this class re-derives the
 * canonical history the plaintext server would have produced, using a local
 * **mirror**: a real `DocumentSession` — the exact authority code the
 * plaintext server runs — fed the DECRYPTED originals in server order.
 *
 * Why this shape: the mirror's `submit()` performs the same
 * transform-against-concurrent → validate → apply → log pipeline the server
 * does, and since every client feeds it the identical ordered inputs, every
 * client derives the identical canonical entries ("the transform moves, the
 * math doesn't change" — made literal: it is the same function, the same
 * class). Those canonical entries then drive an unchanged `ClientReplica`,
 * so optimistic apply, rollback-replay, echoes, and rejection handling are
 * ALL the plaintext code paths — the encryption layer adds sealing at the
 * edges and nothing else. Convergence in encrypted mode therefore inherits
 * every plaintext convergence test by construction.
 *
 * Undecryptable envelopes (malicious participant, tampered blob) consume
 * their seq as a deterministic rejection on every honest client
 * (`ingestOpaqueFailure`) — GCM authentication is deterministic, so all
 * mirrors agree byte-for-byte on the no-op.
 *
 * Mode is derived from the LINK, never the wire (doc 13 §6): this class
 * exists because the caller holds `#k` — a plaintext `welcome` arriving
 * anyway is a downgrade attempt and is hard-refused.
 */
export class EncryptedCollabConnection {
  private mirror: DocumentSession | null = null;
  private replica: ClientReplica | null = null;
  private keys: EpochKeys | null = null;
  private clientSeq = 0;
  private docId = "";
  genesisId: string | null = null;
  docEpoch = 0;
  /** Envelope processing is async (WebCrypto) but MUST be strictly ordered
   * (seq order is the convergence contract) — a serial promise chain. */
  private queue: Promise<void> = Promise.resolve();
  private idCounter = 0;
  private idBase = -1;

  constructor(
    private transport: ClientTransport,
    private clientId: string,
    /** The document master key — the `#k=` fragment value (doc 13 §1). */
    private docKey: string,
    private cb: ConnectionCallbacks = {},
    /** Stretched share code (doc 13 §7) when the doc has one. */
    private stretchedCode?: Uint8Array,
  ) {
    this.transport.onMessage((msg) => this.onServer(msg));
  }

  join(docId: string, token?: string, opts?: { takeover?: boolean; profile?: ParticipantProfile }): void {
    this.docId = docId;
    this.transport.send({
      t: "hello",
      protocolVersion: PROTOCOL_VERSION,
      docId,
      clientId: this.clientId,
      takeover: opts?.takeover,
      token,
      sinceSeq: 0,
      profile: opts?.profile,
      engineVersion: ENGINE_VERSION, // the fence for client-derived canon (doc 13 §2)
    });
  }

  get doc() {
    return this.replica?.doc ?? null;
  }
  get ready(): boolean {
    return this.replica !== null;
  }

  /** Same disjoint carried-id allocation as the plaintext connection. */
  allocIds(n: number): number[] {
    if (this.idBase < 0) {
      let h = 0;
      for (let i = 0; i < this.clientId.length; i++) h = (h * 31 + this.clientId.charCodeAt(i)) >>> 0;
      this.idBase = 1_000_000_000 + (h % 100_000) * 10_000_000;
    }
    const out: number[] = [];
    for (let i = 0; i < n; i++) out.push(this.idBase + this.idCounter++);
    return out;
  }

  /** Submit a local edit (optimistic canonical apply, like the plaintext
   * `submit`); sealing happens on the ordered queue. */
  submit(intent: Omit<Intent, "clientId" | "clientSeq" | "base">): void {
    this.submitFull(intent, false);
  }
  /** Editor-driven path: the mutation is already in the live doc. */
  submitPreApplied(intent: Omit<Intent, "clientId" | "clientSeq" | "base">): void {
    this.submitFull(intent, true);
  }

  private submitFull(intent: Omit<Intent, "clientId" | "clientSeq" | "base">, preApplied: boolean): void {
    if (!this.replica || !this.keys || !this.genesisId) return;
    const full = {
      ...intent,
      clientId: this.clientId,
      clientSeq: ++this.clientSeq,
      base: this.replica.confirmedSeq,
    } as Intent;
    if (preApplied) this.replica.trackLocal(full);
    else this.replica.submitLocal(full);
    // Seal + send ride the serial queue so a fast second keystroke can't
    // overtake the first one's async encryption (order = correctness).
    this.enqueue(async () => {
      const env = await sealIntent(this.keys!.kContent, this.docId, this.genesisId!, full);
      this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
    });
    this.cb.onChange?.();
  }

  setPresence(position: PresencePosition | null): void {
    this.transport.send({ t: "presence", position });
  }

  private enqueue(task: () => Promise<void>): void {
    this.queue = this.queue.then(task).catch(() => {
      // A failed task must not wedge the chain; individual handlers report
      // their own failures (an unopenable envelope is HANDLED, not thrown).
    });
  }

  private onServer(msg: ServerMessage): void {
    switch (msg.t) {
      case "welcome-enc": {
        this.enqueue(async () => {
          this.genesisId = msg.genesisId;
          this.keys = await deriveEpochKeys(this.docKey, msg.genesisId, this.stretchedCode);
          const cp = await openCheckpoint(this.keys.kContent, msg.docId, msg.genesisId, msg.checkpoint.seq, msg.checkpoint);
          const bytes = b64ToBytes(cp.docx);
          // The mirror is the local blind-mode "server": seeded exactly as a
          // rehydrating plaintext server would be — bytes, then sidecar,
          // then the tail replayed through the canonical pipeline.
          this.mirror = new DocumentSession(DocxDocument.load(bytes));
          this.mirror.installSidecar(cp.sidecar as never);
          this.replica = new ClientReplica(bytes, cp.sidecar as never);
          this.replica.confirmedSeq = msg.checkpoint.seq;
          for (const env of msg.tail) await this.ingest(env);
          this.cb.onChange?.();
        });
        return;
      }
      case "broadcast-enc": {
        this.enqueue(async () => {
          if (!this.replica) return; // pre-welcome broadcasts are impossible; guard anyway
          for (const env of msg.entries) await this.ingest(env);
          if (this.replica.reloaded) this.docEpoch++;
          this.cb.onChange?.();
        });
        return;
      }
      case "welcome": {
        // We hold the key, so this doc IS encrypted — a plaintext welcome is
        // a downgrade attempt (malicious server or hostile plaintext
        // re-seed) and is never obeyed (doc 13 §6, round-4 F11).
        this.cb.onRefused?.("mode-downgrade");
        return;
      }
      case "presence": {
        this.cb.onPresence?.(msg.participant, msg.position);
        return;
      }
      case "roster": {
        this.cb.onRoster?.(msg.roster);
        return;
      }
      case "refused": {
        this.cb.onRefused?.(msg.reason);
        return;
      }
    }
  }

  /** Feed one sequenced envelope through mirror + replica (in seq order). */
  private async ingest(env: EnvelopeEntry): Promise<void> {
    let entry;
    try {
      const intent = await openIntent(this.keys!.kContent, this.docId, this.genesisId!, env);
      entry = this.mirror!.submit(intent);
    } catch {
      // Deterministic no-op: unopenable for us ⇒ unopenable for every
      // honest client (GCM auth is a pure function of bytes + key).
      entry = this.mirror!.ingestOpaqueFailure(env.clientId, env.clientSeq);
    }
    // The mirror re-derives seqs in arrival order; they must match the
    // sequencer's. A mismatch means entries were dropped/reordered in
    // transit — surface as a refusal (the resync story, doc 13 §2).
    if (entry.seq !== env.seq) {
      this.cb.onRefused?.("sequence-desync");
      return;
    }
    this.replica!.receive([entry]);
  }
}
