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
import type { RosterEntry } from "./protocol.js";
import type { DocBundle } from "./bundle.js";
import { deriveEpochKeys, openCheckpoint, openIntent, sealIntent, sealCheckpoint, b64ToBytes, bytesToB64, type EpochKeys } from "./e2ee.js";
import { docHash } from "./hash.js";

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
  /** Own doc hashes at gossip points (seq → hash), a small ring — peers'
   * gossip for a seq we also hashed is comparable; anything else is skipped
   * (clients hash at the SAME seq cadence, so overlap is the common case). */
  private ownHashes = new Map<number, string>();
  /** Every K ingested seqs, hash + gossip (doc 13 §2). */
  private static GOSSIP_EVERY = 20;
  /** Server-designated checkpointer flag (doc 13 §3) + cadence. */
  private isCheckpointer = false;
  private static CHECKPOINT_EVERY = 50;
  /** The last intent this connection sealed+sent (one-in-flight makes this
   * the only candidate a `stale-base` refusal can refer to). */
  private lastSent: Intent | null = null;

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

  join(docId: string, token?: string, opts?: { takeover?: boolean; profile?: ParticipantProfile; codeProof?: string; ownerToken?: string }): void {
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
      codeProof: opts?.codeProof,
      ownerToken: opts?.ownerToken, // owner-capability proof (doc 14 §2.5)
      engineVersion: ENGINE_VERSION, // the fence for client-derived canon (doc 13 §2)
    });
  }

  get doc() {
    return this.replica?.doc ?? null;
  }
  get ready(): boolean {
    return this.replica !== null;
  }

  /** Roster + activity parity with the plaintext connection (doc 14). */
  roster: RosterEntry[] = [];
  activity: { seq: number; clientId: string; kind: string }[] = [];

  setProfile(profile: ParticipantProfile): void {
    this.transport.send({ t: "profile", profile } as ClientMessage);
  }

  /** Bundle being resumed from (consumed after the welcome replay). */
  private resuming: DocBundle | null = null;

  /**
   * Rejoin from a persisted bundle (doc 12 §5, encrypted flavor). Identical
   * contract to the plaintext connection: clientSeq watermark restored
   * FIRST; pending replayed fire-and-observe on same-epoch welcomes (the
   * sequencer's plaintext-bookkeeping dedup gives exactly-once, and the
   * whole-epoch welcome-enc replay already reconstructed everything that
   * was sequenced pre-crash); different epoch ⇒ onEpochChange, pending
   * withheld (fork rule). The bundle's confirmed bytes are NOT loaded here
   * — encrypted joiners always rebuild from the sealed checkpoint + tail,
   * which is both simpler and verifiable (hash gossip).
   */
  resume(bundle: DocBundle, token?: string, opts?: { profile?: ParticipantProfile; codeProof?: string; ownerToken?: string }): void {
    this.clientSeq = Math.max(this.clientSeq, bundle.clientSeq);
    this.resuming = bundle;
    this.join(bundle.docId, token, { ...opts, takeover: true });
  }

  /** Durable state as a doc-12 bundle (kDoc rides in it for revival). */
  exportBundle(docId: string): DocBundle | null {
    if (!this.replica || this.genesisId === null) return null;
    return {
      docId,
      genesisId: this.genesisId,
      ...this.replica.exportBundleState(),
      clientSeq: this.clientSeq,
      savedAt: 0,
      lineage: [], // the persister maintains the chain across writes
    };
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
      this.lastSent = full;
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
          this.mirror.setSeqFloor(msg.checkpoint.seq); // numbering continues from the checkpoint
          if (cp.sidecar) this.mirror.installSidecar(cp.sidecar as never);
          this.replica = new ClientReplica(bytes, (cp.sidecar ?? undefined) as never);
          this.replica.confirmedSeq = msg.checkpoint.seq;
          for (const env of msg.tail) await this.ingest(env);
          // Resume epilogue — same semantics as the plaintext connection.
          const resumed = this.resuming;
          this.resuming = null;
          if (resumed) {
            if (resumed.genesisId === msg.genesisId) {
              for (const intent of resumed.pending) {
                const env2 = await sealIntent(this.keys!.kContent, this.docId, msg.genesisId, intent as Intent);
                this.transport.send({ t: "submit-enc", envelope: env2 } as ClientMessage);
              }
            } else {
              this.cb.onEpochChange?.(resumed.genesisId, msg.genesisId);
            }
          }
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
        this.roster = msg.roster;
        this.cb.onRoster?.(msg.roster);
        return;
      }
      case "gossip": {
        this.enqueue(() => this.onGossip(msg.iv, msg.ciphertext));
        return;
      }
      case "checkpointer": {
        this.isCheckpointer = msg.active;
        return;
      }
      case "refused": {
        if (msg.reason === "stale-base") {
          // The server adopted a checkpoint past our envelope's base
          // (quiescent-checkpoint guard). By the time this refusal arrives
          // the broadcasts up to that checkpoint have been processed, so
          // re-seal the SAME intent (same clientId/clientSeq — dedup-safe)
          // with a fresh base and resubmit. One-in-flight discipline makes
          // lastSent the only possible referent.
          const retry = this.lastSent;
          if (retry && this.replica && this.keys && this.genesisId) {
            this.enqueue(async () => {
              const rebased = { ...retry, base: this.replica!.confirmedSeq } as Intent;
              const env = await sealIntent(this.keys!.kContent, this.docId, this.genesisId!, rebased);
              this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
            });
          }
          return;
        }
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
    if (entry.kind === "applied") {
      this.activity.push({ seq: entry.seq, clientId: entry.intent.clientId, kind: entry.intent.kind });
      if (this.activity.length > 100) this.activity.splice(0, this.activity.length - 100);
    }

    // Checkpoint duty (doc 13 §3): the designated client uploads its
    // mirror's canonical state every K seqs. The mirror is confirmed-only
    // by construction (no optimistic edits ever touch it), satisfying the
    // round-4 F17 rule structurally; the server adopts it only if it is
    // quiescent (exactly at the log head) — a miss just retries next round.
    if (this.isCheckpointer && env.seq % EncryptedCollabConnection.CHECKPOINT_EVERY === 0) {
      const cp = this.mirror!.checkpoint();
      const sealed = await sealCheckpoint(this.keys!.kContent, this.docId, this.genesisId!, cp.seq, {
        docx: bytesToB64(cp.docx),
        sidecar: cp.sidecar,
        docHash: await docHash(this.mirror!.doc),
      });
      this.transport.send({ t: "checkpoint", checkpoint: { seq: cp.seq, ...sealed } } as ClientMessage);
    }

    // Divergence DETECTION (doc 13 §2 / round-4 blocker 3): every K seqs,
    // hash the mirror's canonical doc, remember it, and gossip it sealed.
    // Every honest client hashes at the same seq points, so peers' blobs
    // are directly comparable; a mismatch means an engine bug or a
    // tampered history — surfaced as `divergence` (the resync path:
    // rejoin from the welcome, which is always available).
    if (env.seq % EncryptedCollabConnection.GOSSIP_EVERY === 0) {
      const hash = await docHash(this.mirror!.doc);
      this.ownHashes.set(env.seq, hash);
      if (this.ownHashes.size > 5) {
        this.ownHashes.delete(Math.min(...this.ownHashes.keys()));
      }
      const iv = new Uint8Array(12);
      crypto.getRandomValues(iv);
      const body = new TextEncoder().encode(JSON.stringify({ seq: env.seq, hash }));
      const aad = new TextEncoder().encode(`gs:${this.docId}:${this.genesisId}`);
      const ct = await crypto.subtle.encrypt(
        { name: "AES-GCM", iv: iv as BufferSource, additionalData: aad as BufferSource },
        (this.keys as unknown as { kContent: CryptoKey }).kContent,
        body as BufferSource,
      );
      this.transport.send({ t: "gossip", iv: bytesToB64(iv), ciphertext: bytesToB64(new Uint8Array(ct)) } as ClientMessage);
    }
  }

  /** Verify a peer's gossiped hash against our own at the same seq. */
  private async onGossip(iv: string, ciphertext: string): Promise<void> {
    if (!this.keys || !this.genesisId) return;
    let claim: { seq: number; hash: string };
    try {
      const aad = new TextEncoder().encode(`gs:${this.docId}:${this.genesisId}`);
      const pt = await crypto.subtle.decrypt(
        { name: "AES-GCM", iv: b64ToBytes(iv) as BufferSource, additionalData: aad as BufferSource },
        (this.keys as unknown as { kContent: CryptoKey }).kContent,
        b64ToBytes(ciphertext) as BufferSource,
      );
      claim = JSON.parse(new TextDecoder().decode(pt));
    } catch {
      return; // garbage gossip: ignorable (unlike intents, it holds no seq)
    }
    const own = this.ownHashes.get(claim.seq);
    if (own !== undefined && own !== claim.hash) {
      // Two honest clients disagreeing on canonical state at the same seq:
      // the round-4 blocker-3 detection case. Telemetry-worthy always;
      // recovery is a rejoin (fresh welcome-enc replay).
      this.cb.onRefused?.("divergence");
    }
  }
}
