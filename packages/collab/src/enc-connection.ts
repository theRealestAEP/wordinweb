import { DocxDocument } from "@wordinweb/core";
import { DocumentSession, type UndoOutcome } from "./session.js";
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
  sanitizePresencePosition,
} from "./protocol.js";
import type { ClientTransport, ConnectionCallbacks } from "./connection.js";
import type { Scope } from "./apply.js";
import type { RosterEntry } from "./protocol.js";
import type { DocBundle } from "./bundle.js";
import { deriveEpochKeys, openCheckpoint, openIntent, sealIntent, sealCheckpoint, sealMediaBlob, openMediaBlob, sealPresence, openPresence, isSealedPresence, b64ToBytes, bytesToB64, type EpochKeys } from "./e2ee.js";
import { docHash } from "./hash.js";
import { transformIntent } from "./transform.js";
import { MediaClient, applyMediaAddresses, mediaAddressesOf, type MediaTransportOptions } from "./media.js";

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
  /** Relay per-blob upload limit from the welcome, or `null` when the server
   * published none (an older host) — in which case callers SKIP the
   * pre-check rather than inventing a limit. See the plaintext connection's
   * field of the same name. */
  mediaMaxBlobBytes: number | null = null;
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
  /** SELF-HEAL (the B6a typist-drift catch): at quiescence the optimistic
   * replica must equal the mirror's canonical doc — the mirror is LOCAL
   * ground truth (the same authority code the plaintext server runs, fed
   * the identical ordered inputs), so drift in the optimistic
   * rollback-replay machinery is detectable and repairable without a
   * network round-trip. Debounced past the drain; staleness-guarded. */
  private selfCheckTimer: ReturnType<typeof setTimeout> | null = null;
  private lastSelfCheckSeq = -1;
  /** Number of times the replica was rebuilt from the mirror (telemetry). */
  selfHeals = 0;
  /** STUCK-PENDING RECOVERY (swarm finding): under load a submit can be
   * lost between pending-tracking and delivery (or its echo lost), leaving
   * the client permanently un-quiescent — which also blocks the drift
   * self-check. The front pending is watched; if it makes no progress, the
   * RECONCILIATION-TRANSFORMED copy is re-sealed at the current base and
   * resent (server dedup by (clientId, clientSeq) makes a resend of an
   * already-sequenced op a no-op). After MAX retries the pending is dropped
   * and the replica heals from the mirror so the room converges. */
  private stuckTimer: ReturnType<typeof setTimeout> | null = null;
  private stuckKey: string | null = null;
  private stuckRetries = 0;
  private static STUCK_MAX_RETRIES = 3;
  /** Rate-limit recovery (see the "refused" handler): one scheduled
   * full-queue re-drive at a time, exponential backoff 300ms→5s, reset
   * when the queue drains. */
  private redriveTimer: ReturnType<typeof setTimeout> | null = null;
  private redriveBackoffMs = 0;

  /** Out-of-band media transfer (doc 16 §5), sealed under K_media. */
  media: MediaClient | null = null;

  constructor(
    private transport: ClientTransport,
    private clientId: string,
    /** The document master key — the `#k=` fragment value (doc 13 §1). */
    private docKey: string,
    private cb: ConnectionCallbacks = {},
    /** Stretched share code (doc 13 §7) when the doc has one. */
    private stretchedCode?: Uint8Array,
    /** Test seam: self-check debounce delay (default 400ms). */
    private selfCheckDelayMs = 400,
    /** Origin serving the doc-16 §3 media routes; enables media transfer. */
    mediaOpts?: MediaTransportOptions,
  ) {
    this.transport.onMessage((msg) => this.onServer(msg));
    if (mediaOpts) {
      // E2EE media (doc 16 §6): blobs are sealed under K_media, a key
      // DISTINCT from K_content, and the relay only ever sees ciphertext —
      // the sha it verifies addresses that ciphertext, so a blind server can
      // still swap-proof uploads. The IV is returned by the seal and rides
      // in the intent, because re-supply must reproduce the SAME ciphertext.
      this.media = new MediaClient(
        mediaOpts,
        () => this.replica?.doc ?? null,
        {
          // Media installs mutate the doc (pixels land) outside the replica's
          // counter — bump the render signal so the skeleton repaints.
          onChange: () => { this.docVersionBase++; this.cb.onChange?.(); },
          onState: (part, state) => this.cb.onMediaState?.(part, state),
          need: (sha) => this.transport.send({ t: "media-need", sha }),
          have: (shas) => { if (shas.length) this.transport.send({ t: "media-have", shas }); },
        },
        {
          seal: async (plaintext, iv) => {
            if (!this.keys) throw new Error("media seal before keys");
            return sealMediaBlob(this.keys.kMedia, plaintext, iv);
          },
          open: async (blob, iv) => {
            if (!this.keys) throw new Error("media open before keys");
            return openMediaBlob(this.keys.kMedia, blob, iv);
          },
        },
      );
    }
  }

  /** Upload bytes and get the intent's media fields (blobSha over the
   * CIPHERTEXT, plus the IV the re-supply path must reuse), or null when the
   * relay refused — the caller must not emit an insertImage in that case. */
  async uploadMedia(plaintext: Uint8Array): Promise<{ blobSha: string; bytesLen: number; iv?: string } | null> {
    return this.media ? this.media.upload(this.docId, plaintext) : null;
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

  /** Accumulates doc changes across replica rebuilds (self-heal/welcome), so
   * the getter below stays monotonic per connection. */
  private docVersionBase = 0;
  /**
   * Counts changes TO THE RENDERED DOCUMENT — the repaint signal, as opposed
   * to onChange (which also fires for pure bookkeeping: a tracked own echo, a
   * roster-adjacent state change). The editor-driven typing path mutates and
   * paints the doc itself, so its submit + echo leave this untouched; remote
   * applies, optimistic canonical applies (toolbar ops), reloads, heals, and
   * media installs advance it. The react layer repaints on THIS, not on
   * onChange — repainting on onChange queued a whole-document relayout per
   * keystroke, catastrophic past the background-layout page threshold.
   */
  get docVersion(): number {
    return this.docVersionBase + (this.replica?.docVersion ?? 0);
  }

  /** docVersionBase at the last takeRenderScope — a base bump between takes
   * (welcome/heal doc replacement, media install) has no narrow scope. */
  private takenDocVersionBase = 0;

  /**
   * Drain the dirty scope behind the docVersion movement since the last take
   * — parity with the plaintext connection (see connection.ts). Replica
   * applies carry their per-intent scope; base bumps report document scope;
   * null means nothing was recorded (the caller may skip the repaint).
   */
  takeRenderScope(): Scope | null {
    const replicaScope = this.replica?.takeRenderScope() ?? null;
    if (this.docVersionBase !== this.takenDocVersionBase) {
      this.takenDocVersionBase = this.docVersionBase;
      return { kind: "doc" };
    }
    return replicaScope;
  }

  /** Un-confirmed local intents in flight — drained-replay parity with the
   * plaintext connection (doc 15 §2 / rebase.ts). */
  get pendingCount(): number {
    return this.replica?.pendingCount ?? 0;
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

  /** Owner admin op (doc 14 §2.5): honored only if this connection proved
   * the owner token at hello — otherwise the server refuses `not-owner`.
   * The admin channel is PLAINTEXT bookkeeping by design (roles are
   * integrity, not confidentiality — doc 14 §2.5 honest limits), so the
   * encrypted connection sends the same frame the plaintext one does.
   * HISTORY: this method was MISSING here while present on the plaintext
   * connection — the react layer's `as unknown as CollabConnection` cast
   * hid it from the typechecker, so the demo's owner controls were silent
   * TypeErrors on encrypted docs (the default), long misattributed to a
   * Vite prebundle race ("stale session w/o admin"). */
  admin(action:
    | { op: "kick"; clientId: string }
    | { op: "readOnly"; on: boolean }
    | { op: "setRole"; clientId: string; role: "editor" | "viewer" }): void {
    this.transport.send({ t: "admin", action });
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

  /**
   * COLLABORATIVE UNDO (plan doc 03 Phase 8), encrypted flavour.
   *
   * The mirror is a real DocumentSession fed the canonical log, so it already
   * knows the inverse of this client's last action — no new wire message and
   * no server involvement are needed. The inverse is submitted as an ORDINARY
   * sealed intent with `base` set to the ORIGINAL action's seq, which is what
   * makes undo correct under concurrency: the canonical transform rebases it
   * through everything sequenced since, and if the target has been deleted or
   * changed the apply rejects it identically on every replica.
   *
   * Two things this deliberately does NOT do. It does not call the mirror's
   * `undo()`, because that would apply to the mirror out of band and the
   * mirror may only advance by ingesting sequenced envelopes. And it does not
   * paint an undo it can already prove is dead: `targetsResolve` checks the
   * canonical id table first, so the common conflict (someone else removed
   * the thing) is declined up front rather than painted and yanked back.
   */
  undoLast(): UndoOutcome {
    if (!this.mirror || !this.replica || !this.keys || !this.genesisId) return "unavailable";
    const candidate = this.mirror.takeUndo(this.clientId);
    if (candidate.kind !== "undoable") return candidate.kind;
    //
    // TWO WAYS AN UNDO CAN DECLINE, AND THEY GET OPPOSITE TREATMENT. Do not
    // unify them — they look alike (both "undo didn't happen") and are not.
    //
    //   MARKER (`cannot-undo`, decided in takeUndo): the last action has no
    //   inverse yet — say, a bold. Its effect is STILL IN the document, so
    //   reaching past it to undo something older would surprise the user:
    //   their bold stays, their earlier typing vanishes. The entry is NOT
    //   popped; undo stops here. A temporary stop that disappears the day
    //   someone implements that inverse.
    //
    //   OVERTAKEN (`changed-since`, decided just below): the target is
    //   already gone in canonical state — a peer deleted the image first. The
    //   effect this undo would produce is ALREADY TRUE, so consuming the entry
    //   leaves the document exactly where the undo aimed, and the next press
    //   moving to the previous action cannot surprise anyone. The entry IS
    //   popped (takeUndo already did). Refusing would strand undo forever at
    //   an entry that can never become reversible — strictly worse than the
    //   marker's temporary stop.
    //
    if (!this.mirror.targetsResolve(candidate.inverse)) return "changed-since";
    // REBASE HERE, not on the wire. The inverse describes the document as it
    // was at `candidate.seq`; everything sequenced since must be transformed
    // through before it can be applied to the document as it is NOW.
    //
    // Submitting it raw at the old base and letting the canonical transform
    // rebase it — the plaintext session's own trick — does NOT work for an
    // optimistic client: the local apply would paint the stale form (measured:
    // undoing "HELLO" after a peer typed "ZZ" in front of it deleted "abHEL"
    // and left "ZZLO"), and the pending copy would then disagree with the
    // canonical form the server derives, which is the stale-base divergence
    // class this codebase has been bitten by before. Rebasing against the
    // MIRROR — the same canonical log the server would transform against —
    // yields an intent that is correct at the current state, so it submits as
    // an ordinary intent at the ordinary base and every existing mechanism
    // (optimistic apply, echo, rollback-replay) applies unchanged.
    const ahead = this.mirror
      .entriesSince(candidate.seq)
      .filter((e): e is Extract<typeof e, { kind: "applied" }> =>
        e.kind === "applied" && e.intent.clientId !== this.clientId)
      .map((e) => e.intent);
    const rebased = transformIntent(
      {
        ...candidate.inverse,
        clientId: this.clientId,
        // Reserved range: keeps the undo itself off every replica's undo stack.
        clientSeq: this.mirror.nextUndoClientSeq(this.clientId),
        base: candidate.seq,
      } as Intent,
      ahead,
    );
    const full = { ...rebased, base: this.replica.confirmedSeq } as Intent;
    // Optimistic canonical apply, exactly like a toolbar op: the inverse
    // paints immediately and rolls back through the normal machinery if the
    // race it could not foresee (target dies mid-flight) rejects it.
    this.replica.submitLocal(full);
    this.enqueue(async () => {
      try {
        this.lastSent = full;
        const env = await sealIntent(this.keys!.kContent, this.docId, this.genesisId!, full);
        this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
      } catch (err) {
        // ITS OWN CATCH, so the loss accounting can tell submits from ingest.
        // Both ride the same serial chain, and the chain's catch cannot
        // distinguish them — but a seal or send that throws HERE is a LOST
        // EDIT (the replica already applied it optimistically and counted it
        // pending), while an ingest throw is not. Reporting both under one
        // label would make the next B13-style arithmetic ambiguous again,
        // which is the exact failure that investigation turned on.
        this.sendFailures++;
        this.cb.onError?.({ where: "enc.submit", error: err });
      }
    });
    this.scheduleSelfCheck();
    this.cb.onChange?.();
    return "undone";
  }

  /** What the next undo WOULD do, for enabling a button and its tooltip. */
  undoState(): UndoOutcome {
    if (!this.mirror) return "unavailable";
    const state = this.mirror.undoState(this.clientId);
    return state === "undoable" ? "undone" : state;
  }

  /**
   * Throws on the ingest/submit chain — in practice a seal or transport send
   * that failed AFTER the edit was accepted locally and counted as pending.
   *
   * Separate from droppedPreReady on purpose, and the arithmetic is why: that
   * one counts edits refused BEFORE they were ever applied, this one counts
   * edits lost AFTER. Summing them against a swarm run's "lost" figure is how
   * B13's 231 would have been attributed in minutes instead of a night, and
   * merging the two would have made that sum ambiguous again.
   */
  sendFailures = 0;
  /** Submits dropped because the connection was not ready — see
   * ConnectionCallbacks.onSubmitDropped. The window is WIDER here than in the
   * plaintext connection: this one must also derive epoch keys and open the
   * seeded checkpoint before it can honour anything, which on a large
   * document is real time during which the editor's submit path is reachable. */
  droppedPreReady = 0;

  private submitFull(intent: Omit<Intent, "clientId" | "clientSeq" | "base">, preApplied: boolean): void {
    if (!this.replica || !this.keys || !this.genesisId) {
      // LOUD, NOT LOSSY — same contract as the plaintext connection, and the
      // same reasoning against buffering, with one more term: there is no
      // kContent yet, so the intent could not even be sealed.
      this.droppedPreReady++;
      this.cb.onSubmitDropped?.("not-ready");
      return;
    }
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
      try {
        this.lastSent = full;
        const env = await sealIntent(this.keys!.kContent, this.docId, this.genesisId!, full);
        this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
      } catch (err) {
        // ITS OWN CATCH, so the loss accounting can tell submits from ingest.
        // Both ride the same serial chain, and the chain's catch cannot
        // distinguish them — but a seal or send that throws HERE is a LOST
        // EDIT (the replica already applied it optimistically and counted it
        // pending), while an ingest throw is not. Reporting both under one
        // label would make the next B13-style arithmetic ambiguous again,
        // which is the exact failure that investigation turned on.
        this.sendFailures++;
        this.cb.onError?.({ where: "enc.submit", error: err });
      }
    });
    // Arm the stuck-pending watchdog NOW: if the seal/send above throws (the
    // queue's catch swallows task failures by design) or the frame is lost,
    // no broadcast will ever arrive to re-arm the checks — this is the only
    // guaranteed trigger for a lone client's lost op.
    this.scheduleSelfCheck();
    this.cb.onChange?.();
  }

  /**
   * OUTBOUND presence: a COALESCING single-flight sealer, not a queue.
   *
   * Sealing is async, so the send can no longer be a straight-line call, and
   * three shapes were considered. Fire-and-forget lets two rapid caret moves
   * finish encrypting out of order and paint a stale cursor. Reusing
   * `this.queue` would put a high-frequency, low-value stream behind — and
   * ahead of — broadcast ingest, which is the ordering the document's
   * correctness depends on; perf B10 and B13 were both about work piling up
   * on that chain, and this is that lesson applied before the fact rather
   * than after. Chaining every move onto a private promise fixes ordering but
   * still queues N seals for N twitches of the mouse.
   *
   * So: at most ONE seal in flight, and at most ONE position waiting. A move
   * that arrives while a seal is running REPLACES the waiting one — which is
   * not a compromise but the correct semantics, because presence is
   * last-write-wins and a superseded caret has no reader. Bounded by
   * construction: two slots, whatever the input rate.
   */
  /** Serial chain for INBOUND presence opens (see the receive branch). */
  private presenceInQueue: Promise<void> = Promise.resolve();
  private presenceLatest: { position: PresencePosition | null } | null = null;
  private presenceFlushing = false;

  setPresence(position: PresencePosition | null): void {
    // SEALED (doc 13 / #20). Presence used to ride the wire in the clear even
    // in encrypted rooms: the blind sequencer could not read what anyone
    // wrote, but it could watch where everyone pointed, and selection ranges
    // added "how much is selected" on top. Now the position is sealed under
    // K_presence with the sender bound into the AAD, so the relay sees two
    // base64 strings. The structural clamp still runs — on the SENDER here,
    // and again on every receiver after opening, because a hostile
    // participant can seal whatever it likes.
    const clean = sanitizePresencePosition(position);
    const keys = this.keys;
    const genesisId = this.genesisId;
    if (!keys || !genesisId) {
      // Pre-welcome: there is no key yet. Presence is ephemeral and the next
      // caret move re-sends, so dropping is correct and costs nothing — the
      // only thing that would be wrong is leaking it in the clear instead.
      return;
    }
    this.presenceLatest = { position: clean };
    if (this.presenceFlushing) return; // the in-flight loop will pick this up
    this.presenceFlushing = true;
    void (async () => {
      try {
        while (this.presenceLatest) {
          const next = this.presenceLatest;
          this.presenceLatest = null;
          const sealed = await sealPresence(keys.kPresence, this.docId, genesisId, this.clientId, next.position);
          this.transport.send({ t: "presence", position: sealed });
        }
      } catch {
        // A failed seal costs this one caret; the next move re-sends.
      } finally {
        this.presenceFlushing = false;
      }
    })();
  }

  private enqueue(task: () => Promise<void>): void {
    this.queue = this.queue.then(task).catch((err: unknown) => {
      // A failed task must not wedge the chain — but it must not vanish
      // either. Ingest failures ARE handled inside their own tasks (an
      // unopenable envelope becomes a deterministic no-op, not a throw), so
      // anything arriving here is unexpected. It also carries SUBMITS: the
      // seal-and-send half of submitFull rides this same chain, so a throw
      // here is a lost edit with no other trace — the exact shape B13 spent a
      // night chasing.
      this.sendFailures++;
      this.cb.onError?.({ where: "enc.queue", error: err });
    });
  }

  private onServer(msg: ServerMessage): void {
    switch (msg.t) {
      case "welcome-enc": {
        // Read OUTSIDE the serial chain: it is a plain number with no
        // dependency on keys or the mirror, and a consumer asking "how big may
        // this file be" should not wait on a rehydrate to find out.
        this.mediaMaxBlobBytes = msg.mediaMaxBlobBytes ?? null;
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
          this.docVersionBase += (this.replica?.docVersion ?? 0) + 1; // new doc object ⇒ repaint
          this.replica = new ClientReplica(bytes, (cp.sidecar ?? undefined) as never);
          this.replica.confirmedSeq = msg.checkpoint.seq;
          // Doc 16 §6 late-join: the addresses ride inside the SEALED body,
          // so they are available only here, after the checkpoint opened.
          // Tolerantly read — a checkpoint from an older build has none and
          // this replica simply keeps the reserve-but-cannot-fill behavior.
          applyMediaAddresses(this.replica.doc, cp.mediaMeta);
          applyMediaAddresses(this.mirror.doc, cp.mediaMeta);
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
          // Doc 16 §5.3/§5.4: restore media addresses AND pixels from the
          // resumed bundle before volunteering — an E2EE checkpoint usually
          // carries the bytes, but a bundle-resuming holder is the one who
          // can rescue media the relay has already evicted.
          if (resumed?.mediaMeta) {
            const held = DocxDocument.load(resumed.confirmedBytes);
            for (const [part, meta] of resumed.mediaMeta) {
              this.replica.doc.mediaMeta.set(part, meta);
              const pixels = held.media(part);
              if (pixels && this.replica.doc.mediaStatus(part) === "pending") {
                this.replica.doc.installMedia(part, pixels);
              }
            }
          }
          this.media?.volunteer(msg.mediaNeeded);
          void this.media?.fetchPending(this.docId);
        });
        return;
      }
      case "broadcast-enc": {
        this.enqueue(async () => {
          if (!this.replica) return; // pre-welcome broadcasts are impossible; guard anyway
          for (const env of msg.entries) await this.ingest(env);
          if (this.replica.reloaded) this.docEpoch++;
          this.cb.onChange?.();
          this.scheduleSelfCheck();
          // Eager fetch (doc 16 §5.2 step 2) — an applied insertImage leaves
          // a hole this pulls immediately.
          void this.media?.fetchPending(this.docId);
        });
        return;
      }
      case "media-request": {
        this.media?.answerRequest(msg.sha);
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "media-upload": {
        // Chosen re-supplier. In an encrypted room this re-seals with the
        // RECORDED IV so the ciphertext — and therefore the address — is
        // reproduced exactly; see MediaClient.resupply.
        void this.media?.resupply(this.docId, msg.sha);
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "media-ready": {
        void this.media?.onReady(this.docId, msg.sha);
        this.cb.onMediaReady?.(msg.sha);
        return;
      }
      case "media-unavailable": {
        this.media?.onUnavailable(msg.sha);
        this.cb.onMediaUnavailable?.(msg.sha);
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
        const raw = msg.position;
        if (raw === null) {
          this.cb.onPresence?.(msg.participant, null);
          return;
        }
        if (!isSealedPresence(raw)) {
          // A PLAINTEXT position in an encrypted room. The e4 fence should
          // make this unreachable (a peer old enough to send one cannot join),
          // so treat it as noise rather than render it: accepting it would
          // quietly re-open the leak this change closes — the peer that sent
          // it also sent it to the server in the clear.
          return;
        }
        const keys = this.keys;
        const genesisId = this.genesisId;
        if (!keys || !genesisId) return; // pre-welcome presence: nothing to open it with
        // INBOUND opens are serialized on their own chain — separate from the
        // outbound coalescer above, and emphatically not `this.queue`. Serial
        // rather than coalescing because these are OTHER participants: two
        // carets from the same peer must not resolve out of order, and two
        // peers' carets must not be collapsed into one. The chain does not
        // accumulate — each link settles as its decrypt finishes, so its depth
        // is whatever the network delivered, not a backlog this code created.
        this.presenceInQueue = this.presenceInQueue
          .then(async () => {
            // The AAD binds the sender the SERVER stamped, so a relay that
            // re-attributes a blob to another participant fails to decrypt
            // rather than mislabelling a cursor.
            const opened = await openPresence(keys.kPresence, this.docId, genesisId, msg.participant, raw);
            // Post-open clamp: decryption proves WHO sealed it, never that the
            // contents are sane. A hostile participant can seal an unbounded
            // range list under a perfectly valid key.
            this.cb.onPresence?.(msg.participant, sanitizePresencePosition(opened as PresencePosition | null));
          })
          .catch(() => {
            // Undecryptable, tampered, replayed from another room, or garbage:
            // drop THIS participant's caret and nothing else. Never throws
            // into the ingest path — one bad peer must not cost the document.
          });
        return;
      }
      case "roster": {
        this.roster = msg.roster;
        this.cb.onRoster?.(msg.roster);
        return;
      }
      // Session lifecycle warnings are PLAINTEXT bookkeeping and identical in
      // both modes: the deadline is the server's own policy, not document
      // content, so a blind sequencer announces it in the clear. Dispatched
      // directly rather than through the serial queue — nothing here touches
      // keys, the mirror, or the replica.
      case "session-warning": {
        this.cb.onSessionWarning?.({ reason: msg.reason, inMs: msg.inMs });
        return;
      }
      case "session-warning-cleared": {
        this.cb.onSessionWarningCleared?.({ reason: msg.reason });
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
          // re-seal the intent (same clientId/clientSeq — dedup-safe) with a
          // fresh base and resubmit. One-in-flight discipline makes lastSent
          // the only possible referent — but the BODY must be the replica's
          // PENDING copy, not lastSent verbatim: replayPending has already
          // transformed pending against every remote entry ingested since
          // (the same transform the server would have applied had the base
          // been fresh). Resealing the raw original at the bumped base makes
          // the server apply UNTRANSFORMED coordinates verbatim while the
          // optimistic doc holds the transformed form; the own-echo fast
          // path then snapshots the divergent optimistic doc as confirmed —
          // silent permanent divergence at the checkpoint boundary. If the
          // pending copy was dropped during replay, do NOT resubmit: the
          // edit no longer exists locally, and optimistic == canonical only
          // if it stays out of the log on both sides.
          const retry = this.lastSent;
          if (retry && this.replica && this.keys && this.genesisId) {
            this.enqueue(async () => {
              const body = this.replica!.pendingIntent(retry.clientId, retry.clientSeq);
              if (!body) return;
              const rebased = { ...body, base: this.replica!.confirmedSeq } as Intent;
              const env = await sealIntent(this.keys!.kContent, this.docId, this.genesisId!, rebased);
              this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
            });
          }
          return;
        }
        if (msg.reason === "rate-limit") {
          // Throttled, not failed: reset the watchdog retry budget
          // (throttling must never escalate to drop+heal) and RE-DRIVE THE
          // WHOLE PENDING QUEUE after backoff — a refused op in a CHAINED
          // burst (each op addressing ids a previous op allocated) strands
          // everything behind it, and front-only recovery redelivers one op
          // per watchdog window, far too slow for a bulk client. Dedup by
          // (clientId, clientSeq) makes full-queue resends free.
          this.stuckRetries = 0;
          this.scheduleRateLimitRedrive();
          return;
        }
        this.cb.onRefused?.(msg.reason);
        return;
      }
    }
  }

  /** Re-drive the whole pending queue after rate-limit backoff. One timer
   * at a time; backoff doubles per refusal and resets when the queue
   * drains (see scheduleSelfCheck). */
  private scheduleRateLimitRedrive(): void {
    if (this.redriveTimer !== null) return;
    this.redriveBackoffMs = Math.min(this.redriveBackoffMs === 0 ? 300 : this.redriveBackoffMs * 2, 5000);
    this.redriveTimer = setTimeout(() => {
      this.redriveTimer = null;
      this.enqueue(async () => {
        if (!this.replica || !this.keys || !this.genesisId) return;
        for (const body of this.replica.pendingCopies()) {
          const rebased = { ...body, base: this.replica.confirmedSeq } as Intent;
          const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, rebased);
          this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
        }
      });
    }, this.redriveBackoffMs);
  }

  /** Debounce a quiescent self-check: only when nothing is pending (a burst
   * still in flight legitimately diverges live from canonical), re-armed by
   * every broadcast so the check runs once the wire goes quiet. While
   * pending exists, arm the stuck-pending watchdog instead. */
  private scheduleSelfCheck(): void {
    if (!this.replica) return;
    if (this.replica.pendingCount !== 0) {
      this.scheduleStuckCheck();
      return;
    }
    this.redriveBackoffMs = 0; // queue drained — rate-limit backoff resets
    if (this.redriveTimer !== null) {
      clearTimeout(this.redriveTimer);
      this.redriveTimer = null;
    }
    if (this.stuckTimer !== null) {
      clearTimeout(this.stuckTimer);
      this.stuckTimer = null;
    }
    this.stuckKey = null;
    this.stuckRetries = 0;
    if (this.selfCheckTimer !== null) clearTimeout(this.selfCheckTimer);
    this.selfCheckTimer = setTimeout(() => {
      this.selfCheckTimer = null;
      this.enqueue(() => this.selfCheck());
    }, this.selfCheckDelayMs);
  }

  /** Watch the FRONT pending op; progress (a different front, or empty)
   * resets the retry budget. The window is 5× the self-check debounce. */
  private scheduleStuckCheck(): void {
    const front = this.replica?.firstPending();
    if (!front) return;
    const key = `${front.clientId}:${front.clientSeq}`;
    if (this.stuckKey !== key) {
      this.stuckKey = key;
      this.stuckRetries = 0;
    }
    if (this.stuckTimer !== null) clearTimeout(this.stuckTimer);
    this.stuckTimer = setTimeout(() => {
      this.stuckTimer = null;
      this.enqueue(() => this.stuckCheck(front.clientId, front.clientSeq));
    }, this.selfCheckDelayMs * 5);
  }

  /** Runs inside the serial queue: if the same front pending is still
   * stuck, resend its transformed copy at the current base; past the retry
   * budget, drop it and heal from the mirror (convergence over delivery). */
  private async stuckCheck(clientId: string, clientSeq: number): Promise<void> {
    if (!this.replica || !this.mirror || !this.keys || !this.genesisId) return;
    const front = this.replica.firstPending();
    if (!front || front.clientId !== clientId || front.clientSeq !== clientSeq) {
      // Progressed — if fully drained, the normal quiescent path re-arms.
      this.scheduleSelfCheck();
      return;
    }
    if (this.stuckRetries >= EncryptedCollabConnection.STUCK_MAX_RETRIES) {
      const liveHash = await docHash(this.replica.doc);
      const canonicalHash = await docHash(this.mirror.doc);
      const cp = this.mirror.checkpoint();
      this.docVersionBase += this.replica.docVersion + 1; // doc object replaced ⇒ repaint
      this.replica = new ClientReplica(cp.docx, cp.sidecar as never);
      this.replica.confirmedSeq = cp.seq;
      this.docEpoch++;
      this.selfHeals++;
      this.stuckKey = null;
      this.stuckRetries = 0;
      this.cb.onSelfHeal?.({ seq: cp.seq, liveHash, canonicalHash });
      this.cb.onChange?.();
      return;
    }
    this.stuckRetries++;
    const body = this.replica.pendingIntent(clientId, clientSeq);
    if (body) {
      const rebased = { ...body, base: this.replica.confirmedSeq } as Intent;
      const env = await sealIntent(this.keys.kContent, this.docId, this.genesisId, rebased);
      this.transport.send({ t: "submit-enc", envelope: env } as ClientMessage);
    }
    this.scheduleStuckCheck();
  }

  /** Runs INSIDE the serial queue (no ingest can interleave): compare the
   * optimistic replica against the mirror at quiescence; on drift, rebuild
   * the replica from the mirror — the same shape as a reload heal, but
   * automatic, local, and scoped to the one client that drifted. */
  private async selfCheck(): Promise<void> {
    if (!this.replica || !this.mirror) return;
    if (this.replica.pendingCount !== 0) return; // burst resumed — next drain re-arms
    const seq = this.mirror.seq;
    if (seq === this.lastSelfCheckSeq) return; // this state already verified
    const liveHash = await docHash(this.replica.doc);
    const canonicalHash = await docHash(this.mirror.doc);
    if (this.replica.pendingCount !== 0 || this.mirror.seq !== seq) return; // raced a local submit
    this.lastSelfCheckSeq = seq;
    if (liveHash === canonicalHash) return;
    // DRIFT (the B6a class): the optimistic machinery lost/garbled an edit
    // while the canonical derivation is intact. Rebuild the replica from the
    // mirror — byte-exact canonical state + the exact id table (checkpoint()
    // is a pure snapshot; save() is side-effect-free by invariant).
    const cp = this.mirror.checkpoint();
    this.docVersionBase += this.replica.docVersion + 1; // doc object replaced ⇒ repaint
    this.replica = new ClientReplica(cp.docx, cp.sidecar as never);
    this.replica.confirmedSeq = cp.seq;
    this.docEpoch++;
    this.selfHeals++;
    this.cb.onSelfHeal?.({ seq, liveHash, canonicalHash });
    this.cb.onChange?.();
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
    // A rejection is where an edit DIES, and in encrypted rooms this mirror is
    // the only place it happens — the server is blind and cannot validate, so
    // there is no server-side signal at all. Surfacing it here is the whole
    // reason the channel exists: the stale image-size bound rejected every
    // large photo identically on every replica, which means nothing diverged
    // and nothing complained. See ConnectionCallbacks.onIntentRejected.
    if (entry.kind === "rejected" && entry.clientId === this.clientId) {
      this.cb.onIntentRejected?.({ reason: entry.reason, clientSeq: entry.clientSeq });
    }
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
        // Doc 16 §6 late-join, sealed half: the MIRROR never installs pixels,
        // but it holds every REGISTRATION — which is all a joiner needs to go
        // and fetch them. Inside the sealed body, so a blind server still
        // learns nothing about the document's part structure.
        mediaMeta: mediaAddressesOf(this.mirror!.doc),
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
