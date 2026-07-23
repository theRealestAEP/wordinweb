import { ClientReplica } from "./replica.js";
import { DocxDocument } from "@wordinweb/core";
import { Intent } from "./intents.js";
import {
  ClientMessage,
  ServerMessage,
  PROTOCOL_VERSION,
  PresencePosition,
  ParticipantProfile,
  RosterEntry,
} from "./protocol.js";
import type { DocBundle } from "./bundle.js";

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
  /**
   * Resume landed in a DIFFERENT epoch than the bundle's (doc 12 §5 case 2):
   * someone re-seeded while this client was away. The connection has taken
   * the server's state and did NOT replay the bundle's pending intents
   * (they belong to the old epoch — replaying them would be a silent
   * cross-epoch merge, which the fork rule forbids). The consumer decides
   * what to do with the old bundle: doc-15 lineage fast-forward/draft, or
   * pre-lineage, "your offline copy is saved as a draft".
   */
  onEpochChange?: (storedGenesisId: string, currentGenesisId: string) => void;
  /** The session roster changed (join/leave/rename) — full snapshot. */
  onRoster?: (roster: RosterEntry[]) => void;
  /**
   * Resume landed in a different epoch that is a STRICT DESCENDANT of this
   * client's stored head (doc 15 §1 fast-forward): the seed's lineage
   * contains our (genesisId, hash) at ≥ our seq, so adopting the server's
   * state loses nothing — no draft, no banner noise. The consumer should
   * still bank the superseded bundle recoverable-not-gone (version ring /
   * a superseded slot) per doc 15's fabricated-lineage mitigation.
   */
  onFastForward?: (fromGenesisId: string, toGenesisId: string) => void;
  /** Media re-supply control (doc 16 §5): the room asks who holds `sha`;
   * a holder answers via `mediaHave` and uploads over HTTP when chosen. */
  onMediaRequest?: (sha: string) => void;
  /** A needed blob became fetchable / definitively unavailable. */
  onMediaReady?: (sha: string) => void;
  onMediaUnavailable?: (sha: string) => void;
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

  /** The epoch id of the session this connection joined (from the welcome).
   * The resume layer compares it against the bundle's stored genesisId:
   * same ⇒ seamless rejoin; different ⇒ someone re-seeded while away
   * (doc 12 §5 case 2 — take server state, keep the local copy; the doc-15
   * lineage check decides fast-forward vs draft). Null until welcomed. */
  genesisId: string | null = null;
  /** Session encryption mode from the welcome (doc 13 §6). The E2EE layer
   * hard-refuses a value contradicting the link's `#k` fragment. */
  mode: "plaintext" | "encrypted" | null = null;

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

  /**
   * Join a document. The server replies with a welcome (snapshot + sidecar +
   * tail). The hello carries this connection's clientId — the hub binds it to
   * the socket and refuses submits under any other id (doc 11 decision 8).
   * `takeover: true` claims the identity from an existing live connection
   * (the doc-12 §7 "use here instead" path for a second same-profile tab);
   * without it, a duplicate join is refused `already-open`.
   */
  join(docId: string, token?: string, opts?: { takeover?: boolean; profile?: ParticipantProfile; codeProof?: string }): void {
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
    });
  }

  /** The latest roster snapshot (doc 14 §2); empty until the first fan-out. */
  roster: RosterEntry[] = [];

  /**
   * Attribution layer 1 (doc 14 §3): the canonical log IS the attribution
   * record — every applied entry carries its author's bound clientId. This
   * is the bounded client-side derivation an activity panel renders
   * (clientId → roster name/color). Newest last, capped, derived state:
   * never persisted, dies with the connection (zero custody).
   */
  activity: { seq: number; clientId: string; kind: string }[] = [];
  private static ACTIVITY_CAP = 100;

  private recordActivity(entries: { seq: number; kind: string; intent?: Intent }[]): void {
    for (const e of entries) {
      if (e.kind === "applied" && e.intent) {
        this.activity.push({ seq: e.seq, clientId: e.intent.clientId, kind: e.intent.kind });
      }
    }
    if (this.activity.length > CollabConnection.ACTIVITY_CAP) {
      this.activity.splice(0, this.activity.length - CollabConnection.ACTIVITY_CAP);
    }
  }

  /** Ask the room for a blob this replica needs (doc 16 §5.2). */
  mediaNeed(sha: string): void {
    this.transport.send({ t: "media-need", sha });
  }
  /** Volunteer holdings (reply to a request, or after a welcome whose
   * mediaNeeded intersects local media — §5.4). */
  mediaHave(shas: string[]): void {
    if (shas.length) this.transport.send({ t: "media-have", shas });
  }
  /** Shas this replica can re-supply: metadata of parts whose bytes are
   * PRESENT (pending parts are exactly what we can't serve). */
  heldMediaShas(): string[] {
    const doc = this.replica?.doc;
    if (!doc) return [];
    const out: string[] = [];
    for (const [part, meta] of doc.mediaMeta) {
      if (doc.mediaStatus(part) === "ready") out.push(meta.sha);
    }
    return out;
  }

  /** Rename/recolor this participant mid-session. The server sanitizes and
   * fans out the updated roster; the local copy updates on that echo (no
   * optimistic roster — a 1-RTT lag on your own rename is imperceptible and
   * keeps one code path). */
  setProfile(profile: ParticipantProfile): void {
    this.transport.send({ t: "profile", profile });
  }

  /** Bundle being resumed from (set by resume(), consumed at welcome). */
  private resuming: DocBundle | null = null;

  /**
   * Rejoin a document from a persisted bundle (doc 12 §5). Restores the
   * clientSeq watermark FIRST — a fresh counter would reuse already-
   * sequenced (clientId, clientSeq) keys and the server would dedup this
   * client's NEW edits as re-sends (silent edit loss). The welcome decides
   * the case: same epoch ⇒ replay pending (below); different ⇒
   * onEpochChange, pending withheld.
   */
  resume(bundle: DocBundle, token?: string, opts?: { profile?: ParticipantProfile; codeProof?: string }): void {
    this.clientSeq = Math.max(this.clientSeq, bundle.clientSeq);
    this.resuming = bundle;
    this.transport.send({
      profile: opts?.profile,
      codeProof: opts?.codeProof,
      t: "hello",
      protocolVersion: PROTOCOL_VERSION,
      docId: bundle.docId,
      clientId: this.clientId,
      // Resume IS the zombie-takeover case (doc 12 §7): after a crash or
      // refresh, the previous socket for this identity may still be inside
      // the 60s grace window server-side. This connection is that identity's
      // continuation — claiming it is correct, and the incumbent (if any)
      // is a dead tab by definition of "we are resuming from the bundle".
      takeover: true,
      token,
      sinceSeq: bundle.confirmedSeq,
      genesisId: bundle.genesisId,
    });
  }

  /**
   * The current durable state as a doc-12 bundle, or null before welcome.
   * `savedAt` is stamped by the persister at write time (clock injection —
   * this module never reads Date.now itself).
   */
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
        // The sidecar rides with the snapshot (round-4 F10) so the replica's
        // id table matches the server's even across split-created ids.
        this.replica = new ClientReplica(base64ToBytes(msg.snapshot), msg.sidecar);
        this.replica.confirmedSeq = msg.seq;
        this.genesisId = msg.genesisId;
        this.mode = msg.mode;
        if (msg.tail.length) this.replica.receive(msg.tail);
        // Doc 16 §5.3: restore media ADDRESSES from the resumed bundle —
        // the welcome snapshot carries pixels but not shas (metadata is
        // in-memory on the doc); without this a resumed holder could not
        // volunteer or verify. Imported before the §5.4 intersection below.
        if (this.resuming?.mediaMeta) {
          // The welcome snapshot carries the REGISTRATION; the relay may
          // long since have evicted the bytes (that's the re-supply
          // premise) — but this holder's bundle has them. Restore pixels
          // from the bundle's own zip so this client renders immediately
          // AND can serve re-supply for everyone else.
          const held = DocxDocument.load(this.resuming.confirmedBytes);
          for (const [part, meta] of this.resuming.mediaMeta) {
            this.replica.doc.mediaMeta.set(part, meta);
            const bytes = held.media(part);
            if (bytes && this.replica.doc.mediaStatus(part) === "pending") {
              this.replica.doc.installMedia(part, bytes);
            }
          }
        }
        // §5.4: a joining holder resurrects "unavailable" media — intersect
        // the room's needs with local holdings and volunteer.
        if (msg.mediaNeeded?.length) {
          const held = new Set(this.heldMediaShas());
          this.mediaHave(msg.mediaNeeded.filter((sha) => held.has(sha)));
        }
        this.recordActivity(msg.tail as never);

        // Resume epilogue (doc 12 §5). Case 1 — same epoch: replay the
        // bundle's pending intents by RE-SENDING them, untracked. This is
        // deliberately fire-and-observe: an intent the server already
        // sequenced pre-crash dedups to its old entry (seq ≤ welcome.seq,
        // filtered as stale — it's already inside the snapshot); one it
        // never saw sequences fresh and arrives as an ordinary remote
        // broadcast this replica applies in place. Either way the edit ends
        // up in the doc exactly once, with zero client-side bookkeeping —
        // the (clientId, clientSeq) dedup is the whole mechanism.
        // Case 2 — different epoch: someone re-seeded while we were away;
        // pending belongs to the OLD epoch and is withheld (replaying it
        // would silently merge across epochs — the fork rule forbids it);
        // the consumer is told and decides (draft / doc-15 lineage).
        const resumed = this.resuming;
        this.resuming = null;
        if (resumed) {
          if (resumed.genesisId === msg.genesisId) {
            for (const intent of resumed.pending) this.transport.send({ t: "submit", intent });
          } else if (isAncestorOf(resumed, msg.lineage)) {
            // Doc 15 fast-forward: our head is in the seed's chain at >= our
            // seq with a MATCHING hash — the new epoch strictly extends what
            // we have; adopt silently. (Hash equality is the ancestry proof;
            // a fabricated chain without our real hash fails this check and
            // falls through to the fork path below.)
            this.cb.onFastForward?.(resumed.genesisId, msg.genesisId);
          } else {
            this.cb.onEpochChange?.(resumed.genesisId, msg.genesisId);
          }
        }
        this.cb.onChange?.();
        return;
      }
      case "broadcast": {
        this.replica?.receive(msg.entries);
        this.recordActivity(msg.entries as never);
        if (this.replica?.reloaded) this.docEpoch++;
        this.cb.onChange?.();
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
      case "media-request": {
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "media-ready": {
        this.cb.onMediaReady?.(msg.sha);
        return;
      }
      case "media-unavailable": {
        this.cb.onMediaUnavailable?.(msg.sha);
        return;
      }
      case "media-upload": {
        // Chosen as the re-supplier: surface as a request for THIS sha —
        // the app's holder duty uploads over HTTP (bytes never ride the WS).
        this.cb.onMediaRequest?.(msg.sha);
        return;
      }
      case "refused": {
        this.cb.onRefused?.(msg.reason);
        return;
      }
    }
  }
}

/** Doc 15 §1 ancestry check: the rejoiner's own recorded head must appear
 * in the seed's lineage — same epoch, at least our seq, and the HASH of the
 * state we hold. O(chain), no dates involved (clocks are untrusted). */
function isAncestorOf(
  bundle: DocBundle,
  seedLineage: { genesisId: string; seq: number; docHash: string }[] | undefined,
): boolean {
  if (!seedLineage) return false;
  const ownHead = bundle.lineage?.[bundle.lineage.length - 1];
  if (!ownHead) return false;
  return seedLineage.some(
    (h) => h.genesisId === ownHead.genesisId && h.seq >= ownHead.seq && h.docHash === ownHead.docHash,
  );
}
