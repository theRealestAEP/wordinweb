import type { Intent, LogEntry } from "./intents.js";
import type { IdSidecar } from "./session.js";
import type { IntentEnvelope, SealedPresence } from "./e2ee.js";

/** A sequenced envelope in an encrypted room's log (doc 13 §2): the
 * server-assigned seq plus the opaque envelope. The server can read only
 * the bookkeeping; clients derive the canonical form themselves. */
export type EnvelopeEntry = IntentEnvelope & { seq: number };

/**
 * One head in a document's LINEAGE chain (doc 15 §1): the identity of the
 * confirmed state a session epoch ended at. Ancestry is decided by hash
 * membership — never by dates (clocks are self-asserted). A rejoiner whose
 * own head appears in the seed's lineage is a strict ancestor and can
 * FAST-FORWARD silently; anything else is true divergence (draft/fork).
 */
export interface LineageHead {
  genesisId: string;
  seq: number;
  /** sha256 hex of the confirmed docx bytes at (genesisId, seq). */
  docHash: string;
}

/** A sealed checkpoint blob as the wire carries it (doc 13 §3). */
export interface SealedCheckpoint {
  seq: number;
  iv: string;
  ciphertext: string;
}

/**
 * The engine version fence for E2EE rooms (doc 13 §2): canonical forms are
 * CLIENT-derived in encrypted mode, so a room with mixed transform/apply
 * semantics can diverge with no arbiter. Bump whenever those semantics
 * change; the sequencer refuses a hello whose engineVersion mismatches the
 * room's (the stale client gets the download/draft path). Plaintext rooms
 * ignore it (the server's own apply is the arbiter there, and
 * PROTOCOL_VERSION already gates wire compatibility).
 *
 * THE FENCE ONLY WORKS IF THIS ACTUALLY GETS BUMPED. A live fork was traced
 * to two same-"e1" builds with different apply semantics (a stale tab across
 * the offset-basis migration): 63 ops sequenced, zero refusals, silent
 * canonical divergence — the exact failure this constant exists to prevent.
 * Bump ledger (append a line for every semantics change):
 *   e1  initial E2EE ship
 *   e2  separator-inclusive wire offset basis (B1 rev 2); new intents
 *       resizeDrawing/resizeTableColumn/resizeTableRow/moveTable/
 *       removeDrawing; extender/selection-delete emission changes;
 *       canonical insertBreak
 *   e3  header/footer editing: the apply's run index now covers the
 *       header/footer parts (hf paragraphs were addressable by id but
 *       unresolvable, so every hf text intent silently rejected); new
 *       intents ensureHeaderFooter and moveMath
 *   e4  sealed presence in encrypted rooms (#20). A DELIBERATE OVER-
 *       APPLICATION of the rule above: presence is NOT canonical, so a mixed
 *       room would degrade to missing remote carets rather than forking the
 *       document, and the strict rule would not demand a bump. Bumped anyway
 *       to keep dev rooms homogeneous while this lands, after a live fork was
 *       traced to two same-version builds with different semantics. Do NOT
 *       read this line as evidence that presence is canonical — it is not.
 *   e5  text-box story editing: the apply's run index now descends into
 *       w:txbxContent (the e3 gap one level deeper — story runs were
 *       addressable by id but unresolvable, so every text intent aimed at a
 *       text box silently rejected). Squarely canonical: an e4 peer rejects
 *       the very intent an e5 peer applies, so a mixed room forks the
 *       document on the first keystroke in a pleading number column.
 *   e6  padded sealed bodies (traffic-analysis hardening, e2ee.ts): intent,
 *       presence and checkpoint plaintexts now carry a 4-byte length prefix
 *       and zero fill to fixed buckets, so ciphertext size no longer tracks
 *       edit size. Squarely canonical for intents: an e5 openIntent
 *       JSON.parses the RAW plaintext, so a padded body throws there — the
 *       edit no-ops on every e5 mirror while applying on every e6 mirror, a
 *       document fork on the FIRST padded envelope. (Presence and
 *       checkpoints ride the same bump; sealed media is deliberately NOT
 *       padded — its sha addresses require byte-identical re-encryption.)
 *   e7  table text wrapping: tableOp now carries None/Around. An e6 peer
 *       interprets the new object as vertical alignment and diverges, so the
 *       room must contain only clients that know the new canonical mutation.
 *   e8  collaborative suggestion review adds rejectAllRevisions. An e7 peer
 *       cannot apply that sealed intent, so mixed encrypted clients diverge.
 *   e9  splitParagraph can carry suggestion metadata. Older peers apply the
 *       split without its tracked paragraph mark and diverge immediately.
 *   e10 legacy VML drawing edits: existing position/rotation intents now
 *       resolve v:shape/v:rect carriers as well as w:drawing. An e9 peer
 *       rejects the same intent that an e10 peer applies, so a mixed
 *       encrypted room would fork on the first pleading-gutter object edit.
 *   e11 WordArt insertion writes centered Word-compatible text-body
 *       geometry. An e10 peer generates different canonical DOCX bytes for
 *       the same insertWordArt intent, so mixed encrypted clients diverge.
 *   e12 WordArt style edits add schema-enforced glyph color and opacity. An
 *       e11 peer rejects this canonical intent while an e12 peer applies it.
 *   e13 drawing line edits accept a null color to clear an outline. An e12
 *       peer rejects the canonical clear while an e13 peer applies it.
 *   e14 setListType can carry moreBlockIds: a multi-paragraph toggle is ONE
 *       intent whose single apply mints ONE shared numbering definition
 *       (Word's continuity semantics). An e13 peer ignores the field and
 *       formats only the addressed paragraph, so mixed clients diverge on
 *       the first multi-paragraph list toggle.
 *   e15 keyboard-contract + editor-residual waves: new intents
 *       insertSeparator (Shift+Enter / body Tab; the keyboard-contracts
 *       wave shipped it without this bump — covered here) and
 *       deleteSeparator (separator-aware Backspace/Delete), plus new
 *       editor emissions an e14 peer never produces or applies: cross-cell
 *       deletes clear whole cells, and a fully selected table rides
 *       tableOp("deleteTable") from the selection-delete path. An e14 peer
 *       rejects the very intents an e15 peer applies — fork on the first
 *       soft-break delete.
 *   e16 cell merge/split on the wire: tableOp now carries mergeRight/
 *       mergeDown/splitCell (they existed only as local mutations, a silent
 *       local-only divergence in shared documents), plus new registered
 *       operations sortTableRows, convertTextToTable and convertTableToText.
 *       An e15 peer rejects the very intents an e16 peer applies — fork on
 *       the first merged cell, row sort, or conversion.
 *   e17 paragraph/references depth wave: new registered operations
 *       setTabStops (direct w:tabs editing) and setParagraphBorders
 *       (per-edge w:pBdr + w:shd fill). An e16 peer rejects the very
 *       intents an e17 peer applies — fork on the first tab-stop or
 *       paragraph-border edit.
 */
export const ENGINE_VERSION = "e17";

/**
 * Wire protocol between a collab client and the server host. Transport-
 * agnostic: the hub speaks these messages; a WebSocket (or in-memory test)
 * adapter serializes them. Kept deliberately small (plan doc 06).
 */

/** One highlighted stretch of a remote participant's selection: a half-open
 * `[start, end)` inside ONE run, in the same WIRE offset basis carets and
 * suggestRevision use (cumulative within the run, inline separators counting
 * one unit each — see core `EncodedCaret`). The sender emits one entry per
 * selection segment, and a segment never straddles a w:t, so a range stays
 * inside a single run by construction. */
export interface PresenceRange {
  blockId: number;
  runId: number;
  start: number;
  end: number;
}

/** Cap on `PresencePosition.ranges`. Presence crosses a trust boundary (any
 * participant can hand-craft it, and in E2EE mode the server never validates
 * it), so receivers clamp instead of trusting the sender's array length. A
 * selection with more than this many segments renders its first 64. */
export const PRESENCE_MAX_RANGES = 64;

/** An ephemeral cursor/selection position on the presence channel: stable-id
 * addresses (plan doc 03). Never logged, never persisted. */
export interface PresencePosition {
  anchor: { blockId: number; runId: number; offset: number };
  focus?: { blockId: number; runId: number; offset: number };
  /** Selection highlight (Google-Docs style): the stretches the participant
   * has selected, drawn in their presence color under the remote caret.
   * OPTIONAL and additive — a payload from an older client carries only the
   * anchor and still renders a caret. Absent/empty ⇒ no selection. */
  ranges?: PresenceRange[];
}

/** Ignore-what-you-can't-trust filter for an inbound presence payload: drops
 * non-finite/negative/inverted ranges and clamps the list to
 * PRESENCE_MAX_RANGES, leaving the caret (the pre-ranges shape) untouched.
 * Applied client-side on receive because the hub RELAYS presence verbatim —
 * it has never validated caret coordinates either, and in encrypted rooms it
 * couldn't be the arbiter anyway. Returns the payload unchanged when there is
 * nothing to strip, so the common path allocates nothing. */
export function sanitizePresencePosition(pos: PresencePosition | null): PresencePosition | null {
  if (!pos || !pos.ranges) return pos;
  const int = (n: unknown): boolean => typeof n === "number" && Number.isFinite(n) && n >= 0;
  const clean: PresenceRange[] = [];
  for (const r of Array.isArray(pos.ranges) ? pos.ranges : []) {
    if (clean.length >= PRESENCE_MAX_RANGES) break;
    if (!r || typeof r !== "object") continue;
    if (!int(r.blockId) || !int(r.runId) || !int(r.start) || !int(r.end)) continue;
    if (r.end <= r.start) continue;
    clean.push({ blockId: r.blockId, runId: r.runId, start: r.start, end: r.end });
  }
  if (clean.length === pos.ranges.length) return pos;
  return clean.length > 0 ? { ...pos, ranges: clean } : { ...pos, ranges: undefined };
}

/**
 * A participant's self-asserted display profile (plan doc 14 §2). No
 * accounts: anyone with the link can claim any name — stated in the UI
 * ("names are chosen by participants") — but the F4 clientId binding
 * guarantees CONTINUITY: one name maps to one actor per session and nobody
 * can impersonate an existing participant mid-session. The server sanitizes
 * (length, control chars, color shape) and clients still render text-node-
 * only with palette-validated colors (doc 11 XSS vector 7 — defense in
 * depth, since a hostile server or E2EE mode moves sanitization client-side).
 */
export interface ParticipantProfile {
  /** 1–40 chars after server sanitization. */
  name: string;
  /** Hex color `#rrggbb`; invalid values are replaced server-side by a
   * palette color derived from the clientId hash. */
  color: string;
}

/** One roster row (doc 14 §2): keyed by the BOUND clientId — the same
 * identity intents carry, so attribution, presence, and roster share one
 * keyspace. Disconnected participants stay listed (greyed) for the session's
 * lifetime; a reconnect under the same clientId resumes the entry. */
/**
 * Whether a participant's edits will be SEQUENCED, and if not, why.
 *
 * A POSITIVE signal, which is the entire point. Until this existed a client
 * could only discover it was blocked by making an edit and being refused —
 * which cost the user their first keystrokes every time (optimistic local
 * mutation, refusal, heal: a burst of text that appears and vanishes), and
 * left them stuck in viewer mode after a lock was LIFTED, because nothing
 * announced the good news either.
 *
 * The three blocked states arrive as one `read-only` refusal on the wire, so
 * they are separated here — the difference is what a UI can honestly say:
 *  - `owner-lock`  the owner paused editing for the whole room; it may lift
 *  - `demoted`     the owner made THIS participant a viewer
 *  - `viewer-role` the access token grants read access only; nothing the
 *                  owner does in-session changes it
 */
export type WriteStatus = "allowed" | "owner-lock" | "demoted" | "viewer-role";

/** One roster row (doc 14 §2): keyed by the BOUND clientId — the same
 * identity intents carry, so attribution, presence, and roster share one
 * keyspace. Disconnected participants stay listed (greyed) for the session's
 * lifetime; a reconnect under the same clientId resumes the entry. */
export interface RosterEntry {
  clientId: string;
  profile: ParticipantProfile;
  connected: boolean;
  /** Set by the server after a valid one-time AI invitation is redeemed. */
  participantType?: "human" | "agent";
  /**
   * Whether this participant may write, re-fanned on every transition.
   *
   * OPTIONAL because it is additive: an older server omits it. A client that
   * sees `undefined` must NOT assume "allowed" — it should fall back to the
   * refusal-driven behaviour it had before, exactly as the published media
   * limit's `null` means "skip the pre-check". Inventing a permissive default
   * would put the user back in the first-edit-lost loop this removes.
   *
   * PER-CLIENT, not a room flag: an owner keeps writing through their own
   * room-wide lock, so a single broadcast value would be wrong for them.
   */
  write?: WriteStatus;
}

/** Client → server. */
export type ClientMessage =
  | {
      t: "hello";
      protocolVersion: number;
      docId: string;
      /**
       * The identity every subsequent submit/presence message is bound to.
       * The hub registers it for this socket at hello and REFUSES any submit
       * whose intent.clientId differs (plan doc 11 decision 8 / round-4 F4):
       * `(clientId, clientSeq)` is the idempotency key, so an unbound claimed
       * id would let one client poison another's dedup space (silent edit
       * loss) and forge attribution. Registration-on-first-hello is
       * sufficient — the id needs continuity, not global meaning.
       */
      clientId: string;
      /**
       * Single-live-connection rule (plan doc 12 §7): a second hello for the
       * same (docId, clientId) is refused `already-open` — same-profile tabs
       * share localStorage identity and would collide clientSeq counters.
       * With `takeover: true` the NEW connection wins instead: the hub kicks
       * the old socket (it may be a zombie tab) and admits this one.
       */
      takeover?: boolean;
      token?: string;
      sinceSeq: number;
      /** The epoch id from the client's stored bundle, when resuming (plan
       * doc 12 §5). Informational to the server today — epoch comparison is
       * client-side (welcome.genesisId vs the bundle's) — but carried so the
       * server can log resume-vs-fresh joins and future-proof case handling. */
      genesisId?: string;
      /** Display profile (doc 14 §2) — optional; omitted = anonymous default
       * derived server-side. In E2EE mode this becomes an opaque encrypted
       * blob (doc 13/14) — the shape change rides that protocol bump. */
      profile?: ParticipantProfile;
      /** Engine (transform/apply semantics) version — REQUIRED for encrypted
       * rooms, where the fence prevents mixed-semantics divergence. */
      engineVersion?: string;
      /**
       * Share-code proof (doc 13 §7): the PBKDF2-stretched code, base64.
       * The server compares against the verifier registered at seed time and
       * enforces an attempt budget — 10^6 combinations is plenty against 5
       * online tries, and in E2EE mode the code ALSO mixes into key
       * derivation, so even a server bypass yields undecryptable content.
       */
      codeProof?: string;
      /**
       * Owner-capability proof (doc 14 §2.5): the token minted at seed and
       * returned ONLY to the seeder — never part of the shared link. A
       * matching token flags this connection as the epoch's owner
       * (read-only bypass + admin channel). Roles are capabilities, not
       * accounts: lose the bundle, lose the crown; re-seed to reclaim.
       */
      ownerToken?: string;
      /** One-time room invitation created by an already joined participant. */
      agentInvite?: { inviteId: string; token: string };
    }
  /** Update this connection's profile mid-session (rename / recolor). */
  | { t: "profile"; profile: ParticipantProfile }
  /** Register or revoke an AI invitation. The server binds it to this
   * connection's room identity and never broadcasts the token. */
  | { t: "agent-invite"; inviteId: string; token: string; expiresAt: number }
  | { t: "agent-invite-revoke"; inviteId: string }
  /** End-to-end encrypted private chat. The server only routes the envelope
   * between the invited agent and the participant who created its invite. */
  | { t: "agent-chat"; agentClientId: string; messageId: string; iv: string; ciphertext: string }
  /**
   * Owner admin channel (doc 14 §2.5) — only honored from a connection
   * whose hello proved the owner token. All enforceable by a BLIND server
   * (identity/role are plaintext bookkeeping). Honest limits: kick/demote
   * key by clientId (a determined user can mint a fresh id — a nudge, not
   * a wall, without accounts), and in E2EE mode an owner controls WRITES,
   * never reads (re-key is the only read fence).
   */
  | { t: "admin"; action:
        | { op: "kick"; clientId: string }
        | { op: "readOnly"; on: boolean }
        | { op: "setRole"; clientId: string; role: "editor" | "viewer" } }
  | { t: "submit"; intent: Intent }
  /** Encrypted-mode submit (doc 13 §2): opaque body, plaintext bookkeeping.
   * The hub refuses this on plaintext rooms and refuses plaintext `submit`
   * on encrypted rooms — no mixed-mode documents, enforced both ways. */
  | { t: "submit-enc"; envelope: IntentEnvelope }
  /** Upload a sealed checkpoint (doc 13 §3) — accepted ONLY from the
   * connection the server currently designates as checkpointer. */
  | { t: "checkpoint"; checkpoint: SealedCheckpoint }
  /** Media re-supply (doc 16 §3): "I need blob <sha>" (relay miss). */
  | { t: "media-need"; sha: string }
  /** "I hold these blobs" — reply to media-request, or unsolicited right
   * after a welcome whose mediaNeeded intersects local holdings (§5.4). */
  | { t: "media-have"; shas: string[] }
  /** Encrypted hash gossip (doc 13 §2): an OPAQUE sealed {seq, hash} blob
   * the server relays without reading — divergence detection must not leak
   * even document hashes to a blind server (a hash is a stable content
   * fingerprint: confirmable by anyone holding a guess). */
  | { t: "gossip"; iv: string; ciphertext: string }
  /** Presence. In an ENCRYPTED room this is a `SealedPresence` blob and the
   * hub relays it without looking (#20): the server learns who is pointing,
   * never where. Plaintext rooms keep the structured position, which the hub
   * still clamps at the relay. */
  | { t: "presence"; position: PresencePosition | SealedPresence | null }
  /**
   * APPLICATION-LEVEL LIVENESS PROBE (the reconnect arc). The server answers
   * every one with {@link ServerMessage} `pong` carrying the same nonce.
   *
   * WHY THIS EXISTS AT ALL, given TCP and the WebSocket close handshake: a
   * phone that sleeps mid-session leaves a HALF-OPEN connection. The peer is
   * gone, but nothing tells this end — `onclose` never fires, `onerror` never
   * fires, `readyState` still reads OPEN, and every `send()` succeeds into a
   * socket whose bytes will never arrive. The only way to learn the truth is
   * to ask for a reply and notice it did not come.
   *
   * `nonce` is an opaque round-trip token, echoed back unmodified. It exists
   * so a pong can be matched to the ping it answers — without it a stale pong
   * from before a reconnect could silently satisfy the CURRENT probe and mask
   * a genuinely dead socket. It carries no meaning and is never persisted.
   *
   * DELIBERATELY NOT ACTIVITY. The hub answers this WITHOUT calling
   * `noteActivity`, so a heartbeat cannot hold an abandoned room open — the
   * exact hazard the idle-timeout notes call out ("an activity-based clock is
   * what a keepalive script defeats"). A ping proves the socket is alive; it
   * proves nothing about a human being present, and those are different
   * questions with different consequences.
   */
  | { t: "ping"; nonce: number };

/** Server → client. */
export type ServerMessage =
  | {
      t: "welcome";
      docId: string;
      seq: number;
      snapshot: string;
      /**
       * The stable-id sidecar for the snapshot (plan doc 12 §2, round-4 F10).
       * A snapshot NEVER travels without it: the joiner cannot re-derive the
       * id table from parse order once history contains split-created carried
       * ids (round-2 F1) — it would silently mis-address every later intent.
       */
      sidecar: IdSidecar;
      tail: LogEntry[];
      /**
       * The session EPOCH id (plan doc 12): minted fresh every time a doc is
       * seeded/re-seeded. A resuming client compares it against its bundle's
       * stored genesisId — same ⇒ seamless rejoin (case 1); different ⇒
       * someone re-seeded while it was away (case 2: take server state, keep
       * the local copy per doc 15 lineage — NEVER silently merge epochs).
       */
      genesisId: string;
      /** Session encryption mode (plan doc 13 §6). Clients derive the truth
       * from their link (#k present ⇒ encrypted) and HARD-REFUSE a welcome
       * that contradicts it — the wire value must never downgrade a client. */
      mode: "plaintext" | "encrypted";
      /** The seed's lineage chain (doc 15): lets a rejoining holder decide
       * fast-forward vs fork CLIENT-side. Absent for provider-created and
       * legacy rooms (⇒ every epoch mismatch is a fork, the safe default). */
      lineage?: LineageHead[];
      /** Shas with waiters / outstanding unavailability (doc 16 §3): a
       * joining holder intersects with its local media and volunteers —
       * the mechanism behind "reappears when a holder rejoins". */
      mediaNeeded?: string[];
      /** Per-part media ADDRESSES (doc 16 §6 late-join): the declared sha
       * (+ E2EE iv) for every registered media part, so a joiner can fetch
       * pixels its snapshot references but does not contain — without this
       * the address lives only in the already-folded insertImage intent and
       * late-join media is unfetchable. Plaintext rooms only; encrypted
       * rooms carry the same map INSIDE the sealed checkpoint body (the
       * server must not learn part structure beyond what PUT addresses
       * already reveal). Parse-derived holes with unknown sha are omitted. */
      media?: { part: string; sha: string; iv?: string }[];
      /**
       * The relay's configured per-blob size limit, in bytes, so a client can
       * CHECK A FILE BEFORE UPLOADING IT rather than discovering the limit
       * from a 413 afterwards.
       *
       * The number flows FORWARD — server to client, once per session — which
       * is why it lives here rather than being threaded back through the
       * upload's return type. A refusal has to travel up through every layer
       * between the relay and the file picker, and each of those layers has
       * its own idea of what failure looks like; a value published at join
       * time has no such path to get lost on.
       *
       * NOT A SECRET, and nothing is gated on it: the same number already
       * appears in a 413 body, and anyone can discover it by uploading
       * something large. It is per-SERVER-CONFIG, not per-room, so it carries
       * nothing about the room's contents or participants.
       *
       * Optional and additive: an older server omits it, and a client that
       * gets nothing should skip the pre-check rather than invent a limit —
       * the server still enforces the real one.
       */
      mediaMaxBlobBytes?: number;
    }
  | { t: "broadcast"; entries: LogEntry[] }
  /**
   * Encrypted-mode welcome (doc 13 §3): the seed checkpoint (sealed by the
   * epoch's seeder) + the epoch's envelope log after it. BASE-COMPLETE by
   * construction (round-4 blocker 1): the demo server retains the WHOLE
   * epoch log in RAM (doc 12 — rooms are session-scoped), so every entry a
   * joiner needs to re-derive canonical forms is present; client-produced
   * mid-session checkpoints (doc 13 item 6) are a RAM optimization on top,
   * not a correctness requirement, and land with the retention rule.
   */
  | {
      t: "welcome-enc";
      docId: string;
      genesisId: string;
      checkpoint: SealedCheckpoint;
      tail: EnvelopeEntry[];
      mode: "encrypted";
      mediaNeeded?: string[];
      /** Relay per-blob size limit — see the `welcome` field of the same
       * name. Published identically in encrypted rooms: it is server config,
       * not room content, so a blind sequencer can state it in the clear. */
      mediaMaxBlobBytes?: number;
    }
  /** Encrypted-mode broadcast: sequenced opaque envelopes. */
  | { t: "broadcast-enc"; entries: EnvelopeEntry[] }
  /** Relayed gossip blob; `from` is the sender's BOUND clientId. */
  | { t: "gossip"; from: string; iv: string; ciphertext: string }
  /** Media re-supply control (doc 16 §4). request = broadcast "who has
   * it"; upload = to ONE chosen holder; ready/unavailable = to waiters. */
  | { t: "media-request"; sha: string }
  | { t: "media-upload"; sha: string }
  | { t: "media-ready"; sha: string }
  | { t: "media-unavailable"; sha: string }
  /** Checkpointer designation (doc 13 §3): the SERVER assigns the role —
   * the v1 lowest-clientId election was riggable (round-4 blocker 2); an
   * assigned role can only be held by an authenticated connection the
   * server picked. Rotated on disconnect. */
  | { t: "checkpointer"; active: boolean }
  /** `participant` is the sender's bound clientId (round-4 F14) — the same
   * identifier intents carry — so presence joins the roster/attribution
   * keyspace and survives the sender reconnecting on a new socket. */
  | { t: "presence"; participant: string; position: PresencePosition | SealedPresence | null }
  /** Full roster snapshot, fanned out on every join/leave/profile change
   * (rooms are small; a snapshot beats delta bookkeeping). Ephemeral like
   * presence: never logged, never persisted, dies with the room. */
  | { t: "roster"; roster: RosterEntry[] }
  | { t: "agent-invite-registered"; inviteId: string; expiresAt: number }
  | { t: "agent-invite-refused"; inviteId: string; reason: string }
  | { t: "agent-connected"; inviteId: string; agentClientId: string; profile: ParticipantProfile }
  | {
      t: "agent-chat";
      agentClientId: string;
      messageId: string;
      sender: "inviter" | "agent";
      iv: string;
      ciphertext: string;
    }
  /**
   * The session is going to END, and this is the grace period before it does
   * (server lifecycle arc). Sent once per armed deadline to every connection
   * in the room so the UI can run a countdown — a session that vanishes
   * without warning is indistinguishable from a crash, and the whole point of
   * these timeouts is that they are POLICY, visibly applied.
   *
   *  - `idle`     nobody has done anything for the idle window. CANCELLABLE:
   *               any qualifying activity (a submit, an admin action, a media
   *               transfer, someone joining — never presence) resets the clock
   *               and produces `session-warning-cleared`.
   *  - `lifetime` the room has reached its absolute age cap. NOT cancellable
   *               by anything; the deadline was fixed when the room was
   *               created. Ending it does not end the document — any holder
   *               re-seeds from their bundle into a NEW epoch.
   *
   * `inMs` is the ACTUAL remaining time when the frame was built, not the
   * configured lead time: warnings are emitted by a periodic sweep, so the
   * real remainder is up to one sweep interval less than the configured
   * value, and a client counting down from the constant would lie.
   */
  | { t: "session-warning"; reason: "idle" | "lifetime"; inMs: number }
  /**
   * A previously warned deadline is no longer approaching — the countdown
   * should disappear. Only ever `idle`: the lifetime cap cannot be cancelled,
   * and narrowing the type here means a consumer cannot write a branch the
   * server is incapable of producing.
   */
  | { t: "session-warning-cleared"; reason: "idle" }
  /**
   * The answer to a client `ping`, echoing its nonce. Proof that the socket is
   * end-to-end alive RIGHT NOW — not merely that the local OS still believes
   * in it, which is all `readyState === OPEN` ever establishes.
   *
   * Emitted unconditionally and statelessly: no room lookup, no auth check, no
   * activity bookkeeping. A liveness answer that could itself be refused would
   * be indistinguishable from the death it is supposed to detect.
   */
  | { t: "pong"; nonce: number }
  | { t: "refused"; reason: string };

/** Bump when the intent apply/transform semantics change in a way that makes
 * mixed-version clients diverge (plan doc 03 version lockstep). A client whose
 * version differs is refused at hello with "please refresh".
 * v2: hello carries clientId (bound identity) + takeover; welcome carries the
 * id sidecar.
 *
 * v3: ping/pong liveness (the reconnect arc).
 *
 * THIS ONE LOOKS ADDITIVE AND IS NOT, and the reason is worth stating because
 * the instinct to skip the bump is strong and wrong. Neither dispatch switch
 * has a `default:` clause — the hub's (hub.ts) and both clients'
 * (connection.ts / enc-connection.ts) — so an unknown `t` is silently dropped
 * at every hop. That cuts in opposite directions per side:
 *
 *  - An OLD CLIENT receiving `pong` ignores it. Harmless.
 *  - An OLD SERVER receiving `ping` ignores it AND SENDS NOTHING BACK.
 *
 * The second case is fatal to the feature. The heartbeat's entire value is
 * that a missing pong is CONCLUSIVE evidence the socket is dead — that is the
 * only signal that catches the half-open sleeping phone, where onclose and
 * onerror never fire. Against a v2 server every pong is missing, so a new
 * client would diagnose a perfectly healthy session as lost and gate the
 * editor read-only: a silent-data-loss bug traded for a can't-type-at-all bug.
 * Weakening the client to treat silence as inconclusive would restore the very
 * ambiguity the probe exists to remove.
 *
 * Because hello ALREADY hard-refuses a protocolVersion mismatch, bumping makes
 * "I am connected" imply "my peer answers pings" by construction, and a missing
 * pong stays unambiguous. The cost is that v2 clients are refused with "please
 * refresh" — the mechanism working as designed, and the same price every prior
 * bump paid.
 *
 * The v3 liveness change did not bump ENGINE_VERSION because liveness frames
 * cannot change document bytes. Version 4 adds rejectAllRevisions and uses
 * engine e8. Version 5 adds tracked paragraph splits and uses engine e9. */
export const PROTOCOL_VERSION = 6;
