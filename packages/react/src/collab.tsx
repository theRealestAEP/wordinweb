import { useEffect, useMemo, useRef, useState, createElement, type ReactNode } from "react";
import type { DocxDocument, EncodedCaret } from "@wordinweb/core";
import { DocxView, type DocxViewApi } from "./index.js";
import { DocxToolbar, type ToolbarFeature, type ToolbarMode } from "./toolbar.js";
import {
  CollabConnection,
  EncryptedCollabConnection,
  stretchShareCode,
  toSuggestions,
  arrivalMode,
  OFFLINE_TAIL_CAP,
  applyIntentScoped,
  resyncScope,
  createWebSocketTransport,
  monitorTransport,
  BundlePersister,
  draftKey,
  supersededKey,
  type ConnectionState,
  type LivenessOptions,
  type BundleStore,
  type ParticipantProfile,
  type RosterEntry,
  type Intent,
  type PresencePosition,
  type UndoOutcome,
  type WriteStatus,
} from "@wordinweb/collab/client";

/**
 * React binding for wordinweb collaboration. Imported from the SEPARATE
 * `wordinweb/collab` entry so that non-collab apps never pull the collab
 * engine into their bundle (plan doc 07: unreachable beats shakeable — a
 * local-only `import { DocxView } from "wordinweb"` has no path to this file).
 */

export interface UseCollabOptions {
  /** WebSocket URL of the collab server (ws://… / wss://…). */
  url: string;
  /** Document id to join (the magic-link id, plan doc 11). */
  docId: string;
  /** This client's stable id (persist per browser/tab). */
  clientId: string;
  /** Optional auth token (JWT minted by the app's backend, plan doc 07). */
  token?: string;
  /** Construct the socket. Defaults to `new WebSocket(url)`; injectable for
   * tests / custom transports. */
  createSocket?: (url: string) => WebSocket;
  /**
   * Bundle persistence + resume (plan doc 12 §4/§5). When set, the hook:
   * resumes from a stored bundle if one exists (replaying pending intents,
   * idempotently) instead of joining cold; persists the confirmed state on a
   * ~1s throttle after every change, plus a best-effort flush on pagehide/
   * hidden; and on an epoch change (someone re-seeded while away) preserves
   * the superseded bundle as a draft (`<docId>#draft-<oldGenesis>`) before
   * adopting the server's state — the fork rule made visible, never a silent
   * merge. Browser apps pass an `IndexedDbBundleStore` (from `wordinweb`);
   * tests inject the in-memory store. Without it, behavior is unchanged
   * (join cold, keep nothing).
   */
  store?: BundleStore;
  /** Claim the identity from an existing live connection (doc 12 §7 "use
   * here instead"): set after an `already-open` refusal and remount. */
  takeover?: boolean;
  /**
   * E2EE mode (doc 13): the document master key from the link's `#k=`
   * fragment. Its PRESENCE selects the encrypted connection — mode is
   * derived from the link, never the wire; a plaintext welcome is
   * hard-refused (`mode-downgrade`). */
  docKey?: string;
  /** Share code (doc 13 §7) when the doc has one — stretched client-side
   * and mixed into key derivation + sent as the hello proof. */
  shareCode?: string;
  /** Owner capability token (doc 14 §2.5): held by the seeder (returned by
   * go-live, kept in their bundle), never in the shared link. Presence of
   * a valid token unlocks the admin controls. */
  ownerToken?: string;
  /**
   * HTTP origin of the media relay (plan doc 16 §3), e.g.
   * "http://localhost:1234" — the same server the WebSocket points at.
   * Image bytes travel over these routes and NEVER over the sequencer.
   * Without it the room has no media duties and the image toolbar stays
   * inert, so an app that doesn't want media simply omits it.
   */
  httpBase?: string;
  /** Display profile sent at join (doc 14 §2) — self-asserted; persist it in
   * localStorage next to the clientId so identity is stable per browser. */
  profile?: ParticipantProfile;
  /**
   * Heartbeat + reconnect tuning. Defaults suit a real network (15s probe,
   * 10s answer deadline, 6 retries); tests inject short values and a fake
   * clock. `maxRetries` is how many backoff attempts run before the session
   * gives up and reports `lost` — the point at which a human has to act.
   */
  liveness?: LivenessOptions & { maxRetries?: number };
  /**
   * Bound on the offline tail (doc 15 §4.3) — how many intents may be
   * recorded while disconnected before the editor gates read-only with an
   * explicit "offline limit" state. Defaults to OFFLINE_TAIL_CAP; injectable
   * so tests can reach the cap without recording thousands of edits.
   */
  offlineTailCap?: number;
}

export interface CollabSession {
  /** The live document to render (null until the welcome arrives). */
  doc: DocxDocument | null;
  /** Monotonically increases on every reconciled change — a cheap re-render
   * signal (the doc object may be reloaded by reconciliation). */
  version: number;
  /**
   * Increases only when the RENDERED DOCUMENT changed outside the editor —
   * remote applies, optimistic toolbar ops, reloads, self-heals, media
   * installs. The editor-driven typing path (submit + own echo) leaves it
   * untouched: the editor already mutated AND painted the doc, so repainting
   * again on `version` queued a redundant whole-document relayout per
   * keystroke — invisible on 5 pages, catastrophic past DocxView's
   * background-layout threshold (a 500-page doc relaid out asynchronously
   * behind an inert container, per keystroke). DocxView's `renderSignal`
   * rides THIS; `version` remains the observer/bookkeeping signal.
   */
  renderVersion: number;
  /** Increases only when reconciliation RELOADED the document (a true
   * conflict). The editor re-mounts on this; between reloads it updates in
   * place (no flash for the common non-conflicting edits). */
  docEpoch: number;
  /** True once joined and the snapshot is loaded. */
  ready: boolean;
  /** Submit a local edit the editor ALREADY applied to the live doc
   * (bookkeeping filled by the connection; no local re-apply). */
  submit: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void;
  /** Submit an operation NOT yet applied locally: it is applied optimistically
   * through the same canonical applyIntent code the server runs, so the local
   * result is byte-identical to every replica by construction. Used for
   * toolbar/API commands (insert chart, set link, page layout, ...). */
  submitOp: (intent: Omit<Intent, "clientId" | "clientSeq" | "base">) => void;
  /** Broadcast this client's cursor/selection. */
  setPresence: (pos: PresencePosition | null) => void;
  /** Allocate carried node ids (sub-range format / split / insert). */
  allocIds: (n: number) => number[];
  /** Remote participants' latest cursor/selection positions. */
  presence: Record<string, PresencePosition | null>;
  /** Set if the server refused the connection (e.g. version mismatch). */
  refused: string | null;
  /**
   * WHETHER THIS CLIENT CAN REACH THE SEQUENCER AT ALL.
   *
   * Distinct from `refused`, which means the server answered and said no —
   * this means nobody answered. A refusal ends the session; this is a
   * transport fault that usually heals by itself.
   *
   *  - `live`         the last heartbeat round trip completed.
   *  - `reconnecting` the socket dropped; a backoff retry is scheduled. Most
   *                   drops pass through here and back to `live` unnoticed,
   *                   so UI should stay quiet or subtle here.
   *  - `lost`         retries exhausted. Show something modal: the document is
   *                   safe but nothing further will sync until a human acts.
   *
   * A NON-LIVE CONNECTION DOES NOT BLOCK WRITES (doc 12 §5: offline editing
   * is first-class). "I cannot reach the server" is a different fact from
   * "the server said no": nothing typed offline is doomed — it applies
   * locally, is recorded to the offline tail, persists with the bundle, and
   * replays on reconnect (doc 15's arrival ladder). {@link offline} is the
   * state a consumer renders for it; {@link writesBlocked} stays reserved
   * for genuine refusals, where accepting keystrokes would be a lie.
   */
  connection: ConnectionState;
  /**
   * OFFLINE EDITING IS ACTIVE (doc 15 §2 / doc 12 §5): the sequencer is
   * unreachable, a document is in hand, and nothing has refused this
   * client's writes — so edits apply locally and accumulate in the durable
   * offline tail instead of being dropped.
   *
   *  - `editsHeld` how many intents the tail currently holds. Show it: the
   *    difference between "editing is paused" and "your changes are being
   *    kept (N so far)" is the whole point of the state.
   *  - `capped` the tail hit its bound (doc 15 §4.3). Recording has STOPPED
   *    and {@link writesBlocked} is raised — loudly, with this flag saying
   *    why, never by silently discarding keystrokes. The escape hatches are
   *    reconnecting or downloading a copy.
   *
   * Null while the connection is live, and null when a refusal already
   * blocks writes (a viewer's link does not become writable by going
   * through a tunnel). On reconnect the tail replays: silently as ordinary
   * edits when the session is the same epoch, else via {@link arrival}.
   */
  offline: { editsHeld: number; capped: boolean } | null;
  /**
   * Retry the connection immediately, resetting the backoff and the attempt
   * budget. Wired to a "Try again" control on the `lost` UI — from `lost`
   * nothing else will ever retry, because exhausting the budget is precisely
   * the statement that automatic recovery has stopped.
   */
  reconnect: () => void;
  /**
   * The server announced it will END this session, and this is the grace
   * period before it does. `inMs` is the measured remainder at the moment the
   * warning arrived — tick a countdown from it locally rather than expecting
   * refreshed messages.
   *
   *  - `idle`     cancellable: any qualifying activity (an accepted edit, an
   *               admin action, a media transfer, a join — never presence)
   *               resets the clock and this goes back to null on its own.
   *  - `lifetime` NOT cancellable by anything; it counts down to the ending.
   *
   * Null whenever no deadline is approaching — including after ANY session
   * end (kick, refusal), because a countdown that outlived its session would
   * keep promising an ending that already happened.
   */
  sessionWarning: { reason: "idle" | "lifetime"; inMs: number } | null;
  /** True after an edit was refused under the owner's read-only lock
   * (doc 14 §2.5). NON-FATAL: the live view keeps working.
   *
   * LEGACY/FALLBACK now: sticky until reload or `retryWrites`, because a
   * refusal can only ever say "blocked". Prefer {@link writesBlocked}, which
   * reads the server's positive roster status when one is published and only
   * falls back to this against an older server. */
  readOnlyBlocked: boolean;
  /**
   * The server has told this client it is NOT the owner of this document.
   *
   * Holding an `ownerToken` string is not the same as it being valid — a room
   * re-seeded into a new epoch mints a fresh one, leaving the old token
   * truthy in a browser that now has no admin rights. The welcome carries no
   * owner flag, so until an admin op is attempted this is genuinely unknown;
   * a `not-owner` refusal is the authoritative answer.
   *
   * Consumers should HIDE admin controls when this is true rather than
   * disable them: an offered control that cannot work is worse than an absent
   * one, and the refusal it produces is not a session failure.
   */
  notOwner: boolean;
  /**
   * THE SERVER WILL NOT ACCEPT THIS CLIENT'S WRITES. Consumers must render the
   * editor READ-ONLY on this, never merely annotate it: an editable surface
   * over a server that refuses the writes applies every keystroke locally and
   * then loses it, which is silent data loss that looks like it worked.
   *
   * One predicate rather than a per-reason check, so a new blocking state
   * inherits the gate instead of needing to remember it.
   *
   * NOW READS THE SERVER'S ROSTER STATUS (`RosterEntry.write`), which closed
   * the seam this comment used to describe. That matters in both directions:
   * the status is present AT JOIN, so edit 1 is gated rather than applied and
   * then healed away; and it is POSITIVE, so a lift arrives on its own instead
   * of leaving the client in viewer mode until reload. Against a server that
   * publishes no status it falls back to the refusal-derived flag — never to
   * "allowed", which would restore both bugs at once.
   *
   * The status also distinguishes the three server conditions the single
   * `read-only` refusal could not (owner lock, per-client demotion, viewer
   * token), so a UI can finally say which one applies rather than inventing a
   * distinction it could not see.
   *
   * A DEAD CONNECTION IS DELIBERATELY NOT A BLOCK — that is the refusal /
   * unreachable split (doc 12 §5). This predicate means "nothing you type
   * can ever land"; a dropped socket means "nothing you type can land YET",
   * and the offline tail is what keeps the two honest: disconnected edits
   * apply locally, persist durably, and replay on reconnect, so accepting
   * them is not a lie. The one offline condition that does fold in is the
   * tail CAP ({@link offline}.capped): past it an edit genuinely cannot be
   * kept, so the gate closes loudly rather than recording silently stopping.
   * An earlier revision folded every non-live connection in here; that
   * over-reached — it was written for server refusals and extended to
   * transport faults, which made first-class offline editing impossible.
   *
   * A REFUSED SESSION FOLDS IN TOO, by the same rule. A refusal ends the
   * session — no socket, nothing to sequence against — so any document still
   * on screen (e.g. behind a "session ended" dialog) must be read-only, and
   * it inherits that from this predicate rather than from a refusal-specific
   * gate someone could forget to compose.
   */
  writesBlocked: boolean;
  /**
   * WHY the server refuses this client's writes, straight from the roster.
   *
   * UNDEFINED means the server published nothing (an older build), NOT that
   * writing is allowed — the same contract as the media limit's null. A
   * consumer must keep its copy generic in that case rather than naming a
   * cause it was never told.
   *
   * The three causes are genuinely different to the person reading the banner:
   * `owner-lock` may lift at any moment, `demoted` is about them specifically,
   * and `viewer-role` is a property of their link that nothing the owner does
   * in-session will change. Telling that last group to wait for the owner is
   * telling them to wait for something that will never happen.
   */
  writeStatus?: WriteStatus;
  /**
   * Optimistically clear the FALLBACK block and let the user try again.
   *
   * BELT AND BRACES now, not the primary escape: where the server publishes a
   * write status, a lift arrives on its own and the editor becomes writable
   * without anyone clicking anything. This remains for the older-server
   * fallback path — where the block is sticky, no lift is announced, and
   * attempting a write is the only way to discover the lock is gone — and as a
   * manual override if a status update is ever missed.
   */
  retryWrites: () => void;
  /**
   * Set when a resume landed in a different epoch than the stored bundle
   * (doc 12 §5 case 2): the session took the server's state; the old copy
   * was saved as a draft. UI copy: "restored by another participant — your
   * offline copy is saved as a draft." Null otherwise.
   */
  epochChanged: { from: string; to: string } | null;
  /** Session roster (doc 14 §2): everyone who joined this session, with
   * connection state — the identity keyspace presence/attribution share. */
  roster: RosterEntry[];
  /** Rename/recolor this participant (server sanitizes + fans out). */
  setProfile: (profile: ParticipantProfile) => void;
  /** Attribution layer 1 (doc 14 §3): recent applied entries
   * {seq, clientId, kind} — join clientId to `roster` for names/colors. */
  activity: { seq: number; clientId: string; kind: string }[];
  /** Owner admin ops (doc 14 §2.5) — no-op unless this connection proved
   * the owner token; the server refuses `not-owner` otherwise. */
  admin: (action:
    | { op: "kick"; clientId: string }
    | { op: "readOnly"; on: boolean }
    | { op: "setRole"; clientId: string; role: "editor" | "viewer" }) => void;
  /**
   * Offline reconciliation (doc 15 arrival ladder). When a rejoin lands in
   * a DIFFERENT epoch than the tail was recorded against, the recommended
   * mode is `suggest` (replay as tracked changes; ≤ the suggest threshold)
   * or `draft` (large tails). `structural` counts the tail intents that
   * have no suggestion form (splits/merges/tables/format — rebase.ts):
   * they are NOT replayed by `suggest` and survive in the banked draft
   * slot instead — say so, or their absence reads as silent loss. A
   * same-epoch rejoin never surfaces here: its tail replays silently
   * through the ordinary submit path, like a large pending queue.
   * Null when there's nothing to reconcile. `reconcile()` runs the choice.
   */
  arrival: { mode: "suggest" | "draft"; tailLength: number; structural: number } | null;
  reconcile: (mode: "suggest" | "draft") => void;
  /** How many times the connection self-healed a drifted optimistic replica
   * (encrypted mode's quiescent mirror check — the B6a class). Telemetry;
   * the heal itself is automatic and already reflected in doc/docEpoch. */
  selfHeals: number;
  /**
   * Submits the connection DROPPED because it was not ready yet (no replica
   * to apply against, no confirmed seq, no key to seal with). Should stay 0:
   * the editor is gated on `ready`, so nothing user-driven can reach that
   * guard. A non-zero value means some caller is submitting before the
   * welcome — which is a silent edit loss, and was the shape of B13's
   * 231-intent gap. Telemetry, deliberately visible rather than swallowed.
   */
  droppedPreReady: number;
  /**
   * Persistence writes that FAILED (quota exceeded, blocked storage, private
   * mode). Non-zero means this browser's stored bundle is stale — and in a
   * zero-custody design that bundle IS the durable copy, so "saved" is
   * currently a lie. Consumers should SHOW this, not just log it: the user
   * can still rescue the document with the download button, but only if they
   * are told. Sticky (never decremented): the doubt persists until reload.
   */
  persistErrors: number;
  /**
   * Submits whose seal or transport send THREW after the edit was already
   * applied optimistically — i.e. lost edits, counted at the far end of the
   * path from droppedPreReady.
   *
   * The two together are a PARTITION, which is the point: droppedPreReady is
   * "refused before it was ever applied", this is "accepted locally then lost
   * on the way out". A swarm run that loses N intents can now subtract both
   * from N and see what remains unexplained, instead of a night of guessing —
   * see B13, where exactly that arithmetic was missing.
   */
  sendFailures: number;
  /** Upload image bytes to the relay and get the address fields an
   * insertImage intent must carry (doc 16 §5.1). Null = the relay refused,
   * and the caller must not reserve anything. Always null when the app
   * supplied no `httpBase`. */
  uploadMedia: (bytes: Uint8Array) => Promise<{ blobSha: string; bytesLen: number; iv?: string } | null>;
  /**
   * Largest single upload the relay accepts, in bytes, as published in the
   * welcome — so an oversized file can be refused locally instead of after a
   * full seal-hash-upload round trip.
   *
   * NULL MEANS "NO PUBLISHED LIMIT", and the only correct response to it is to
   * skip the check: a server that publishes nothing still enforces its real
   * limit, so a client that invents a default either blocks uploads that
   * would have succeeded or promises a size that will be refused.
   */
  mediaMaxBlobBytes: number | null;
  /**
   * Reverse this participant's last SEQUENCED action (plan doc 03 Phase 8).
   * The inverse is computed locally and submitted as an ORDINARY intent, so
   * it converges, broadcasts, and — if the target has since been deleted or
   * changed — is rejected cleanly on every replica alike.
   *
   * The outcome is reported rather than thrown so the UI can say the right
   * thing: `cannot-undo` means the last action has no inverse yet (a
   * different message from having nothing to undo), and `unavailable` means
   * this connection has no collaborative undo at all — today the plaintext
   * connection, whose authority lives on the server.
   */
  undoLast: () => UndoOutcome;
}


/**
 * Connect to a collab document and track its live state. Returns a
 * CollabSession the app injects into `<DocxView collab={session} />` (the
 * DocxView `collab` prop is typed structurally, so the react package needs no
 * runtime dependency on this module).
 */
export function useCollab(opts: UseCollabOptions): CollabSession {
  const { url, docId, clientId, token, createSocket, store, profile, takeover, docKey, shareCode, ownerToken, httpBase } = opts;
  const connRef = useRef<CollabConnection | null>(null);
  const [doc, setDoc] = useState<DocxDocument | null>(null);
  const [version, setVersion] = useState(0);
  /** Repaint signal (see CollabSession.renderVersion): follows the
   * connection's docVersion, plus local bumps for offline toolbar applies. */
  const [renderVersion, setRenderVersion] = useState(0);
  /** Last connection docVersion folded into renderVersion — the two sources
   * (connection counter, local offline bumps) must not collide. */
  const lastDocVersionRef = useRef(0);
  const [docEpoch, setDocEpoch] = useState(0);
  const [ready, setReady] = useState(false);
  const [presence, setPresence] = useState<Record<string, PresencePosition | null>>({});
  const [refused, setRefused] = useState<string | null>(null);
  /**
   * Announced endings as ABSOLUTE deadlines, one slot per reason.
   *
   * Not a single slot, because the two deadlines can be pending at once (a
   * room nearing its age cap that also goes quiet), and a single slot loses
   * one of them: the idle warning would overwrite the lifetime warning, and
   * the `cleared` that follows an edit would then take the lifetime countdown
   * down with it — leaving no warning at all for an ending that is still
   * minutes away and cannot be cancelled.
   *
   * Absolute rather than remaining, so the surviving deadline resumes with the
   * time it actually has left rather than the remainder it was announced with.
   */
  const [deadlines, setDeadlines] = useState<{ idle: number | null; lifetime: number | null }>({ idle: null, lifetime: null });
  /** Both slots empty, without re-rendering when they already are. */
  const clearDeadlines = () =>
    setDeadlines((d) => (d.idle === null && d.lifetime === null ? d : { idle: null, lifetime: null }));
  /**
   * The SOONEST pending ending — that is the one the user needs, and when it
   * is the cancellable one, cancelling it reveals the other rather than
   * clearing the screen. `inMs` is measured here and only here: the memo
   * recomputes exactly when a deadline is set or cleared, so the number the
   * consumer starts its countdown from is fresh at the moment it changes, and
   * the object identity is stable in between (a consumer keying an interval on
   * it must not be restarted by an unrelated render).
   */
  const sessionWarning = useMemo<{ reason: "idle" | "lifetime"; inMs: number } | null>(() => {
    const { idle, lifetime } = deadlines;
    const reason = idle !== null && (lifetime === null || idle <= lifetime) ? "idle" : lifetime !== null ? "lifetime" : null;
    if (!reason) return null;
    return { reason, inMs: Math.max(0, (reason === "idle" ? idle! : lifetime!) - Date.now()) };
  }, [deadlines]);
  /** True after an edit was refused `read-only` (owner lock, doc 14 §2.5).
   * Sticky until reload — there is no lift signal on the wire yet, and the
   * lock only re-manifests per-edit; the banner is advisory while the live
   * view keeps working. */
  const [readOnlyBlocked, setReadOnlyBlocked] = useState(false);
  const [notOwner, setNotOwner] = useState(false);
  const [epochChanged, setEpochChanged] = useState<{ from: string; to: string } | null>(null);
  const [roster, setRoster] = useState<RosterEntry[]>([]);
  const [arrival, setArrival] = useState<{ mode: "suggest" | "draft"; tailLength: number; structural: number } | null>(null);
  const [selfHeals, setSelfHeals] = useState(0);
  const [droppedPreReady, setDroppedPreReady] = useState(0);
  const [persistErrors, setPersistErrors] = useState(0);
  const [sendFailures, setSendFailures] = useState(0);
  /**
   * THE OFFLINE TAIL (doc 15 §2): intents recorded while the sequencer was
   * unreachable. The ref — not the connection — owns it, because the tail
   * must survive the teardown-and-rebuild cycle every reconnect performs.
   * Entries carry no wire bookkeeping; the replay's ordinary submit path
   * assigns clientId/clientSeq/base when they finally go out.
   */
  const offlineTailRef = useRef<import("@wordinweb/collab/client").DocBundle["offlineTail"]>(undefined);
  /** The genesisId the tail was recorded against (bundle.offlineTailEpoch):
   * same epoch on rejoin ⇒ silent replay; different ⇒ arrival ladder. */
  const offlineEpochRef = useRef<string | null>(null);
  /** Tail length mirrored into state so the UI re-renders as it grows. */
  const [offlineHeld, setOfflineHeld] = useState(0);
  /** The genesis of the state on screen — what a capture stamps the tail
   * with when recording starts. Set on every welcomed onChange. */
  const lastGenesisRef = useRef<string | null>(null);
  /** The live persister, reachable from capture/replay code that outlives
   * any single effect run. Null while no store or between effect runs. */
  const persisterRef = useRef<BundlePersister | null>(null);
  /** The connection whose welcome already resolved the tail (replay started
   * or arrival offered) — the welcome-time decision must run once per
   * connection, not once per broadcast. */
  const tailConnRef = useRef<object | null>(null);
  /** Trailing-throttle timer for the OFFLINE tail writer (below). */
  const offlinePersistTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [connection, setConnection] = useState<ConnectionState>("live");
  /**
   * Bumped to force the connection effect to tear down and rebuild — the
   * reconnect itself.
   *
   * REBUILDING RATHER THAN RE-HANDSHAKING THE EXISTING CONNECTION is the
   * whole design. The teardown flushes the bundle and the fresh run reads it
   * back, so a reconnect travels the ORDINARY resume path — the one doc 12 §5
   * specifies and the bundle-resume suite already covers — with
   * `sinceSeq = bundle.confirmedSeq` and the pending queue replayed
   * idempotently by (clientId, clientSeq). Swapping a socket underneath a live
   * connection would instead need a second, bespoke rejoin path that no
   * existing test exercises, to re-derive state the bundle already holds.
   */
  const [reconnectNonce, setReconnectNonce] = useState(0);
  /** Backoff + attempt budget, in refs: they must survive the effect teardown
   * that the reconnect itself performs. */
  const backoffRef = useRef(0);
  const attemptRef = useRef(0);
  const retryTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  /**
   * The teardown's bundle flush, awaited by the next effect run before it
   * reads the store.
   *
   * WITHOUT THIS THE RECONNECT RACES ITS OWN SAVE: `flush()` is async and the
   * teardown cannot await it, so the rebuilt connection could call
   * `store.get()` first and resume from a bundle up to one throttle window
   * stale — re-sending confirmed intents and, on a store with nothing written
   * yet, joining COLD when a resume was available.
   */
  const flushDoneRef = useRef<Promise<void>>(Promise.resolve());
  const maxRetries = opts.liveness?.maxRetries ?? 6;
  // Read through a ref so the effect does not list an options object literal
  // (a fresh identity every render) in its deps and reconnect forever — the
  // same hazard the `profile` note below records.
  const livenessRef = useRef(opts.liveness);
  livenessRef.current = opts.liveness;

  /** Retry now: clear the budget and rebuild. Used by the `lost` UI and by
   * the wake-up path, both of which mean "a human is here, try again". */
  const reconnect = () => {
    if (retryTimerRef.current !== null) clearTimeout(retryTimerRef.current);
    retryTimerRef.current = null;
    backoffRef.current = 0;
    attemptRef.current = 0;
    setConnection("reconnecting");
    setReconnectNonce((n) => n + 1);
  };

  /**
   * THE OFFLINE TAIL WRITER — persistence for the window where the
   * BundlePersister cannot write (a rebuilt connection has no welcome yet,
   * so exportBundle is null). Read-modify-write of the STORED bundle,
   * chained on flushDoneRef so it serializes against teardown flushes and
   * the next run's resume read. While a welcomed connection exists the
   * persister carries the tail instead (via its offlineTail getter), so the
   * two writers are never active at once.
   */
  const persistOfflineTail = (): void => {
    if (!store) return;
    flushDoneRef.current = flushDoneRef.current
      .then(async () => {
        const prev = await store.get(docId);
        if (!prev) return; // never persisted: nothing durable to attach to
        const tail = offlineTailRef.current ?? [];
        const next = { ...prev };
        if (tail.length) {
          next.offlineTail = [...tail];
          next.offlineTailEpoch = offlineEpochRef.current ?? undefined;
        } else {
          delete next.offlineTail;
          delete next.offlineTailEpoch;
        }
        await store.put(next);
      })
      .catch((err: unknown) => {
        // Same contract as the persister's onError: in a zero-custody design
        // a failed write means the durable copy is stale — never silent.
        setPersistErrors((n) => n + 1);
        // eslint-disable-next-line no-console
        console.error("[wordinweb] offline-tail-persist", err);
      });
  };
  /** Throttled (trailing, 1s) — the doc 12 §4 discipline: bounds loss to
   * the window without deferring forever under sustained typing. */
  const scheduleOfflinePersist = (): void => {
    if (!store || offlinePersistTimerRef.current !== null) return;
    offlinePersistTimerRef.current = setTimeout(() => {
      offlinePersistTimerRef.current = null;
      persistOfflineTail();
    }, 1000);
  };
  /** Run any armed offline write NOW (pagehide, teardown). */
  const flushOfflinePersist = (): void => {
    if (offlinePersistTimerRef.current === null) return;
    clearTimeout(offlinePersistTimerRef.current);
    offlinePersistTimerRef.current = null;
    persistOfflineTail();
  };

  /**
   * Record one offline edit (doc 15 §2): already applied to the live doc by
   * the editor (`preApplied`), or applied here through the SAME canonical
   * applyIntent code every replica runs (the toolbar/API path). The intent
   * then joins the tail, the UI repaints, and the tail is persisted through
   * whichever writer can currently reach the store.
   */
  const captureOffline = (intent: Omit<Intent, "clientId" | "clientSeq" | "base">, preApplied: boolean): void => {
    if (!preApplied) {
      if (!doc) return;
      const ids = doc.enableStableIds();
      const res = applyIntentScoped(doc, ids, intent as Intent);
      if (!res.applied) return; // nothing changed ⇒ nothing to record
      resyncScope(doc, ids, res);
      // Applied HERE, not by the editor — the repaint signal must move.
      setRenderVersion((v) => v + 1);
    }
    const tail = offlineTailRef.current ?? [];
    tail.push(intent as Intent);
    offlineTailRef.current = tail;
    // Stamp the epoch the tail is being recorded against, once per tail.
    offlineEpochRef.current ??= lastGenesisRef.current;
    setOfflineHeld(tail.length);
    setVersion((v) => v + 1); // offline edits fire no onChange — repaint here
    if (connRef.current?.ready) persisterRef.current?.notify();
    else scheduleOfflinePersist();
  };

  /**
   * Drain the offline tail through the ordinary submit path, one-in-flight
   * (doc 03's supported discipline — the reason no sibling math is needed).
   * `plain` replays entries as the ordinary edits they were (same-epoch
   * rejoin: to the session this client is indistinguishable from one that
   * was disconnected with a large pending queue — doc 15 §2). `suggest`
   * converts each entry to its tracked-change form first (diverged rejoin,
   * the owner-adopted default); structural entries have no suggestion form
   * and are popped WITHOUT submitting — they survive in the banked draft
   * slot, and `arrival.structural` told the user so.
   *
   * EXACTLY-ONCE across a kill: each step pops the entry from the tail,
   * submits it (it becomes a pending intent with a clientSeq), and flushes
   * the persister — so the stored bundle moves the intent from `offlineTail`
   * to `pending` in one write. A resume then replays it from `pending`,
   * verbatim, deduplicated by (clientId, clientSeq). A hard kill inside the
   * sub-second flush window keeps the ordinary doc 12 §4 RPO bound.
   */
  const runTailReplay = (conn: CollabConnection, kind: "plain" | "suggest"): void => {
    const author = profile?.name || clientId;
    const date = new Date().toISOString();
    const step = (): void => {
      // Torn down or dropped again: stop. The remainder is persisted and
      // resumes on the next welcome (same epoch now ⇒ silent continuation;
      // suggest remainders re-offer via the arrival ladder).
      if (connRef.current !== conn || !conn.ready) return;
      if (conn.pendingCount > 0) {
        setTimeout(step, 25);
        return;
      }
      const tail = offlineTailRef.current ?? [];
      if (!tail.length) {
        // Done. The stored tail is already erased: the final step's flush
        // (below) wrote the bundle after the last pop emptied the getter.
        offlineTailRef.current = undefined;
        offlineEpochRef.current = null;
        setOfflineHeld(0);
        setArrival(null);
        return;
      }
      const next = tail[0];
      offlineTailRef.current = tail.slice(1);
      setOfflineHeld(tail.length - 1);
      if (kind === "suggest") {
        const { suggestions } = toSuggestions([next], author, date);
        if (suggestions.length) conn.submit(suggestions[0]);
      } else {
        conn.submit(next);
      }
      void persisterRef.current?.flush();
      setTimeout(step, 0);
    };
    step();
  };

  useEffect(() => {
    // Every dep change here means a DIFFERENT connection, so a deadline
    // announced by the previous one no longer applies to anything.
    clearDeadlines();
    const socket = (createSocket ?? ((u: string) => new WebSocket(u)))(url);
    let disposed = false;
    /** One verdict per socket: close, error and heartbeat-timeout all funnel
     * here and only the first is acted on. */
    let dropHandled = false;
    // Assigned immediately below; held in a `let` so `dropped` can stop it
    // without a temporal-dead-zone reference to a const declared after it.
    let monitor: { stop: () => void; start: () => void; probe: () => void } | null = null;
    const dropped = (): void => {
      if (dropHandled || disposed) return;
      dropHandled = true;
      monitor?.stop();
      // Close it ourselves. A half-open socket is still OPEN as far as the
      // browser is concerned, and leaving it around means the next reconnect
      // adds a second live socket for the same clientId — which the hub's
      // single-live-connection rule would refuse `already-open`.
      try {
        socket.close?.();
      } catch {
        // Already dead; the reconnect below is what matters.
      }
      // Budget exhausted ⇒ stop retrying and say so. Retrying forever would
      // leave a phone in a tunnel burning battery on a socket that will not
      // open, and would never give the user the modal that tells them their
      // work is safe.
      if (attemptRef.current >= maxRetries) {
        setConnection("lost");
        return;
      }
      attemptRef.current += 1;
      // Backoff matching scheduleRateLimitRedrive (enc-connection.ts): 300ms
      // doubling to a 5s ceiling, reset on success.
      backoffRef.current = Math.min(backoffRef.current === 0 ? 300 : backoffRef.current * 2, 5000);
      setConnection("reconnecting");
      retryTimerRef.current = setTimeout(() => {
        retryTimerRef.current = null;
        setReconnectNonce((n) => n + 1);
      }, backoffRef.current);
    };
    // ONCLOSE/ONERROR AS PROPERTIES, not addEventListener: the property form
    // needs no matching removeEventListener in the teardown and cannot
    // double-register if this effect ever re-runs against the same socket.
    socket.onclose = () => dropped();
    socket.onerror = () => dropped();
    // The heartbeat catches what those two cannot — see LivenessMonitor: a
    // suspended phone leaves a half-open socket where NEITHER handler ever
    // fires and readyState still reads OPEN.
    const monitored = monitorTransport(
      createWebSocketTransport(socket),
      () => dropped(),
      livenessRef.current ?? {},
    );
    const transport = monitored.transport;
    monitor = monitored.monitor;
    // Probe the moment the tab wakes. This is the motivating scenario, and
    // waiting up to a full probe interval to discover a socket that died
    // during sleep is precisely the window in which the user types into a
    // void. Registered unconditionally (the persister's own visibilitychange
    // listener below is separate and store-gated).
    const onVisible = (): void => {
      if (typeof document !== "undefined" && document.visibilityState !== "visible") return;
      monitor?.probe();
    };
    if (typeof document !== "undefined") document.addEventListener("visibilitychange", onVisible);
    monitor.start();
    let persister: BundlePersister | null = null; // assigned below iff store
    const callbacks: import("@wordinweb/collab/client").ConnectionCallbacks = {
      onChange: () => {
        const c = connRef.current;
        if (!c) return;
        setDoc(c.doc);
        setReady(c.ready);
        setVersion((v) => v + 1); // signal a re-render on every reconciled change
        // Repaint only when the DOC changed outside the editor (docVersion
        // moved) — an own-echo onChange must not queue a whole-document
        // relayout (see CollabSession.renderVersion).
        const dv = c.docVersion;
        if (dv !== lastDocVersionRef.current) {
          lastDocVersionRef.current = dv;
          setRenderVersion((v) => v + 1);
        }
        setDocEpoch(c.docEpoch); // bumps only on a reload (true conflict)
        persister?.notify(); // throttled bundle write (doc 12 §4)
        // A welcome that LANDED is the only proof a reconnect worked. Reset
        // the budget here rather than when the socket opens: an open socket
        // that is then refused (`already-open`, version mismatch) has not
        // recovered anything, and crediting it would let a refusal loop retry
        // forever without ever reaching `lost`.
        if (c.ready) {
          backoffRef.current = 0;
          attemptRef.current = 0;
          setConnection((s) => (s === "live" ? s : "live"));
          lastGenesisRef.current = c.genesisId;
          // Resolve the offline tail ONCE per welcomed connection (doc 15
          // arrival ladder). Same epoch as the tail was recorded against ⇒
          // this client is just a rejoiner with a large pending queue:
          // replay silently through the ordinary drained submit path.
          // Different epoch (or a tail from a build that recorded none) ⇒
          // true divergence: offer the choice, never silently bulldoze.
          if (tailConnRef.current !== c) {
            tailConnRef.current = c;
            const tail = offlineTailRef.current ?? [];
            if (tail.length) {
              if (offlineEpochRef.current === c.genesisId) {
                runTailReplay(c, "plain");
              } else {
                const structural = toSuggestions(tail, "", "").dropped.length;
                // diverged=true never yields "fast-forward"; narrow for the type.
                const mode = arrivalMode(tail.length, true) === "draft" ? "draft" as const : "suggest" as const;
                setArrival({ mode, tailLength: tail.length, structural });
              }
            }
          }
        }
      },
      onPresence: (participant, pos) =>
        setPresence((prev) => ({ ...prev, [participant]: pos })),
      onRefused: (reason) => {
        // Read-only is a PER-EDIT refusal, not a connection failure (doc 14
        // §2.5): the blocked participant keeps the live view — "you can keep
        // reading" — so it must never flip the full-page refused state.
        // The optimistic local edit is dropped by the stuck-pending watchdog
        // (refused resends exhaust the retry budget → drop + heal from the
        // mirror), so the room converges without the victim's edit; typing
        // simply works again once the owner lifts the lock.
        if (reason === "read-only") {
          setReadOnlyBlocked(true);
          return;
        }
        // NOT-OWNER is a REJECTED ACTION, not a dead session, and treating it
        // as fatal was a real bug: the admin op bounced, the connection stayed
        // perfectly healthy, and the whole editor was replaced by "Connection
        // refused: not-owner. Refresh to retry." — losing the live document
        // view over a button press that should simply have done nothing.
        //
        // It is reachable without anyone being devious. `ownerToken` is a
        // string a client holds; holding one is not the same as it being
        // VALID. A room re-seeded into a new epoch mints a fresh owner token,
        // so the previous one stays truthy in this browser, keeps rendering
        // the admin controls, and fails the moment it is used.
        //
        // The refusal is also the authoritative answer to a question the
        // client cannot otherwise ask — the welcome carries no owner flag, so
        // "am I the owner" is unknowable until an admin op is tried. Recording
        // it lets the UI stop offering controls it now knows will fail.
        if (reason === "not-owner") {
          setNotOwner(true);
          return;
        }
        // The session is over, so any countdown toward its ending is now a
        // lie — including the `idle-timeout` / `session-expired` kicks that
        // ARE the deadline arriving. Cleared before the refusal so no render
        // can ever show both.
        clearDeadlines();
        setRefused(reason);
      },
      onSessionWarning: ({ reason, inMs }) => setDeadlines((d) => ({ ...d, [reason]: Date.now() + inMs })),
      // Only ever `idle` — the lifetime deadline cannot be cancelled, so
      // there is deliberately no per-reason branch here to get wrong.
      onSessionWarningCleared: () => setDeadlines((d) => (d.idle === null ? d : { ...d, idle: null })),
      onRoster: (r) => setRoster(r),
      onSelfHeal: () => setSelfHeals((n) => n + 1),
      // A submit that never left the client. Counted, never swallowed.
      onSubmitDropped: () => setDroppedPreReady((n) => n + 1),
      // Async failures with nowhere to return (seal, transport send). Surfaced
      // rather than swallowed — this callback exists because three bugs this
      // session lived inside empty catches. console.error is the floor: an app
      // embedding this should report it properly.
      onError: (info) => {
        // A submit that threw on its way out is a LOST EDIT — counted so the
        // loss accounting partitions (see CollabSession.sendFailures). Other
        // sites are reported but not counted here; they are not edit loss.
        if (info.where === "enc.submit") setSendFailures((n) => n + 1);
        // eslint-disable-next-line no-console
        console.error(`[wordinweb] ${info.where}`, info.error);
      },
      onEpochChange: (from, to) => {
        // A recorded offline tail is handled by the welcome-time resolution
        // in onChange (the arrival ladder keys on the tail's RECORDED epoch
        // vs the welcome's, which also covers a reload that lands after
        // this callback already fired in a previous session).
        // Someone re-seeded while we were away (doc 12 §5 case 2): the
        // connection already took the server's state and withheld our
        // old-epoch pending. Preserve the superseded bundle as a DRAFT
        // before the persister overwrites the main slot — this is the
        // "your offline copy is saved as a draft" guarantee.
        if (store) {
          void store.get(docId).then((old) => {
            if (old && old.genesisId === from) {
              return store.put({ ...old, docId: draftKey(docId, from) });
            }
          });
        }
        setEpochChanged({ from, to });
      },
      onFastForward: (from, _to) => {
        // Doc 15 fast-forward: nothing was lost, so no draft banner — but
        // the superseded state is BANKED recoverable-not-gone (the
        // fabricated-lineage mitigation: worst case is restorable).
        if (store) {
          void store.get(docId).then((old) => {
            if (old && old.genesisId === from) {
              return store.put({ ...old, docId: supersededKey(docId, from) });
            }
          });
        }
      },
    };
    // Mode from the LINK (doc 13 §6): a docKey selects the encrypted
    // connection; both classes expose the same session surface, so the rest
    // of the hook (and DocxView) is mode-blind. Construction is async only
    // when a share code must be stretched (PBKDF2, once per join).
    const flush = () => {
      void persister?.flush();
      // The offline writer's armed write, if any (it is armed only while
      // the persister cannot export a bundle, so at most one of these two
      // produces a write).
      flushOfflinePersist();
    };
    /**
     * This run is a REJOIN after a detected drop, not a first join — so it
     * must claim the identity the way `resume()` already does.
     *
     * The dropped socket may still be registered server-side: a half-open
     * connection is invisible to the server too, and it holds this clientId
     * under the single-live-connection rule until its own timeout. Without the
     * claim the rejoin is refused `already-open` and the session is stuck
     * behind a socket that is provably dead — the failure this whole path
     * exists to recover from. A reconnect IS that identity's continuation, by
     * definition, which is the same argument resume() makes.
     */
    const rejoining = reconnectNonce > 0;
    void (async () => {
      const stretched = docKey && shareCode ? await stretchShareCode(shareCode, docId) : undefined;
      const codeProof = stretched ? btoa(String.fromCharCode(...stretched)) : undefined;
      if (disposed) return;
      // Media relay origin (doc 16 §3). Absent ⇒ the connection has no media
      // duties at all and images stay a local-only feature, which is exactly
      // how every non-media caller behaves today.
      const mediaOpts = httpBase ? { httpBase } : undefined;
      const conn: CollabConnection = docKey
        ? (new EncryptedCollabConnection(transport, clientId, docKey, callbacks, stretched, undefined, mediaOpts) as unknown as CollabConnection)
        : new CollabConnection(transport, clientId, callbacks, mediaOpts);
      connRef.current = conn;
      if (store) {
        persister = new BundlePersister(conn, store, docId, {
          // The browser's bundle IS the durable copy in a zero-custody design,
          // so a failed write (quota, blocked storage, private mode) means
          // "saved" is a lie. Never silent.
          onError: (err) => {
            // Surfaced to the UI as well as the console: a failed write is
            // data-loss-class here, and a console line nobody is watching is
            // the same silence this arc exists to remove.
            setPersistErrors((n) => n + 1);
            // eslint-disable-next-line no-console
            console.error("[wordinweb] bundle-persist", err);
          },
          // The offline tail rides every bundle write (doc 15 §4.3 "tail
          // hygiene"): without this, the first post-resume write would
          // ERASE a stored tail the user had not reconciled yet.
          offlineTail: () => {
            const tail = offlineTailRef.current ?? [];
            return tail.length ? { tail: [...tail], epoch: offlineEpochRef.current ?? undefined } : null;
          },
        });
        persisterRef.current = persister;
        if (typeof window !== "undefined") {
          window.addEventListener("pagehide", flush);
          document.addEventListener("visibilitychange", flush);
        }
        // A RECONNECT MUST NOT OVERTAKE THE SAVE THAT PRECEDED IT. The
        // previous run's teardown started a flush it could not await; reading
        // the store before that lands would resume from a stale bundle — or,
        // on the very first save, find nothing and join COLD, discarding
        // exactly the state the resume path exists to carry across the drop.
        await flushDoneRef.current;
        if (disposed || connRef.current !== conn) return;
        // Resume if a bundle exists, else join cold. The get() is async;
        // the editor stays !ready (input disabled) until the welcome.
        const bundle = await store.get(docId);
        if (disposed || connRef.current !== conn) return;
        // The stored tail seeds the in-memory one only when memory holds
        // nothing (a fresh page). Within a page's lifetime the ref is always
        // at least as fresh as the store (the store trails by the throttle
        // window), so a reconnect must never clobber it with a stale read.
        if (!offlineTailRef.current?.length) {
          offlineTailRef.current = bundle?.offlineTail;
          offlineEpochRef.current = bundle?.offlineTailEpoch ?? null;
          setOfflineHeld(bundle?.offlineTail?.length ?? 0);
        }
        // resume() already claims takeover for this exact reason (doc 12 §7).
        if (bundle) conn.resume(bundle, token, { profile, codeProof, ownerToken });
        else conn.join(docId, token, { profile, takeover: takeover || rejoining, codeProof, ownerToken });
      } else {
        conn.join(docId, token, { profile, takeover: takeover || rejoining, codeProof });
      }
    })();
    return () => {
      disposed = true;
      monitor?.stop(); // no probe may outlive the socket it was measuring
      if (typeof document !== "undefined") document.removeEventListener("visibilitychange", onVisible);
      if (store && typeof window !== "undefined") {
        window.removeEventListener("pagehide", flush);
        document.removeEventListener("visibilitychange", flush);
      }
      // The offline writer's armed write, if any, joins the flush handoff
      // (it chains itself onto flushDoneRef; armed only while the persister
      // cannot export, so the two flushes never both produce a write).
      flushOfflinePersist();
      // Published so the NEXT run can await it before reading the store —
      // this is the handoff that makes a reconnect a resume (see above).
      flushDoneRef.current = Promise.all([flushDoneRef.current, persister?.flush() ?? Promise.resolve()]).then(() => {});
      persister?.stop(); // …then detach (no timers left behind)
      if (persisterRef.current === persister) persisterRef.current = null;
      connRef.current = null;
    };
    // Reconnect when the target document or endpoint changes, or when
    // `reconnectNonce` is bumped by a detected drop — the reconnect IS a
    // teardown-and-rebuild, which is what routes it through resume().
    // profile intentionally omitted from deps (an inline object literal would
    // reconnect every render); renames go through setProfile, not re-join.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [url, docId, clientId, token, createSocket, store, takeover, docKey, shareCode, ownerToken, httpBase, reconnectNonce]);

  /** Cancel a pending retry if the hook unmounts between attempts. */
  useEffect(() => () => {
    if (retryTimerRef.current !== null) clearTimeout(retryTimerRef.current);
    retryTimerRef.current = null;
  }, []);

  /**
   * This client's own write status from the roster, or undefined when the
   * server publishes none. Read from the roster rather than tracked
   * separately so there is exactly one source for it.
   */
  const myWrite = roster.find((r) => r.clientId === clientId)?.write;

  /**
   * THE REFUSAL / UNREACHABLE SPLIT (doc 12 §5). `refusalBlocked` is "the
   * server said no" — a dead session, an owner lock, a demotion, a viewer
   * link — and it is the only thing that makes accepting keystrokes a lie.
   * A merely unreachable server is `offlineActive`: edits apply locally and
   * are recorded to the durable tail, EXCEPT past the tail cap, where
   * recording genuinely cannot continue and the gate closes loudly.
   *
   * The refusal check stays live while offline on purpose: a viewer-role
   * link (or a lock this client last heard was on) does not become writable
   * by going through a tunnel — offline editing must not punch a hole in
   * any refusal.
   */
  const refusalBlocked =
    refused !== null || (myWrite !== undefined ? myWrite !== "allowed" : readOnlyBlocked);
  const offlineActive = !refusalBlocked && connection !== "live" && ready;
  const offlineCapped = offlineActive && offlineHeld >= (opts.offlineTailCap ?? OFFLINE_TAIL_CAP);

  return {
    doc,
    version,
    renderVersion,
    docEpoch,
    ready,
    // DocxView's editor applies every command to the live doc BEFORE emitting
    // its intent, so the connection must not optimistically re-apply it (that
    // doubled each keystroke: "Hello" rendered as "Hello" + its reversal, and
    // the corrupted offsets got everything after the first char rejected
    // server-side). Pre-applied semantics track pending + send only.
    //
    // OFFLINE, the same intent is CAPTURED instead (doc 15 §2): the edit is
    // already in the doc, so it only needs recording. Past the cap nothing
    // routes here — writesBlocked has withdrawn the editor — and anything
    // that slips to the connection anyway lands in the loud droppedPreReady
    // backstop rather than vanishing.
    submit: (intent) => {
      if (offlineActive && !offlineCapped) return captureOffline(intent, /*preApplied*/ true);
      connRef.current?.submitPreApplied(intent);
    },
    // Toolbar/API ops: canonical-apply optimistic path (see CollabSession).
    // Offline, the capture applies through the SAME canonical applyIntent
    // code before recording, so the local result stays byte-identical to
    // what every replica will derive when the tail replays.
    submitOp: (intent) => {
      if (offlineActive && !offlineCapped) return captureOffline(intent, /*preApplied*/ false);
      connRef.current?.submit(intent);
    },
    setPresence: (pos) => connRef.current?.setPresence(pos),
    allocIds: (n) => connRef.current?.allocIds(n) ?? [],
    uploadMedia: async (bytes) => (await connRef.current?.uploadMedia(bytes)) ?? null,
    // Read straight off the connection rather than mirrored into state: it is
    // set once when the welcome lands and never changes for the session, and
    // in the encrypted connection it is assigned OUTSIDE the serial rehydrate
    // chain, so it is readable before `ready` — which is what lets the picker
    // open and the pre-check run while the document is still rehydrating.
    mediaMaxBlobBytes: connRef.current?.mediaMaxBlobBytes ?? null,
    // Collaborative undo lives on the ENCRYPTED connection: it needs the
    // mirror (a local re-derivation of the canonical log) to compute the
    // inverse. A plaintext room's authority is the server, so undo there
    // needs a wire message and is not wired yet — the optional call reports
    // "unavailable" rather than pretending.
    // The ref is TYPED as the plaintext class (the encrypted one is cast in
    // at construction — a pre-existing modelling wrinkle, not this arc's), so
    // the access is structural. The TYPE is the real one now, so a change to
    // UndoOutcome breaks here rather than silently disagreeing.
    undoLast: () => (connRef.current as Partial<{ undoLast: () => UndoOutcome }> | null)?.undoLast?.() ?? "unavailable",
    presence,
    refused,
    sessionWarning,
    readOnlyBlocked,
    notOwner,
    // Derived, not stored: one place decides "writes won't land", so the gate
    // and the banner can never disagree about whether the user may type.
    //
    // THE SERVER'S ROSTER SIGNAL WINS WHEN PRESENT, and that is the fix for a
    // real user-visible bug: `readOnlyBlocked` is refusal-driven, so it can
    // only ever say "blocked" — nothing tells it a lock was LIFTED. A blocked
    // participant therefore sat in viewer mode until reload, and needed the
    // "Try editing again" button to escape. The roster status is POSITIVE: it
    // says `allowed` too, and it is re-fanned on every transition, so a lift
    // reaches this client without anyone attempting an edit to discover it.
    //
    // `undefined` means an older server that publishes no status. Fall back to
    // the refusal-driven flag rather than assuming `allowed` — a permissive
    // default would put the user's first keystrokes straight back in the
    // appear-then-vanish loop this exists to remove.
    //
    // A NON-LIVE CONNECTION DOES NOT BLOCK: unreachable is not refused (see
    // refusalBlocked above, and the field's own doc). What DOES fold in is
    // the offline tail CAP — past it an edit cannot be durably kept, which
    // is the same "cannot land" fact this predicate exists to gate, and
    // folding it here is what keeps `editable: … && !writesBlocked` the one
    // gate every editor surface already reads.
    //
    // A REFUSAL BLOCKS WRITES FIRST OF ALL: the session is dead, and note the
    // hub does NOT close the socket after a kick, so `connection` can still
    // read "live" and the roster can still hold a stale `allowed` — neither
    // may win over a refusal. This is what keeps a document rendered behind a
    // "session ended" dialog read-only through the same single predicate.
    writesBlocked: refusalBlocked || offlineCapped,
    // Offline editing state (doc 15 §2) — see the CollabSession doc. Null
    // whenever a refusal already owns the story: the two must never both
    // claim the banner.
    offline: offlineActive ? { editsHeld: offlineHeld, capped: offlineCapped } : null,
    // Left UNSET by a disconnect on purpose. These four values describe what
    // the SERVER decided about this participant, and a client that cannot
    // reach the server has not been told anything new — reporting a stale
    // `allowed` beside `writesBlocked: true` would be a straight
    // contradiction, and inventing a fifth value would put a transport fault
    // into a permissions vocabulary. `connection` is where that story lives.
    writeStatus: myWrite,
    connection,
    reconnect,
    retryWrites: () => setReadOnlyBlocked(false),
    epochChanged,
    roster,
    setProfile: (p) => connRef.current?.setProfile(p),
    activity: connRef.current?.activity ?? [],
    admin: (action) => connRef.current?.admin(action),
    selfHeals,
    droppedPreReady,
    persistErrors,
    sendFailures,
    arrival,
    reconcile: (mode) => {
      const conn = connRef.current;
      const tail = offlineTailRef.current ?? [];
      if (!conn || !tail.length) {
        offlineTailRef.current = undefined;
        offlineEpochRef.current = null;
        setOfflineHeld(0);
        setArrival(null);
        return;
      }
      setArrival(null); // the choice is made; the banner comes down now
      if (mode === "suggest") {
        // Replay the tail as tracked changes on the crowd's doc, drained
        // one-in-flight, popping + persisting per intent (see runTailReplay
        // — a kill mid-replay resumes from the store, and the remainder is
        // re-OFFERED rather than silently replayed, because the tail keeps
        // its recorded epoch until it fully drains).
        runTailReplay(conn, "suggest");
      } else {
        // Draft: the offline copy (confirmed state + this tail) was banked
        // to the draft/superseded slot when the epoch change surfaced;
        // clear the live tail so it is not offered again.
        offlineTailRef.current = undefined;
        offlineEpochRef.current = null;
        setOfflineHeld(0);
        void persisterRef.current?.flush(); // erase the tail from the live slot
      }
    },
  };
}

/**
 * A complete collaborative editor: joins a document over the collab server and
 * renders its live, reconciled state, forwarding local edits as intents. This
 * is the inbound + outbound integration composed for you — the app supplies
 * only the connection params.
 *
 * Rendering strategy: DocxView renders the replica's LIVE document object
 * directly (`collab.doc`). Broadcasts mutate that same instance in place and
 * bump `version`, which DocxView receives as `renderSignal` and turns into a
 * single in-place repaint — no per-broadcast serialize/parse round-trip. The
 * `source` bytes are computed only on a true-conflict reload (keyed on
 * `docEpoch`) as the placeholder/fallback. The visual rendering runs in the
 * browser (as all of DocxView does); the protocol/convergence/binding it rides
 * on are covered by the headless test suites.
 */
/** Toolbar groups OFF by default in the collab editor. Uploads (image/icon/
 * screenshot/3D/media/object) are excluded per the demo threat model (no
 * object upload → no stored-payload surface); the rest are commands not yet
 * routed through intents — showing them would make local-only edits that
 * silently diverge from other participants. Flip any of these on explicitly
 * via `toolbarFeatures` once wired. */
export const COLLAB_TOOLBAR_DEFAULTS: Partial<Record<ToolbarFeature, boolean>> = {
  // Images ride the doc-16 media relay (bytes out of band, address committed
  // in the intent). The rest of the upload surface stays closed per the demo
  // threat model — no object upload, no stored-payload surface.
  icon: false, screenshot: false, model3D: false, media: false, object: false,
  history: false, // toolbar undo/redo drives the LOCAL history; collaborative undo is server-side
  drawing: false, // ink strokes have no intent yet
  arrange: false, // selected-object arrange ops are not collab-anchored yet
};

export function CollabEditor(opts: UseCollabOptions & {
  editable?: boolean;
  /** Render the Word-style ribbon toolbar above the page (default true). */
  toolbar?: boolean;
  /** Ribbon mode: "simple" (Home only) or "advanced" (default). */
  toolbarMode?: ToolbarMode;
  /** Per-group overrides merged over the collab-safe defaults. */
  toolbarFeatures?: Partial<Record<ToolbarFeature, boolean>>;
  /** Observe the imperative document API (find/replace, inserts, ...). */
  onReady?: (api: DocxViewApi) => void;
  /** Observe the live CollabSession (roster, activity, epochChanged,
   * doc-for-download) — how an app shell renders chips/banners/buttons
   * around the editor without re-implementing its composition. */
  onSession?: (session: CollabSession) => void;
  /** Custom refusal UI (e.g. already-open -> "use here instead",
   * no-session -> "bring it back live"). Default: a refresh notice.
   *
   * `ctx.docVisible` says which surface the content lands on. False (the
   * pre-live refusals: already-open, room-full, version mismatch, a code
   * gate): the content IS the page, exactly as before — render a full
   * screen. True (post-live refusals: idle-timeout, session-expired): the
   * document is still in hand and stays on screen READ-ONLY behind the
   * content, which is rendered in an overlay covering it — render a
   * dialog/scrim, not a full page. Existing one-argument hosts keep their
   * old markup untouched; the second argument only ADDS the information. */
  refusedContent?: (reason: string, ctx?: { docVisible: boolean }) => ReactNode;
}): ReactNode {
  const session = useCollab(opts);
  useEffect(() => {
    opts.onSession?.(session);
    // The session object is a fresh literal each render; observing on
    // version/ready/roster/refusal changes is what consumers need.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    // writesBlocked belongs here for the same reason `refused` does: the shell
    // renders the banner and any retry affordance off it, and without it the
    // observer only learns the editor went read-only when some UNRELATED
    // change (the next version bump) happens to re-fire this effect.
    // `connection` belongs here for the same reason `writesBlocked` does: the
    // shell renders the disconnect modal off it, and without it the observer
    // would only learn the socket died when some unrelated change happened to
    // re-fire this effect — which, on a dead connection, may be never.
  }, [session.version, session.ready, session.refused, session.roster, session.epochChanged, session.sessionWarning, session.writesBlocked, session.connection]);
  const [api, setApi] = useState<DocxViewApi | null>(null);
  // Caret survival across a true-conflict reload: the docEpoch key change
  // below remounts the whole DocxView, killing the caret's node references.
  // The stable-id encoding survives (the replica reproduces the id table via
  // the sidecar), so capture it from the OUTGOING view the moment the epoch
  // bump reaches this render — the old view is still mounted here, and its
  // own doc/id table was left untouched by the reload (restoreConfirmed
  // builds a fresh document object) — then restore it into the replacement
  // when its onReady fires. Nothing is captured when the user had no caret,
  // so an idle viewer never has focus stolen.
  const pendingCaretRef = useRef<EncodedCaret | null>(null);
  const seenEpochRef = useRef(session.docEpoch);
  if (session.docEpoch !== seenEpochRef.current) {
    seenEpochRef.current = session.docEpoch;
    pendingCaretRef.current = api?.getEncodedCaret() ?? null;
  }
  // Placeholder bytes only — DocxView renders session.doc directly (below), so
  // this is re-serialized ONLY on a reload (docEpoch), never per broadcast.
  const bytes = useMemo(
    () => (session.doc ? session.doc.save() : null),
    [session.doc, session.docEpoch],
  );

  if (session.refused) {
    // A refusal that arrives AFTER a live session (idle-timeout,
    // session-expired) leaves the document in hand — keep it on screen
    // behind the refusal content instead of replacing the page with it.
    // Pre-live refusals (already-open, room-full, version mismatch, code
    // gates) have nothing loaded, and faking an empty page behind a dialog
    // would be a lie: those keep the full-screen refusal exactly as before.
    const docVisible = !!(session.doc && bytes);
    const content = opts.refusedContent
      ? opts.refusedContent(session.refused, { docVisible })
      : createElement(
          "div",
          // Legible over a rendered page in the overlay case; inert extra
          // styling on the plain full-screen path costs nothing.
          docVisible
            ? { style: { margin: 16, padding: "10px 14px", width: "fit-content", background: "rgba(255,255,255,.95)", borderRadius: 8, font: "14px system-ui", boxShadow: "0 2px 8px rgba(0,0,0,.25)" } }
            : undefined,
          `Please refresh — ${session.refused}.`,
        );
    if (!docVisible) {
      return createElement("div", { className: "dxw-collab-refused" }, content);
    }
    return createElement(
      "div",
      { className: "dxw-collab-refused-host", style: { position: "relative", height: "100%" } },
      createElement(DocxView, {
        source: bytes,
        // NO `collab` prop: a refused session is dead — no socket, nothing to
        // submit to — so this is a plain document view, not a live editor.
        // And the SAME single gate as the live path below: refusal folds into
        // writesBlocked (see CollabSession), so this is always false here —
        // an editable surface over a dead session would apply keystrokes
        // locally and lose them.
        editable: (opts.editable ?? true) && !session.writesBlocked,
        style: { height: "100%" },
      }),
      // The overlay covers the document (blocking pointer interaction with a
      // dead page); the host's content styles itself on top of it. The
      // load-bearing `dxw-collab-refused` class keeps its meaning — "the
      // element holding the refusal content" — with the overlay placement
      // carried by the additional class. DELIBERATELY NO z-index: it would
      // create a stacking context that traps a host's `position: fixed`
      // dialog/scrim at this layer, silently re-scoping the host's own
      // z-values. DOM order (after the DocxView) already paints it above the
      // document.
      createElement(
        "div",
        { className: "dxw-collab-refused dxw-collab-refused-overlay", style: { position: "absolute", inset: 0 } },
        content,
      ),
    );
  }
  if (!session.ready || !bytes || !session.doc) {
    // Surface a dead server instead of spinning forever: if the welcome hasn't
    // arrived after a grace period, say so (the socket errored or nothing is
    // listening — the demo's most common local failure is a stopped server).
    return createElement(ConnectingNotice, null);
  }

  const view = createElement(DocxView, {
    source: bytes,
    onReady: (a: DocxViewApi) => {
      setApi(a);
      opts.onReady?.(a);
      // Re-place the caret captured before a docEpoch remount (see above).
      const caret = pendingCaretRef.current;
      if (caret) {
        pendingCaretRef.current = null;
        a.setCaretFromEncoded(caret);
      }
    },
    // Render the live doc object directly; repaint in place on each version
    // bump. submit + presence + id allocator flow out; DocxView draws carets.
    collab: {
      submit: session.submit,
      submitOp: session.submitOp as (intent: { kind: string } & Record<string, unknown>) => void,
      presence: session.presence,
      allocIds: session.allocIds,
      // Image bytes go out of band over the relay (doc 16); its presence is
      // what makes the toolbar's image button real in a room.
      uploadMedia: session.uploadMedia,
      mediaMaxBlobBytes: session.mediaMaxBlobBytes,
      // Cmd+Z: reverse my last sequenced action over the wire.
      undoLast: () => { session.undoLast(); },
      doc: session.doc,
      // The repaint signal is renderVersion, NOT version: version moves on
      // every onChange (including the editor's own submit + echo, which the
      // editor already painted), and DocxView answers each renderSignal move
      // with a whole-document relayout — behind an inert container past the
      // background-layout page threshold. Riding version made every keystroke
      // in a 500-page room queue that relayout: the "collab editor is
      // unusable on a big document" report.
      renderSignal: session.renderVersion,
      // Outbound presence: the editor reports caret moves; remote tabs draw
      // this user's cursor (inbound presence above draws theirs here).
      setPresence: session.setPresence,
      // Name flags on remote carets (doc 14 §2): presence and roster share
      // the bound-clientId keyspace, so this join is exact.
      participantNames: Object.fromEntries(session.roster.map((r) => [r.clientId, r.profile.name])),
    },
    // VIEWER MODE when the server will not take this client's writes. Not a
    // banner over a live editor: that shape applied every keystroke locally
    // and lost it, which is why the read-only banner could be on screen while
    // the user's text appeared and then vanished.
    editable: (opts.editable ?? true) && !session.writesBlocked,
    // Re-key only on docEpoch (a true-conflict reload) — NOT on every change.
    // Between reloads the live doc mutates in place and the key stays stable,
    // so DocxView repaints in place instead of re-mounting (no flash/jump).
    key: session.docEpoch,
  });

  if (opts.toolbar === false) return view;
  return createElement(
    "div",
    { className: "dxw-collab-shell", style: { display: "flex", flexDirection: "column", height: "100%" } },
    createElement(DocxToolbar, {
      api,
      mode: opts.toolbarMode ?? "advanced",
      features: { ...COLLAB_TOOLBAR_DEFAULTS, ...opts.toolbarFeatures },
    }),
    createElement("div", { style: { flex: 1, minHeight: 0 } }, view),
  );
}

/** "Connecting…" that upgrades to a clear failure message after 5s — a demo
 * pointed at a stopped server previously spun forever with no signal. */
function ConnectingNotice(): ReactNode {
  const [slow, setSlow] = useState(false);
  useEffect(() => {
    const t = setTimeout(() => setSlow(true), 5000);
    return () => clearTimeout(t);
  }, []);
  // The spinner is an EMPTY element carrying a class, not inline styles: a host
  // that styles it gets a spinner, a host that does not sees nothing extra
  // (an empty span renders as nothing). The package should not ship an opinion
  // about the look, and it must not require CSS to be legible either — hence
  // the text stands on its own.
  return createElement(
    "div",
    {
      className: "dxw-collab-connecting",
      style: { padding: 16, font: "14px system-ui" },
      role: "status",
      "aria-live": "polite",
      "aria-busy": "true",
    },
    createElement("span", { className: "dxw-collab-spinner", "aria-hidden": "true", key: "s" }),
    slow
      ? "Still connecting — is the collab server running? (start it, then reload this page)"
      : "Connecting…",
  );
}

/** Structural type for the DocxView `collab` prop — mirrors CollabSession so
 * DocxView can accept an injected session via a type-only import (no runtime
 * dependency on the collab engine). */
export type { CollabSession as InjectedCollabSession };

// Browser bundle store (IndexedDB) + the seams apps need to wire persistence
// themselves. Exported from the collab entry only — the local-only `wordinweb`
// entry has no path to any of this (doc 07 tree-shaking rule).
export type { UndoOutcome } from "@wordinweb/collab/client";
export { IndexedDbBundleStore } from "./bundle-store.js";
export { InMemoryBundleStore, BundlePersister, parseBundleKey, versionKey, draftKey, supersededKey } from "@wordinweb/collab/client";
export type { StoredDocSummary, ParsedBundleKey, StoredDocKind } from "@wordinweb/collab/client";
export { mintDocKey, docKeyFromFragment, deriveEpochKeys, sealCheckpoint, stretchShareCode, bytesToB64, docHash, mediaAddressesOf } from "@wordinweb/collab/client";
export type { BundleStore, DocBundle } from "@wordinweb/collab/client";
