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
  createWebSocketTransport,
  BundlePersister,
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
}

export interface CollabSession {
  /** The live document to render (null until the welcome arrives). */
  doc: DocxDocument | null;
  /** Monotonically increases on every reconciled change — a cheap re-render
   * signal (the doc object may be reloaded by reconciliation). */
  version: number;
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
   * Offline reconciliation (doc 15 arrival ladder). When a resume lands in
   * a diverged epoch and the bundle has an offline tail, the recommended
   * mode is `suggest` (replay as tracked changes). Null when there's
   * nothing to reconcile. `reconcile()` runs the chosen mode.
   */
  arrival: { mode: "suggest" | "draft"; tailLength: number } | null;
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
  const [epochChanged, setEpochChanged] = useState<{ from: string; to: string } | null>(null);
  const [roster, setRoster] = useState<RosterEntry[]>([]);
  const [arrival, setArrival] = useState<{ mode: "suggest" | "draft"; tailLength: number } | null>(null);
  const [selfHeals, setSelfHeals] = useState(0);
  const [droppedPreReady, setDroppedPreReady] = useState(0);
  const [persistErrors, setPersistErrors] = useState(0);
  const [sendFailures, setSendFailures] = useState(0);
  const offlineTailRef = useRef<import("@wordinweb/collab/client").DocBundle["offlineTail"]>(undefined);

  useEffect(() => {
    // Every dep change here means a DIFFERENT connection, so a deadline
    // announced by the previous one no longer applies to anything.
    clearDeadlines();
    const socket = (createSocket ?? ((u: string) => new WebSocket(u)))(url);
    const transport = createWebSocketTransport(socket);
    let persister: BundlePersister | null = null; // assigned below iff store
    const callbacks: import("@wordinweb/collab/client").ConnectionCallbacks = {
      onChange: () => {
        const c = connRef.current;
        if (!c) return;
        setDoc(c.doc);
        setReady(c.ready);
        setVersion((v) => v + 1); // signal a re-render on every reconciled change
        setDocEpoch(c.docEpoch); // bumps only on a reload (true conflict)
        persister?.notify(); // throttled bundle write (doc 12 §4)
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
        // Arrival ladder (doc 15): a diverged rejoin with a recorded offline
        // tail offers rebase-as-suggestions (small) or draft (large).
        const tail = offlineTailRef.current ?? [];
        if (tail.length) {
          const mode = arrivalMode(tail.length, true);
          if (mode !== "fast-forward") setArrival({ mode, tailLength: tail.length });
        }
        // Someone re-seeded while we were away (doc 12 §5 case 2): the
        // connection already took the server's state and withheld our
        // old-epoch pending. Preserve the superseded bundle as a DRAFT
        // before the persister overwrites the main slot — this is the
        // "your offline copy is saved as a draft" guarantee.
        if (store) {
          void store.get(docId).then((old) => {
            if (old && old.genesisId === from) {
              return store.put({ ...old, docId: `${docId}#draft-${from}` });
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
              return store.put({ ...old, docId: `${docId}#superseded-${from}` });
            }
          });
        }
      },
    };
    // Mode from the LINK (doc 13 §6): a docKey selects the encrypted
    // connection; both classes expose the same session surface, so the rest
    // of the hook (and DocxView) is mode-blind. Construction is async only
    // when a share code must be stretched (PBKDF2, once per join).
    let disposed = false;
    const flush = () => void persister?.flush();
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
        });
        if (typeof window !== "undefined") {
          window.addEventListener("pagehide", flush);
          document.addEventListener("visibilitychange", flush);
        }
        // Resume if a bundle exists, else join cold. The get() is async;
        // the editor stays !ready (input disabled) until the welcome.
        const bundle = await store.get(docId);
        if (disposed || connRef.current !== conn) return;
        offlineTailRef.current = bundle?.offlineTail;
        if (bundle) conn.resume(bundle, token, { profile, codeProof, ownerToken });
        else conn.join(docId, token, { profile, takeover, codeProof, ownerToken });
      } else {
        conn.join(docId, token, { profile, takeover, codeProof });
      }
    })();
    return () => {
      disposed = true;
      if (store && typeof window !== "undefined") {
        window.removeEventListener("pagehide", flush);
        document.removeEventListener("visibilitychange", flush);
      }
      void persister?.flush(); // last write on unmount…
      persister?.stop(); // …then detach (no timers left behind)
      connRef.current = null;
    };
    // Reconnect when the target document or endpoint changes.
    // profile intentionally omitted from deps (an inline object literal would
    // reconnect every render); renames go through setProfile, not re-join.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [url, docId, clientId, token, createSocket, store, takeover, docKey, shareCode, ownerToken, httpBase]);

  /**
   * This client's own write status from the roster, or undefined when the
   * server publishes none. Read from the roster rather than tracked
   * separately so there is exactly one source for it.
   */
  const myWrite = roster.find((r) => r.clientId === clientId)?.write;

  return {
    doc,
    version,
    docEpoch,
    ready,
    // DocxView's editor applies every command to the live doc BEFORE emitting
    // its intent, so the connection must not optimistically re-apply it (that
    // doubled each keystroke: "Hello" rendered as "Hello" + its reversal, and
    // the corrupted offsets got everything after the first char rejected
    // server-side). Pre-applied semantics track pending + send only.
    submit: (intent) => connRef.current?.submitPreApplied(intent),
    // Toolbar/API ops: canonical-apply optimistic path (see CollabSession).
    submitOp: (intent) => connRef.current?.submit(intent),
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
    writesBlocked: myWrite !== undefined ? myWrite !== "allowed" : readOnlyBlocked,
    writeStatus: myWrite,
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
      if (!conn || !tail.length) { setArrival(null); return; }
      if (mode === "suggest") {
        // Replay the offline tail as tracked changes on the crowd's doc
        // (drained one-in-flight — the supported discipline). The author
        // is this participant's display name (doc 14 attribution).
        const author = profile?.name || clientId;
        const date = new Date().toISOString();
        const { suggestions } = toSuggestions(tail as never, author, date);
        let i = 0;
        const step = () => {
          if (i >= suggestions.length) { offlineTailRef.current = undefined; setArrival(null); return; }
          conn.submit(suggestions[i++] as never);
          // Drain: the loopback/echo advances confirmedSeq; poll a tick.
          setTimeout(step, 30);
        };
        step();
      } else {
        // Draft: the superseded bundle is already banked by onEpochChange;
        // just dismiss (the offline edits live in the draft slot).
        offlineTailRef.current = undefined;
        setArrival(null);
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
   * no-session -> "bring it back live"). Default: a refresh notice. */
  refusedContent?: (reason: string) => ReactNode;
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
  }, [session.version, session.ready, session.refused, session.roster, session.epochChanged, session.sessionWarning, session.writesBlocked]);
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
    return opts.refusedContent
      ? createElement("div", { className: "dxw-collab-refused" }, opts.refusedContent(session.refused))
      : createElement("div", { className: "dxw-collab-refused" }, `Please refresh — ${session.refused}.`);
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
      renderSignal: session.version,
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
  return createElement(
    "div",
    { className: "dxw-collab-connecting", style: { padding: 16, font: "14px system-ui" } },
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
export { InMemoryBundleStore, BundlePersister } from "@wordinweb/collab/client";
export { mintDocKey, docKeyFromFragment, deriveEpochKeys, sealCheckpoint, stretchShareCode, bytesToB64, docHash, mediaAddressesOf } from "@wordinweb/collab/client";
export type { BundleStore, DocBundle } from "@wordinweb/collab/client";
