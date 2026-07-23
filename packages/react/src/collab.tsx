import { useEffect, useMemo, useRef, useState, createElement, type ReactNode } from "react";
import type { DocxDocument } from "@wordinweb/core";
import { DocxView, type DocxViewApi } from "./index.js";
import { DocxToolbar, type ToolbarFeature, type ToolbarMode } from "./toolbar.js";
import {
  CollabConnection,
  EncryptedCollabConnection,
  stretchShareCode,
  createWebSocketTransport,
  BundlePersister,
  type BundleStore,
  type ParticipantProfile,
  type RosterEntry,
  type Intent,
  type PresencePosition,
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
}

/**
 * Connect to a collab document and track its live state. Returns a
 * CollabSession the app injects into `<DocxView collab={session} />` (the
 * DocxView `collab` prop is typed structurally, so the react package needs no
 * runtime dependency on this module).
 */
export function useCollab(opts: UseCollabOptions): CollabSession {
  const { url, docId, clientId, token, createSocket, store, profile, takeover, docKey, shareCode } = opts;
  const connRef = useRef<CollabConnection | null>(null);
  const [doc, setDoc] = useState<DocxDocument | null>(null);
  const [version, setVersion] = useState(0);
  const [docEpoch, setDocEpoch] = useState(0);
  const [ready, setReady] = useState(false);
  const [presence, setPresence] = useState<Record<string, PresencePosition | null>>({});
  const [refused, setRefused] = useState<string | null>(null);
  const [epochChanged, setEpochChanged] = useState<{ from: string; to: string } | null>(null);
  const [roster, setRoster] = useState<RosterEntry[]>([]);

  useEffect(() => {
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
      onRefused: (reason) => setRefused(reason),
      onRoster: (r) => setRoster(r),
      onEpochChange: (from, to) => {
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
      const conn: CollabConnection = docKey
        ? (new EncryptedCollabConnection(transport, clientId, docKey, callbacks, stretched) as unknown as CollabConnection)
        : new CollabConnection(transport, clientId, callbacks);
      connRef.current = conn;
      if (store) {
        persister = new BundlePersister(conn, store, docId);
        if (typeof window !== "undefined") {
          window.addEventListener("pagehide", flush);
          document.addEventListener("visibilitychange", flush);
        }
        // Resume if a bundle exists, else join cold. The get() is async;
        // the editor stays !ready (input disabled) until the welcome.
        const bundle = await store.get(docId);
        if (disposed || connRef.current !== conn) return;
        if (bundle) conn.resume(bundle, token, { profile, codeProof });
        else conn.join(docId, token, { profile, takeover, codeProof });
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
  }, [url, docId, clientId, token, createSocket, store, takeover, docKey, shareCode]);

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
    presence,
    refused,
    epochChanged,
    roster,
    setProfile: (p) => connRef.current?.setProfile(p),
    activity: connRef.current?.activity ?? [],
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
const COLLAB_TOOLBAR_DEFAULTS: Partial<Record<ToolbarFeature, boolean>> = {
  image: false, icon: false, screenshot: false, model3D: false, media: false, object: false,
  history: false, // toolbar undo/redo drives the LOCAL history; collaborative undo is server-side
  drawing: false, // ink strokes have no intent yet
  arrange: false, // selected-object arrange ops are not collab-anchored yet
  headerFooter: false, // no header/footer intents yet
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
  }, [session.version, session.ready, session.refused, session.roster, session.epochChanged]);
  const [api, setApi] = useState<DocxViewApi | null>(null);
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
    onReady: (a: DocxViewApi) => { setApi(a); opts.onReady?.(a); },
    // Render the live doc object directly; repaint in place on each version
    // bump. submit + presence + id allocator flow out; DocxView draws carets.
    collab: {
      submit: session.submit,
      submitOp: session.submitOp as (intent: { kind: string } & Record<string, unknown>) => void,
      presence: session.presence,
      allocIds: session.allocIds,
      doc: session.doc,
      renderSignal: session.version,
      // Outbound presence: the editor reports caret moves; remote tabs draw
      // this user's cursor (inbound presence above draws theirs here).
      setPresence: session.setPresence,
      // Name flags on remote carets (doc 14 §2): presence and roster share
      // the bound-clientId keyspace, so this join is exact.
      participantNames: Object.fromEntries(session.roster.map((r) => [r.clientId, r.profile.name])),
    },
    editable: opts.editable ?? true,
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
export { IndexedDbBundleStore } from "./bundle-store.js";
export { InMemoryBundleStore, BundlePersister } from "@wordinweb/collab/client";
export { mintDocKey, docKeyFromFragment, deriveEpochKeys, sealCheckpoint, stretchShareCode, bytesToB64, docHash } from "@wordinweb/collab/client";
export type { BundleStore, DocBundle } from "@wordinweb/collab/client";
