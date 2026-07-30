import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { CollabEditor, IndexedDbBundleStore, InMemoryBundleStore, versionKey, type BundleStore, type CollabSession, type DocBundle, type StoredDocSummary } from "wordinweb/collab";
import { type DocxViewApi } from "wordinweb";
import { reviveEncrypted } from "./e2ee-flows";
import { pruneVersions, versionByteBudget, VERSION_COUNT_CAP } from "./version-retention";
import { useStartupReclaim } from "./startup-reclaim";
import { FileMenu, fmtSize, savedDocName } from "./file-menu";
import { PerfHud } from "./perf/hud";
import { perfMonitor, type DocStats } from "./perf/metrics";

/**
 * The zero-custody demo app (plan doc 12): everything around the editor —
 * roster chips, the download escape hatch, the epoch-change draft banner,
 * and the two refusal flows the lifecycle produces:
 *
 *  - `already-open`  → "Use here instead" (doc 12 §7 single-tab takeover)
 *  - `no-session`    → "Bring it back live" (PUT this browser's bundle —
 *                       the browsers ARE the recovery machinery)
 *
 * The server keeps nothing: this page's IndexedDB bundle is the durable
 * copy, so getting a file out has to stay reachable (doc 12 §4 durability
 * honesty). It now lives one click deep under File — the same place every
 * other editor keeps it — and the two screens where the copy is actually AT
 * RISK keep their own download button in reach: the ConnectionLost dialog and
 * the lifetime countdown, both of which appear over the document.
 */

function b64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

export function App({ url, httpBase, docId, clientId, name, docKey, ownerToken, initialShareCode, onNewDocument, onDisconnect, onNameChange, onShareLink, shareCopied, store: storeProp }: {
  url: string;
  httpBase: string;
  docId: string;
  clientId: string;
  name: string;
  /** E2EE mode (doc 13): present iff the link carried `#k=` — the mode is
   * the LINK's fact; the editor hard-refuses a contradicting welcome. */
  docKey?: string;
  /** Owner capability (doc 14 §2.5): held by the doc's creator; unlocks
   * the admin controls below. Never in the shared link. */
  ownerToken?: string;
  /** Share code already known on THIS device (the creator's own code, or a
   * joiner's previously-entered one). Seeds the join so a code-gated doc
   * connects without re-prompting. Never comes from the shared link. */
  initialShareCode?: string;
  /** Leave this document and start a fresh one — the "I forgot the code"
   * escape (a code-locked session can't be entered, but a new one can). */
  onNewDocument?: () => void;
  /** Leave the session and hand the CURRENT document bytes back to the local
   * editor. Null bytes when the session never became ready — there is simply
   * nothing to carry out, and the caller reopens the blank template. */
  onDisconnect?: (bytes: Uint8Array | null) => void;
  /** Rename this participant (persisted by the caller). */
  onNameChange?: (name: string) => void;
  /** Copy the share link — the URL is never rendered, only copied. */
  onShareLink?: () => void;
  /** True briefly after a successful copy, for the button's confirmation. */
  shareCopied?: boolean;
  /** Where bundles/versions/drafts persist (shared with the local editor).
   * Defaults to this browser's IndexedDB; tests inject the in-memory store. */
  store?: BundleStore;
}) {
  const store = useMemo(() => storeProp ?? new IndexedDbBundleStore(), [storeProp]);
  const [session, setSession] = useState<CollabSession | null>(null);
  const socketRef = useRef<WebSocket | null>(null);
  /**
   * The demo builds its own socket so Disconnect can CLOSE it: the hook's
   * teardown flushes the bundle and drops its reference but never closes the
   * transport, so without this the server would keep counting a participant
   * who had already left, and the room would stay alive on a phantom.
   *
   * STABLE IDENTITY IS LOAD-BEARING — `createSocket` is in useCollab's effect
   * deps, so an inline arrow is a new function every render and the session
   * reconnects in a loop. It never stabilises, the editor never mounts, and
   * the symptom is simply "no page" (measured: 6 of 7 browser scenarios
   * failed, including every joiner, until this was hoisted).
   */
  const makeSocket = useCallback((u: string) => {
    const sock = new WebSocket(u);
    socketRef.current = sock;
    return sock;
  }, []);

  /**
   * Push a rename to the room. The `profile` prop is READ ONLY AT JOIN — the
   * hook deliberately keeps it out of its effect deps so an inline object
   * literal can't reconnect the session on every render — so without this an
   * edited name would sit in localStorage and never reach anyone else's
   * roster. Debounced because the control is a text field: one frame per
   * keystroke would be one roster fan-out per keystroke.
   */
  const joinedName = useRef(name);
  useEffect(() => {
    if (!session?.ready || name === joinedName.current) return;
    const t = setTimeout(() => {
      joinedName.current = name;
      session.setProfile({ name, color: "" }); // server assigns the palette color
    }, 400);
    return () => clearTimeout(t);
  }, [name, session]);

  /** Leave the room, keep the copy: snapshot the live document, close the
   * socket, and hand the bytes up. Nothing durable is lost — the server was
   * never holding it. */
  const leaveSession = () => {
    const bytes = session?.doc ? session.doc.save() : null;
    try { socketRef.current?.close(); } catch { /* already closed */ }
    socketRef.current = null;
    onDisconnect?.(bytes);
  };
  // E2E test hook (dev harness only — anon-share is never published): expose
  // the live CollabSession so browser tests can inject intents through the
  // REAL connection and read the converged doc, independent of the editor UI.
  useEffect(() => {
    const w = window as unknown as { __ww?: unknown };
    w.__ww = session
      ? {
          _session: session, // raw session for E2E deep inspection (dev only)
          submitOp: (i: unknown) => session.submitOp(i as never),
          admin: (act: unknown) => session.admin(act as never),
          allocIds: (n: number) => session.allocIds(n),
          ready: session.ready,
          roster: () => session.roster,
          // Base64 of the canonical docx — the byte-identical convergence
          // check across clients (save() is deterministic by design).
          saveB64: () => {
            if (!session.doc) return null;
            const bytes = session.doc.save();
            let bin = "";
            for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
            return btoa(bin);
          },
          // Live perf snapshot (null unless the perf HUD is open) — lets the
          // stress suite log the same numbers the HUD shows.
          perf: () => perfMonitor.snapshot(),
          // Times the connection self-healed a drifted optimistic replica
          // (B6a catch-and-resync) — the swarm harness reads this to tell
          // "drifted then healed" from "still drifted".
          selfHeals: () => session.selfHeals,
          // Submits the connection dropped because it was not ready (B13).
          // Should always read 0 — the editor is gated on `ready`, so a
          // non-zero value means edits are being lost before the wire.
          droppedPreReady: () => session.droppedPreReady,
          // Submits lost when their seal/send threw (B13 family). With
          // droppedPreReady this partitions edit loss: refused-before-applied
          // vs applied-then-lost-on-the-way-out.
          sendFailures: () => session.sendFailures,
          // Concatenated text (readable diffs on failure).
          text: () => {
            if (!session.doc) return "";
            const walk = (el: { name: string; text: string; children: unknown[] }): string =>
              (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
            return walk(session.doc.docRoot as never);
          },
        }
      : null;
  }, [session]);
  // Document facts for the perf HUD: a paragraph COUNT and the serialized
  // byte length — never text. Called at most every 2s, on idle (see metrics.ts).
  const docStats = useCallback((): DocStats | null => {
    const doc = session?.doc;
    if (!doc) return null;
    let paragraphs = 0;
    const count = (el: { name: string; children: unknown[] }): void => {
      if (el.name === "w:p") paragraphs++;
      for (const c of el.children as { name: string; children: unknown[] }[]) count(c);
    };
    count(doc.docRoot as never);
    return { paragraphs, bytes: doc.save().byteLength };
  }, [session]);
  const [takeover, setTakeover] = useState(false);
  const [reviveState, setReviveState] = useState<"idle" | "reviving" | "no-copy">("idle");
  // Share code (doc 13 §7): seeded from this device's known code (creator's own,
  // or a previously-entered one) so a code-gated join doesn't re-prompt; also
  // entered on a code-required/invalid refusal. Remounting with it retries (and
  // in E2EE mode mixes into key derivation).
  const [shareCode, setShareCode] = useState<string | undefined>(() => initialShareCode);
  const [codeDraft, setCodeDraft] = useState(initialShareCode ?? "");
  const [readOnly, setReadOnly] = useState(false);
  // HOLDING a token is not the same as it being VALID. A room re-seeded into
  // a new epoch mints a fresh owner token, leaving the old one truthy here
  // with no rights attached — which used to render the admin controls and
  // then blow the whole editor away on the first click (a `not-owner`
  // refusal was treated as a dead session). The server's refusal is the
  // authoritative answer, so once it arrives the controls stop being offered.
  const isOwner = !!ownerToken && !session?.notOwner;
  // Remount key: bumps to retry the connection after takeover/revival.
  const [attempt, setAttempt] = useState(0);
  // "Keep editing offline" was chosen on the ConnectionLost dialog. Reset
  // the moment the connection recovers, so a LATER loss re-raises the
  // dialog rather than inheriting a dismissal from a different outage.
  const [lostDismissed, setLostDismissed] = useState(false);
  useEffect(() => {
    if (session?.connection !== "lost") setLostDismissed(false);
  }, [session?.connection]);
  const [showActivity, setShowActivity] = useState(false);
  // Versions (doc 14 §1): frozen restore points beside the live bundle.
  // Stored in the same IndexedDB db under version-suffixed keys — a
  // DocVersion is bundle-shaped enough for the demo (docx + sidecar).
  const [versions, setVersions] = useState<{ key: string; label: string; savedAt: number }[]>([]);
  /** A version save/read failed — surfaced in the version strip, because a
   * swallowed quota error means the user believes a restore point exists
   * when it does not (the browser is the only place it could exist). */
  const [versionError, setVersionError] = useState<string | null>(null);
  /** Older versions were deleted by retention (see version-retention.ts) —
   * always said on screen. Silently deleting a restore point is worse than
   * refusing to make a new one. */
  const [versionNotice, setVersionNotice] = useState<string | null>(null);
  /** Startup reclaim (startup-reclaim.ts): copies leaked before retention
   * existed are removed on load, and what was removed is reported here. */
  const reclaim = useStartupReclaim(store);
  /** What the store holds, measured when a persist write fails — the banner
   * must say what is taking the space, not just that space ran out. */
  const [storedSummary, setStoredSummary] = useState<{ count: number; bytes: number } | null>(null);
  /** Bumped by the banner's "Manage saved copies" to open the File menu. */
  const [savedMenuRequest, setSavedMenuRequest] = useState(0);
  const persistFailing = (session?.persistErrors ?? 0) > 0;
  useEffect(() => {
    if (!persistFailing) return;
    let alive = true;
    store.list().then(
      (all) => {
        if (alive) setStoredSummary({ count: all.length, bytes: all.reduce((n, s) => n + s.byteLength, 0) });
      },
      () => {
        // Unlistable storage: the banner still shows, just without numbers.
        if (alive) setStoredSummary(null);
      },
    );
    return () => { alive = false; };
  }, [persistFailing, store]);

  // Seed the strip FROM THE STORE, not from empty component state: saved
  // versions are durable, so they must still be on screen after a reload.
  useEffect(() => {
    let alive = true;
    store.list().then(
      (all) => {
        if (!alive) return;
        setVersions(
          all
            .filter((s) => s.kind === "version" && s.docId === docId)
            .map((s) => ({ key: s.key, label: s.label ?? "(auto)", savedAt: s.versionSavedAt ?? s.savedAt }))
            .sort((a, b) => a.savedAt - b.savedAt),
        );
      },
      () => {
        if (alive) setVersionError("Saved versions couldn’t be read from this browser’s storage.");
      },
    );
    return () => { alive = false; };
  }, [store, docId]);

  // `window.prompt` blocks the whole page — in a live session that stalls the
  // socket pump behind a browser dialog we do not control, and it is the one
  // piece of chrome here that cannot be styled. The modal below replaces it.
  const [versionModal, setVersionModal] = useState(false);
  const [versionLabel, setVersionLabel] = useState("");
  const versionRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (!versionModal) return;
    versionRef.current?.focus();
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setVersionModal(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [versionModal]);

  /** What a dropped version is called in the retention notice. */
  const versionName = (s: StoredDocSummary) =>
    s.label ?? new Date(s.versionSavedAt ?? s.savedAt).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });

  const saveVersion = async (label: string) => {
    if (!session?.doc) return;
    const savedAt = Date.now();
    // Everything retention deleted in this save, oldest first — reported on
    // screen whether or not the save itself succeeded (see the catch).
    let dropped: StoredDocSummary[] = [];
    try {
      const bundle = await store.get(docId);
      if (!bundle) {
        // The persister has not landed a bundle — there is nothing durable
        // to freeze, and pretending otherwise is the failure mode this
        // feature exists to avoid.
        setVersionError("Nothing is saved in this browser yet to freeze — storage may be blocked. Download a copy instead.");
        return;
      }
      const key = versionKey(docId, savedAt, label);
      const next = { ...bundle, docId: key, savedAt };
      const budget = await versionByteBudget();
      try {
        await store.put(next);
        // Landed. Retention runs AFTER the save, so a restore point is only
        // ever deleted once the newer one it makes way for durably exists.
        dropped = await pruneVersions(store, docId, budget);
      } catch (err) {
        // Out of space (or blocked). Free THIS doc's older versions — the one
        // kind of stored copy that is redundant by construction — and retry
        // once. keepNewest: 0 because the version being made room for is the
        // incoming one; the budget leaves it headroom.
        dropped = await pruneVersions(store, docId, Math.max(0, budget - bundle.confirmedBytes.byteLength), { keepNewest: 0, countCap: VERSION_COUNT_CAP - 1 });
        if (dropped.length === 0) throw err; // nothing to free — surface the original failure
        await store.put(next);
      }
      setVersionError(null);
      const droppedKeys = new Set(dropped.map((d) => d.key));
      setVersions((v) => [...v.filter((x) => !droppedKeys.has(x.key)), { key, label: label || "(auto)", savedAt }]);
    } catch {
      setVersionError("Couldn’t save this version — storage is full or blocked. Download a copy instead.");
    } finally {
      if (dropped.length) {
        const names = dropped.map(versionName);
        const shown = names.length > 4 ? `${names.slice(0, 3).join(", ")} and ${names.length - 3} more` : names.join(", ");
        setVersionNotice(
          `Removed ${dropped.length} older version${dropped.length === 1 ? "" : "s"} to stay within this browser’s storage: ${shown}. Download a version to keep it outside the browser.`,
        );
      }
    }
  };
  const downloadVersion = async (key: string, label: string, savedAt: number) => {
    try {
      const v = await store.get(key);
      if (!v) throw new Error("gone");
      const blob = new Blob([v.confirmedBytes as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      // Restore-as-DRAFT (doc 14 §1): a version opens as a new file, never
      // mutates the live doc — "back to then", not "merge then into now".
      a.download = `${label === "(auto)" ? "version" : label}-${savedAt}.docx`;
      a.click();
      URL.revokeObjectURL(a.href);
    } catch {
      setVersionError("Couldn’t read that version from this browser’s storage.");
    }
  };

  /** Download any saved entry from the File menu without opening it — the
   * one honest way to get at a stored copy while the session is live. */
  const downloadSaved = async (s: StoredDocSummary) => {
    try {
      const b = await store.get(s.key);
      if (!b) throw new Error("gone");
      const blob = new Blob([b.confirmedBytes as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${savedDocName(s)}.docx`;
      a.click();
      URL.revokeObjectURL(a.href);
    } catch {
      setVersionError("Couldn’t read that saved copy from this browser’s storage.");
    }
  };

  const revive = async () => {
    setReviveState("reviving");
    const bundle: DocBundle | null = await store.get(docId);
    if (!bundle) {
      // A viewer who cleared storage (or never had the doc) can't revive —
      // stated plainly (doc 12 §4: "this document lives in your browser").
      setReviveState("no-copy");
      return;
    }
    if (docKey) {
      // Encrypted revival: seal the bundle under a fresh epoch client-side
      // (the server can't — no keys). 409 = someone else won; join theirs.
      await reviveEncrypted(httpBase, bundle, docKey, shareCode);
    } else {
      await fetch(`${httpBase}/docs/${docId}`, {
        method: "PUT",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ docx: b64(bundle.confirmedBytes), sidecar: bundle.confirmedSidecar, lineage: bundle.lineage }),
      });
    }
    // Whether we won or someone else did (409), a session now exists — join it.
    setReviveState("idle");
    setAttempt((a) => a + 1);
  };

  /**
   * Editing vs Suggesting, live (doc 05). Suggesting mode already replicates:
   * typing emits `insertText` carrying `suggest{author,date}` (or a
   * `suggestRevision`), and peers apply it as a real tracked change. All that
   * was missing is the client-side control, so this is UI over an existing
   * intent surface — no new intent kind, no changed apply semantics, so no
   * ENGINE_VERSION bump (that fence guards canonical divergence).
   *
   * THE MODE IS THE EDITOR'S, NOT REACT'S. Every render reads it back with
   * `api.isSuggesting()`. A mirrored React boolean would disagree with what
   * typing actually does the moment anything else calls setSuggesting — and
   * `writesBlocked` really does drop the editor and reset the mode, which a
   * mirror would hide. `modeTick` exists only to ask for a repaint after a
   * click; it is never the source of truth.
   */
  const [api, setApi] = useState<DocxViewApi | null>(null);
  const [, setModeTick] = useState(0);
  // writesBlocked is THE write gate in this codebase (it is what turns the
  // editor read-only). View-only folds INTO it rather than sitting beside it.
  const writesBlocked = !!session?.writesBlocked;
  const suggesting = !writesBlocked && (api?.isSuggesting() ?? false);
  const mode = writesBlocked ? "view" : suggesting ? "suggest" : "edit";
  // A full tracked-change walk of the document, re-run whenever this shell
  // re-renders (i.e. per broadcast). Cheap next to the layout each broadcast
  // already triggers, and there is no incremental counter to read instead.
  const pendingSuggestions = api?.revisionCount() ?? 0;
  const toggleSuggesting = () => {
    if (!api || writesBlocked) return;
    // The participant alias is the revision author, so peers see WHO
    // suggested rather than a generic "You" on everybody's screen.
    api.setSuggesting(!api.isSuggesting(), name);
    setModeTick((n) => n + 1);
  };

  const download = () => {
    if (!session?.doc) return;
    const blob = new Blob([session.doc.save() as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${docId.slice(0, 8)}.docx`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const refusedContent = (reason: string): ReactNode => {
    if (reason === "already-open") {
      return (
        <div className="gate">
          <div className="gate-card">
          <h2>Already open in another tab</h2>
          <p>A document can only be live in one tab at a time. Taking over here disconnects the other one.</p>
          <button onClick={() => { setTakeover(true); setAttempt((a) => a + 1); }}>Use here instead</button>
          </div>
        </div>
      );
    }
    if (reason === "no-session") {
      return (
        <div>
          <p>This session has ended — the server keeps nothing between sessions.</p>
          {reviveState === "no-copy" ? (
            <p>This browser has no saved copy of the document, so it can’t revive it. Any participant who edited here before can.</p>
          ) : (
            <button disabled={reviveState === "reviving"} onClick={() => void revive()}>
              {reviveState === "reviving" ? "Bringing it back…" : "Bring it back live"}
            </button>
          )}
        </div>
      );
    }
    if (reason === "code-required" || reason === "code-invalid" || reason === "code-locked") {
      return (
        <div className="gate">
          <div className="gate-card">
            <h2>
              {reason === "code-locked" ? "Too many tries" : reason === "code-invalid" ? "That code didn\u2019t match" : "Share code required"}
            </h2>
            <p>
              {reason === "code-locked"
                ? "Wait a minute, then re-enter the share code."
                : reason === "code-invalid"
                  ? "Ask whoever shared the document \u2014 the code is sent separately from the link."
                  : "This document is locked with a code, sent separately from the link."}
            </p>
            <div className="gate-row">
          <input
            data-testid="join-share-code"
            value={codeDraft}
            onChange={(e) => setCodeDraft(e.target.value)}
            // Mirrors the field that CREATED the code (local-editor's modal):
            // codes are free-form text, not digits. `inputMode="numeric"` used
            // to live here and popped a NUMBER PAD on phones, so a passphrase
            // code was literally untypeable on the device most likely to be
            // reading a link someone sent them.
            placeholder="e.g. redwood"
            maxLength={64}
          />{" "}
          <button
            data-testid="join-submit"
            disabled={!codeDraft || reason === "code-locked"}
            onClick={() => {
              // Remember it on this device so a reload doesn't re-prompt (a
              // wrong code just gets refused again and overwritten next try).
              try { localStorage.setItem(`wordinweb-code-${docId}`, codeDraft); } catch { /* private mode */ }
              setShareCode(codeDraft);
              setAttempt((a) => a + 1);
            }}
          >
            Join
          </button>
            </div>
            {onNewDocument && (
              <p className="gate-foot">
                <span>Forgot the code?</span>
                <button className="ghost" data-testid="refused-new-document" onClick={onNewDocument}>Start a new document</button>
              </p>
            )}
          </div>
        </div>
      );
    }
    if (reason === "idle-timeout" || reason === "session-expired") {
      // THE DEADLINE ARRIVING, not a refusal. The countdown that preceded it
      // explained what was about to happen, so this must read as the ending
      // it announced — "refresh to retry" would be both wrong (there is
      // nothing to retry) and a lie about where the document lives.
      //
      // A MODAL over the (read-only) document rather than a replacement
      // screen: the document is still right there, and the dialog forces the
      // one decision that matters — where does editing continue?
      return (
        <SessionEndedModal
          reason={reason}
          reviveState={reviveState}
          onRevive={() => void revive()}
          onKeepLocal={onDisconnect ? leaveSession : undefined}
          onNewDocument={onNewDocument}
        />
      );
    }
    if (reason === "room-full") {
      return <p>This document already has the maximum number of participants (10). Ask someone to leave, then refresh.</p>;
    }
    if (reason === "engine-version-mismatch" || reason === "engine-version-required") {
      return (
        <div>
          <p>This tab is running a different version of the editor than the session — joining would silently corrupt the document.</p>
          <button data-testid="refused-reload" onClick={() => window.location.reload()}>Reload this tab</button>
        </div>
      );
    }
    return (
      <div className="gate">
        <div className="gate-card">
          <h2>Can\u2019t open this document</h2>
          {/* The reason is protocol vocabulary, not prose — shown because a
              specific word someone can search or quote beats "something went
              wrong", and these are the cases with no tailored screen yet. */}
          <p>The server refused the connection: <code>{reason}</code>.</p>
          <button onClick={() => window.location.reload()}>Reload this tab</button>
        </div>
      </div>
    );
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      {/* `sessionbar` carries the button styling. Without it these controls
          sit outside <header> and fall back to raw browser buttons, next to a
          header whose buttons are styled — the inconsistency reads as broken
          rather than plain. */}
      <div data-testid="toolbar" className="sessionbar" style={{ display: "flex", gap: 8, alignItems: "center", padding: "7px 14px", flexWrap: "wrap" }}>
        {/* Open is disabled in a session. Loading a file replaces the whole
            document locally without emitting an intent that describes the
            replacement, so this browser would fork from every peer while both
            sides still believed they were in the same room — the same failure
            as the three insert APIs that mutated locally under collab without
            emitting (fixed in 1087298).

            No Print: this screen holds a `CollabSession`, whose `doc` is a
            DocxDocument (save/docRoot) and not the view api that owns
            `print()`. An item that cannot work is worse than no item.

            Saved documents are LISTED here but never OPENABLE: opening a
            stored copy replaces the whole document without an intent, which
            forks the room exactly like Open would. Download-only is the
            honest offer while the session is live. The current doc's live
            bundle is excluded — it is the document on screen. */}
        <FileMenu
          onNew={onNewDocument}
          onDownload={download}
          disabled={!session?.ready}
          openDisabled
          listSaved={() => store.list().then((all) => all.filter((s) => s.key !== docId))}
          savedOpenDisabled
          onDownloadSaved={(s) => void downloadSaved(s)}
          onDeleteSaved={(s) => store.delete(s.key)}
          openRequest={savedMenuRequest}
        />
        <span className="bar-sep" aria-hidden="true" />
        <span style={{ flex: 1 }} />
        {session?.arrival ? (
          <span data-testid="arrival-banner" style={{ fontSize: 12, background: "#d1ecf1", padding: "2px 8px", borderRadius: 6 }}>
            You edited offline ({session.arrival.tailLength} change{session.arrival.tailLength === 1 ? "" : "s"}).{" "}
            {session.arrival.mode === "suggest" ? (
              <>
                {/* Structural edits (splits, tables, formatting) have no
                    tracked-change form — they stay in the saved draft copy
                    rather than replaying. Said up front, or their absence
                    after "add as suggestions" reads as silent loss. */}
                {session.arrival.structural > 0 && (
                  <span data-testid="arrival-structural">
                    {session.arrival.structural} of them {session.arrival.structural === 1 ? "is" : "are"} structural and will stay in your saved draft copy.{" "}
                  </span>
                )}
                <button onClick={() => session.reconcile("suggest")}>Add my changes as suggestions</button>{" "}
                <button onClick={() => session.reconcile("draft")}>Keep as draft</button>
              </>
            ) : (
              <>
                Too many to suggest cleanly — <button onClick={() => session.reconcile("draft")}>keep as a draft</button>.
              </>
            )}
          </span>
        ) : session?.storeSlow ? (
          // A DIFFERENT CLAIM from "storage is full", and worth its own words:
          // the read timed out, so this session joined without whatever was
          // saved here. Nothing has failed to write, and the document on
          // screen is fine — but a resume did not happen, and saying "full"
          // sent someone hunting through a store holding a few megabytes.
          <span data-testid="store-slow-banner" style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            This browser’s storage didn’t respond, so this document opened without
            anything previously saved here. Editing works normally.
          </span>
        ) : (session?.persistErrors ?? 0) > 0 ? (
          // DATA-LOSS WARNING, deliberately ahead of the other banners. This
          // browser's stored bundle IS the durable copy (doc 12 §4) — the
          // server keeps nothing — so a failed write means the document may
          // exist nowhere but this tab's memory. Quota exceeded, private
          // mode, or storage blocked all land here.
          //
          // The button is INLINE rather than a pointer to File > Download,
          // and it is the one place in this row that keeps its own. Moving
          // Download into the menu is right everywhere else — but this banner
          // fires exactly when the durable copy may not exist, and telling
          // someone to go find a menu item is the wrong instruction to give at
          // the moment their only copy is at risk. Same reasoning as the
          // ConnectionLost and countdown screens, which also kept theirs.
          <span data-testid="persist-banner" style={{ fontSize: 12, background: "#f8d7da", padding: "2px 8px", borderRadius: 6 }}>
            This document may not be saved in this browser — download a copy.
            {/* WHAT is taking the space, and the route to act on it. Someone
                stuck in this state needs an action, not a diagnosis: the
                button opens File > Saved, where delete lives (two-step). */}
            {storedSummary && storedSummary.count > 0 && (
              <span data-testid="persist-stored-summary">
                {" "}{storedSummary.count} saved cop{storedSummary.count === 1 ? "y uses" : "ies use"} {fmtSize(storedSummary.bytes)} in
                this browser — deleting old ones can free space.
              </span>
            )}{" "}
            <button data-testid="persist-download" onClick={download} disabled={!session?.ready}>
              Download .docx
            </button>{" "}
            <button data-testid="persist-manage-saved" onClick={() => setSavedMenuRequest((n) => n + 1)}>
              Manage saved copies
            </button>
          </span>
        ) : session?.offline?.capped ? (
          // THE TAIL CAP (doc 15 §4.3): recording stopped and the editor is
          // gated — loudly, with the way out, never by silently dropping
          // keystrokes. Ahead of the generic offline chip because this is
          // the moment the promise "your changes are being kept" ends.
          <span data-testid="offline-cap-banner" style={{ fontSize: 12, background: "#f8d7da", padding: "2px 8px", borderRadius: 6 }}>
            Offline change limit reached — editing is paused so nothing goes missing.
            Your {session.offline.editsHeld} saved changes are safe in this browser;
            reconnect to send them, or download a copy to keep working.{" "}
            <button data-testid="offline-cap-download" onClick={download} disabled={!session?.ready}>
              Download .docx
            </button>
          </span>
        ) : session?.offline ? (
          // OFFLINE, EDITING CONTINUES (doc 12 §5: offline is first-class).
          // Not "editing is paused" — that copy became false the day the
          // offline tail became real. The chip says what is actually
          // happening: edits apply here, are saved here, and replay on
          // reconnect (as suggestions if the document moved on — doc 15).
          //
          // `reconnecting` is deliberately a quiet chip rather than a modal:
          // most drops (a tunnel, a lid closed for ten seconds) heal in under
          // a second, and a dialog that flashes on every blip trains people to
          // dismiss the one that matters. `lost` gets the modal below.
          <span data-testid="reconnect-banner" data-state={session.connection}
            style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            {session.connection === "lost" ? "Connection lost" : "Reconnecting…"} — you can keep
            editing; your changes are saved in this browser
            {session.offline.editsHeld > 0
              ? ` (${session.offline.editsHeld} so far) `
              : " "}
            and will be added back when you reconnect — as suggestions to review if others edited meanwhile.
          </span>
        ) : session && session.connection !== "live" ? (
          // Non-live with offline editing UNAVAILABLE (never welcomed, or a
          // refusal-blocked participant — a viewer's link stays view-only in
          // a tunnel too). Nothing is being recorded, so say the quiet thing.
          <span data-testid="reconnect-banner" data-state={session.connection}
            style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            Reconnecting…
          </span>
        ) : session?.writesBlocked ? (
          // Owner lock (doc 14 §2.5): NOT a dead session — the doc stays live
          // and readable; editing returns when the owner lifts.
          //
          // Keyed to `writesBlocked`, the SAME predicate that gates the editor,
          // rather than to `readOnlyBlocked`. That is now load-bearing: the
          // gate engages from the roster signal BEFORE any edit is attempted,
          // so no refusal ever arrives and the refusal-driven flag stays false.
          // Keying the banner to it would leave the editor correctly frozen
          // with nothing on screen explaining why — the worst of both.
          <span data-testid="readonly-banner" data-reason={session.writeStatus ?? "unknown"}
            style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            {/* The three causes need different words because they imply
                different things to WAIT for. Telling a view-only link holder
                that the owner paused editing tells them to wait for something
                that will never happen. Where the server publishes no status
                (older build) the copy stays deliberately vague rather than
                naming a cause we were not told. */}
            {session.writeStatus === "owner-lock"
              ? "The owner paused editing for everyone — you can read along, and editing returns by itself when they lift it."
              : session.writeStatus === "demoted"
                ? "The owner set you to view-only in this session — you can read along."
                : session.writeStatus === "viewer-role"
                  ? "Your link is view-only. You can read and download this document, but editing isn’t part of this link."
                  : "This document is read-only for you right now."}
            {" "}The editor is in view mode, so nothing you type can be silently lost.{" "}
            {/* Only offered where trying could actually change the answer. A
                viewer-role link is a property of the link itself — no amount of
                retrying alters it, and a button that cannot work is the same
                false promise as the copy above. Against an older server (no
                status) this is the ONLY escape, because the block is sticky and
                no lift is announced. */}
            {session.writeStatus !== "viewer-role" && (
              <button data-testid="retry-writes" onClick={() => session?.retryWrites()}>
                Try editing again
              </button>
            )}
          </span>
        ) : session?.epochChanged ? (
          <span style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            Restored by another participant — your offline copy is saved as a draft.
          </span>
        ) : null}
        {onDisconnect && (
          <button
            data-testid="disconnect"
            title="Leave this session and keep editing your copy locally"
            onClick={leaveSession}
          >
            Disconnect
          </button>
        )}
        <button onClick={() => { setVersionLabel(""); setVersionModal(true); }} disabled={!session?.ready}>Save version</button>
        <button onClick={() => setShowActivity((v) => !v)}>{showActivity ? "Hide" : "Show"} activity</button>
        {isOwner && (
          <button
            data-testid="readonly-toggle"
            className="owner"
            aria-pressed={readOnly}
            title="Owner: block all editors (you keep writing)"
            onClick={() => { const on = !readOnly; setReadOnly(on); session?.admin({ op: "readOnly", on }); }}
          >
            {readOnly ? "🔒 Read-only ON" : "🔓 Make read-only"}
          </button>
        )}
        {/* THREE honest states, because view-only is already real. Editing and
            Suggesting are the two a person can pick; View only is what
            writesBlocked forces, and the control says so instead of pretending
            to offer a choice it cannot honour. The ON tint is the markup
            renderer's revision ink, so the chip matches what suggestions look
            like on the page. */}
        <button
          data-testid="suggest-toggle"
          data-mode={mode}
          aria-pressed={suggesting}
          disabled={writesBlocked || !api}
          title={
            writesBlocked
              ? "This session is view-only, so there is nothing to suggest into."
              : suggesting
                ? "You are suggesting. Your edits arrive as tracked changes for others to review."
                : "Switch to suggesting: your edits arrive as tracked changes instead of changing the text."
          }
          style={suggesting ? { background: "#fce8e6", borderColor: "#d93025", color: "#a50e0e" } : undefined}
          onClick={toggleSuggesting}
        >
          {mode === "view" ? "View only" : suggesting ? "✎ Suggesting" : "Editing"}
        </button>
        {pendingSuggestions > 0 && (
          <span data-testid="suggestion-count" style={{ fontSize: 12, color: "#a50e0e" }}>
            {pendingSuggestions} pending suggestion{pendingSuggestions === 1 ? "" : "s"}
          </span>
        )}
        {suggesting && (
          // KNOWN TRAP, stated where it bites. Paste in suggesting mode inside
          // a session is a complete no-op (core editor.ts, the paste collab
          // gate): Ctrl+V does nothing at all. A tooltip would only reach
          // people who hover, and silently swallowing a paste is the same
          // family as silent data loss — so it is on screen the whole time
          // the mode is on.
          <span data-testid="suggest-paste-warning" style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            Paste is unavailable in this mode.
          </span>
        )}
        <span className="bar-sep" aria-hidden="true" />
        <input
          className="alias"
          data-testid="alias"
          value={name}
          placeholder="Your name"
          maxLength={40}
          title="The name other participants see. Anyone can pick any name — it is a label, not an account."
          onChange={(e) => onNameChange?.(e.target.value)}
        />
        {onShareLink && (
          <button className="primary" data-testid="share-collab" onClick={onShareLink} title="Copy the share link to your clipboard">
            {shareCopied ? "Copied!" : "Share Collab"}
          </button>
        )}
        {onNewDocument && (
          <button data-testid="new-document" onClick={onNewDocument} title="Leave this document and start a fresh one">
            New document
          </button>
        )}
        {/* Roster last and hard right. It is the only part of this bar whose
            width is not ours to choose — a busy room can hold ten names of any
            length — so it scrolls WITHIN itself rather than wrapping the row
            or pushing the controls off the edge. */}
        <span className="roster-scroll">
          {(session?.roster ?? []).map((r) => (
            <span key={r.clientId} data-testid="roster-chip" data-connected={r.connected} style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
              <span
                title="Names are chosen by participants"
                style={{
                  padding: "2px 10px", borderRadius: 999, fontSize: 12, color: "#fff",
                  background: r.profile.color, opacity: r.connected ? 1 : 0.35, whiteSpace: "nowrap",
                }}
              >
                {r.profile.name}
              </span>
              {isOwner && r.clientId !== clientId && (
                <>
                  <button className="chip" title="Demote to viewer"
                    onClick={() => session?.admin({ op: "setRole", clientId: r.clientId, role: "viewer" })}>👁</button>
                  <button className="chip" title="Kick"
                    onClick={() => session?.admin({ op: "kick", clientId: r.clientId })}>✕</button>
                </>
              )}
            </span>
          ))}
        </span>
      </div>
      {reclaim.notice && (
        // Startup reclaim's report — same contract as the retention notice
        // below: stored copies were deleted (or could not be), so it is said
        // on screen and stays until dismissed.
        <div className="versionbar" data-testid="reclaim-bar">
          <span data-testid="reclaim-notice" style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            {reclaim.notice}{" "}
            <button className="ghost" data-testid="reclaim-dismiss" onClick={reclaim.dismiss}>Dismiss</button>
          </span>
        </div>
      )}
      {(versions.length > 0 || versionError || versionNotice) && (
        <div className="versionbar" data-testid="versions">
          <span className="versionbar-label">Saved versions</span>
          {versions.map((v) => (
            <button key={v.key} className="version-chip"
              title="Downloads as a new file — never merges into the live document"
              onClick={() => void downloadVersion(v.key, v.label, v.savedAt)}>
              {/* "(auto)" is the STORAGE sentinel for an unnamed version, so
                  it is renamed for display here rather than at the source. */}
              <span className="version-name">{v.label === "(auto)" ? "Unnamed" : v.label}</span>
              <time className="version-time">{new Date(v.savedAt).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}</time>
            </button>
          ))}
          {versionError && (
            <span data-testid="version-error" style={{ fontSize: 12, background: "#f8d7da", padding: "2px 8px", borderRadius: 6 }}>
              {versionError}
            </span>
          )}
          {versionNotice && (
            <span data-testid="version-retention-notice" style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
              {versionNotice}{" "}
              <button className="ghost" onClick={() => setVersionNotice(null)}>Dismiss</button>
            </span>
          )}
        </div>
      )}
      {versionModal && (
        <div
          className="modal-scrim"
          data-testid="version-modal"
          onMouseDown={(e) => { if (e.target === e.currentTarget) setVersionModal(false); }}
        >
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="version-modal-title">
            <h2 id="version-modal-title">Save a version</h2>
            <p>
              A frozen copy of the document as it stands now, kept in this browser beside the
              live one. Downloading a version never merges it back — it opens as its own file.
            </p>
            <label htmlFor="version-label-input">Name (optional)</label>
            <input
              id="version-label-input"
              ref={versionRef}
              data-testid="version-label"
              value={versionLabel}
              placeholder="e.g. before the edits"
              maxLength={64}
              onChange={(e) => setVersionLabel(e.target.value)}
              onKeyDown={(e) => {
                if (e.key !== "Enter") return;
                setVersionModal(false);
                void saveVersion(versionLabel.trim());
              }}
            />
            <div className="row">
              <button className="ghost" data-testid="version-cancel" onClick={() => setVersionModal(false)}>
                Cancel
              </button>
              <button
                data-testid="version-save"
                onClick={() => { setVersionModal(false); void saveVersion(versionLabel.trim()); }}
              >
                Save version
              </button>
            </div>
          </div>
        </div>
      )}
      {showActivity && (
        <div style={{ maxHeight: 120, overflowY: "auto", padding: "0 8px 4px", fontSize: 12, fontFamily: "monospace" }}>
          {/* Attribution L1 (doc 14 §3): the canonical log IS the record —
              clientId joins the roster for a name + color. */}
          {[...(session?.activity ?? [])].reverse().map((a) => {
            const who = session?.roster.find((r) => r.clientId === a.clientId);
            return (
              <div key={a.seq}>
                <span style={{ color: who?.profile.color ?? "#888" }}>{who?.profile.name ?? a.clientId}</span>
                {" "}#{a.seq} {a.kind}
              </div>
            );
          })}
        </div>
      )}
      <div style={{ flex: 1, minHeight: 0 }}>
        <CollabEditor
          key={attempt}
          url={url}
          createSocket={makeSocket}
          docId={docId}
          clientId={clientId}
          store={store}
          takeover={takeover}
          docKey={docKey}
          shareCode={shareCode}
          ownerToken={ownerToken}
          // Media relay origin (doc 16 §3) — the same server the socket
          // points at. Image bytes ride these HTTP routes, never the
          // sequencer; without it the image button stays inert.
          httpBase={httpBase}
          profile={{ name, color: "" /* server assigns a palette color */ }}
          onSession={setSession}
          onReady={setApi}
          refusedContent={refusedContent}
          editable
        />
      </div>
      {session?.sessionWarning && (
        <SessionCountdown warning={session.sessionWarning} onDownload={download} />
      )}
      {session?.connection === "lost" && !lostDismissed && (
        <ConnectionLost
          onRetry={() => session.reconnect()}
          onDownload={download}
          // The safety promise is CONDITIONAL on the durable copy actually
          // existing. Passing this rather than hardcoding "your work is safe"
          // is the whole point: see the copy in the component.
          saved={(session.persistErrors ?? 0) === 0}
          // Offline editing is only offered when it is actually available
          // (a document in hand, no refusal): dismissing into a page that
          // cannot record anything would betray the button's promise.
          onKeepEditing={session.offline ? () => setLostDismissed(true) : undefined}
        />
      )}
      {/* Dev perf menu (Ctrl+Shift+P / ?perf=1) — inert until opened. */}
      <PerfHud clientId={clientId} docStats={docStats} />
    </div>
  );
}

/**
 * The session ended (idle timeout / lifetime cap), shown as a MODAL over the
 * document — which is still on screen behind it, read-only (`writesBlocked`
 * folds refusal in, so the editor gate handles that; this dialog only has to
 * present the ways forward).
 *
 * GENUINELY BLOCKING, unlike the other dialogs here (Save-version, the
 * local editor's share modal), which close on Escape and a scrim click.
 * There is deliberately no dismiss: dismissing would strand the user on a
 * dead document with the three ways forward gone. Every exit is one of the
 * three buttons, so focus moves INTO the dialog on mount and Tab is trapped
 * inside it — that trap is what makes swallowing Escape defensible for a
 * keyboard user (the document behind is read-only; tabbing into it serves
 * nothing).
 *
 * The `no-copy` state keeps only "Start fresh": with no stored bundle there
 * is nothing to revive and nothing to keep editing. It stops being a ghost
 * there — with one action, that action is the primary.
 */
function SessionEndedModal({ reason, reviveState, onRevive, onKeepLocal, onNewDocument }: {
  reason: "idle-timeout" | "session-expired";
  reviveState: "idle" | "reviving" | "no-copy";
  onRevive: () => void;
  onKeepLocal?: () => void;
  onNewDocument?: () => void;
}) {
  const dialogRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    const buttons = () => Array.from(dialog.querySelectorAll<HTMLElement>("button:not([disabled])"));
    buttons()[0]?.focus();
    // Queried per keypress, not captured once: the button set changes when
    // reviveState moves (reviving disables bring-back; no-copy removes it).
    const onKeydown = (e: KeyboardEvent) => {
      if (e.key !== "Tab") return;
      const els = buttons();
      if (els.length === 0) return;
      const idx = els.indexOf(document.activeElement as HTMLElement);
      if (e.shiftKey && idx <= 0) { e.preventDefault(); els[els.length - 1].focus(); }
      else if (!e.shiftKey && (idx === -1 || idx === els.length - 1)) { e.preventDefault(); els[0].focus(); }
    };
    window.addEventListener("keydown", onKeydown, true);
    return () => window.removeEventListener("keydown", onKeydown, true);
  }, []);
  return (
    <div className="modal-scrim" data-testid="session-ended">
      <div ref={dialogRef} className="modal" role="dialog" aria-modal="true" aria-labelledby="session-ended-title">
        <h2 id="session-ended-title">
          {reason === "idle-timeout" ? "This session has ended" : "This session reached its time limit"}
        </h2>
        <p>
          {reason === "idle-timeout"
            ? "This session ended after a long stretch with no edits. Rooms close when nobody is using them — the server keeps nothing between sessions."
            : "This session reached its time limit and ended. Every room has a maximum age, so none of them lives on the server indefinitely."}
        </p>
        {reviveState === "no-copy" ? (
          <p>This browser has no saved copy of the document, so it can’t bring it back. Any participant who edited here before can.</p>
        ) : (
          <p>Your copy is still here in this browser — nothing was lost.</p>
        )}
        <div className="row">
          {reviveState !== "no-copy" && onKeepLocal && (
            <button className="ghost" data-testid="ended-keep-local" onClick={onKeepLocal}>Keep editing on your own</button>
          )}
          {onNewDocument && (
            <button className={reviveState === "no-copy" ? undefined : "ghost"} data-testid="ended-new-document" onClick={onNewDocument}>
              Start fresh
            </button>
          )}
          {reviveState !== "no-copy" && (
            <button data-testid="bring-back" disabled={reviveState === "reviving"} onClick={onRevive}>
              {reviveState === "reviving" ? "Bringing it back…" : "Bring it back live"}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

/**
 * The connection is gone and automatic recovery has stopped.
 *
 * MODAL, unlike the `reconnecting` chip, because the two states demand
 * different things from the user. Reconnecting resolves itself; `lost` never
 * will — the retry budget is exhausted, so nothing further happens unless a
 * human acts. A quiet chip for that state would leave someone waiting forever
 * for a recovery that has already given up.
 *
 * WHY THE REASSURANCE IS HONEST, and where its limits are. The server keeps
 * nothing (zero custody, doc 12): this browser's IndexedDB bundle IS the
 * durable copy. It holds the confirmed document plus the pending queue —
 * edits made before the drop that the server had not yet acknowledged — and
 * those replay verbatim on rejoin, deduplicated by (clientId, clientSeq), so
 * reconnecting cannot double-apply them. Nothing typed AFTER the drop is at
 * risk either, because those edits are recorded to the durable offline tail
 * (doc 15 §2) and replayed on reconnect; that capture is the reason this
 * dialog can make a promise at all — and the reason it now offers "keep
 * editing offline" instead of holding the document hostage.
 *
 * If the room was re-seeded while this browser was away, the rejoin lands in a
 * different epoch and doc 15's arrival ladder takes over — the offline edits
 * come back as tracked-change SUGGESTIONS rather than overwriting the crowd's
 * text, or as a saved draft when there are too many to review. Hence "as
 * suggestions" below rather than a flat "your edits will be restored", which
 * would over-promise the destructive-rebase behaviour the ladder deliberately
 * does not do.
 *
 * `saved` is the one thing that can falsify all of it. A failed bundle write
 * (quota, private mode, blocked storage) means the durable copy silently
 * stopped updating, and the copy MUST change rather than repeat a promise the
 * storage layer did not keep — the download button is the only true escape in
 * that case, so it is the one the dialog leads with. `onKeepEditing` is
 * absent in that case too at the call site's discretion — and always absent
 * when offline editing is unavailable (no document, refusal-blocked).
 */
function ConnectionLost({ onRetry, onDownload, saved, onKeepEditing }: {
  onRetry: () => void;
  onDownload: () => void;
  saved: boolean;
  onKeepEditing?: () => void;
}) {
  return (
    <div
      data-testid="lost-modal"
      role="dialog"
      aria-modal="true"
      aria-labelledby="dxw-lost-title"
      style={{
        position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)",
        display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000,
      }}
    >
      <div style={{
        background: "#fff", borderRadius: 10, padding: 20, maxWidth: 420,
        font: "14px system-ui", boxShadow: "0 8px 32px rgba(0,0,0,0.3)",
      }}>
        <h2 id="dxw-lost-title" style={{ margin: "0 0 8px", fontSize: 16 }}>
          Connection lost
        </h2>
        {saved ? (
          <p style={{ margin: "0 0 12px", lineHeight: 1.5 }}>
            You’ve been disconnected. <strong>Your work is saved in this
            browser</strong> — and you can keep editing: changes you make while
            offline are saved here too. When you reconnect, everything is sent
            automatically — and if someone restored the document while you were
            away, your changes come back as suggestions to review rather than
            overwriting theirs.
          </p>
        ) : (
          <p style={{ margin: "0 0 12px", lineHeight: 1.5 }}>
            You’ve been disconnected, and{" "}
            <strong>this browser could not save a copy of the document</strong>{" "}
            (storage may be full or blocked). Download it now — that is the only
            copy you can rely on.
          </p>
        )}
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button data-testid="lost-retry" onClick={onRetry} autoFocus>
            Try to reconnect
          </button>
          {onKeepEditing && saved && (
            <button data-testid="lost-keep-editing" onClick={onKeepEditing}>
              Keep editing offline
            </button>
          )}
          <button data-testid="lost-reload" onClick={() => location.reload()}>
            Reload the page
          </button>
          <button data-testid="lost-download" onClick={onDownload}>
            Download .docx
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * The grace period before an announced session ending, shown as a live clock.
 *
 * TICKS LOCALLY from the `inMs` the server measured when it sent the warning:
 * the wire carries ONE message per deadline, not a stream of refreshed
 * remainders, so a component waiting to be told the new number would display a
 * frozen one for the entire grace period.
 *
 * The two endings are genuinely different and the copy says so rather than
 * sharing a euphemism:
 *
 *  - `idle`     is cancellable, and typing is what cancels it — so the popup
 *               asks for the one action that works, and DISAPPEARS on its own
 *               when the server clears the deadline. That vanishing IS the
 *               confirmation; there is deliberately no dismiss button, because
 *               dismissing it would hide the only signal that it worked.
 *  - `lifetime` cannot be cancelled by anything, so the copy makes no
 *               suggestion at all. It points at the document's durability
 *               instead (the browser's copy, the download beside it) because
 *               that is the true reassurance — the session ends, the document
 *               does not.
 */
function SessionCountdown({ warning, onDownload }: {
  warning: { reason: "idle" | "lifetime"; inMs: number };
  onDownload: () => void;
}) {
  const idle = warning.reason === "idle";
  const [remaining, setRemaining] = useState(warning.inMs);
  useEffect(() => {
    // Anchor on an absolute deadline, not a decrementing counter: a
    // backgrounded tab throttles timers, and a counter would drift into
    // promising time that had already passed.
    const deadline = Date.now() + warning.inMs;
    setRemaining(warning.inMs);
    const id = setInterval(() => setRemaining(Math.max(0, deadline - Date.now())), 1000);
    return () => clearInterval(id);
    // A fresh warning object = a fresh deadline (the state holds the callback's
    // object, so the identity is stable between warnings).
  }, [warning]);
  const secs = Math.ceil(remaining / 1000);
  const clock = `${Math.floor(secs / 60)}:${String(secs % 60).padStart(2, "0")}`;
  return (
    <div className="countdown" data-testid="session-countdown" data-reason={warning.reason} role="status" aria-live="polite">
      <b>{idle ? "Session closing due to inactivity" : "This session is reaching its time limit"}</b>
      <span className="clock" data-testid="countdown-clock">{clock}</span>
      <p>
        {idle
          ? "Keep editing to stay connected — any edit resets the clock for everyone."
          : "Your copy is safe in this browser; you can re-share it as a new session afterwards."}
      </p>
      {!idle && <button onClick={onDownload}>Download .docx</button>}
    </div>
  );
}
