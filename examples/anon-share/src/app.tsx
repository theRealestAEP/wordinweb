import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { CollabEditor, IndexedDbBundleStore, InMemoryBundleStore, type CollabSession, type DocBundle } from "wordinweb/collab";
import { reviveEncrypted } from "./e2ee-flows";
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
 * copy, which is why the download button is a first-class control and not
 * a menu item (doc 12 §4 durability honesty).
 */

function b64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

export function App({ url, httpBase, docId, clientId, name, docKey, ownerToken, initialShareCode, onNewDocument, onDisconnect }: {
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
}) {
  const store = useMemo(() => new IndexedDbBundleStore(), []);
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
  const isOwner = !!ownerToken;
  // Remount key: bumps to retry the connection after takeover/revival.
  const [attempt, setAttempt] = useState(0);
  const [showActivity, setShowActivity] = useState(false);
  // Versions (doc 14 §1): frozen restore points beside the live bundle.
  // Stored in the same IndexedDB db under version-suffixed keys — a
  // DocVersion is bundle-shaped enough for the demo (docx + sidecar).
  const [versions, setVersions] = useState<{ label: string; savedAt: number }[]>([]);

  const saveVersion = async () => {
    if (!session?.doc) return;
    const label = window.prompt("Version label (optional):") ?? "";
    const savedAt = Date.now();
    const bundle = await store.get(docId);
    if (!bundle) return;
    await store.put({ ...bundle, docId: `${docId}#version-${savedAt}${label ? `-${label}` : ""}` });
    setVersions((v) => [...v, { label: label || "(auto)", savedAt }].slice(-25));
  };
  const downloadVersion = async (savedAt: number, label: string) => {
    const v = await store.get(`${docId}#version-${savedAt}${label !== "(auto)" ? `-${label}` : ""}`);
    if (!v) return;
    const blob = new Blob([v.confirmedBytes as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    // Restore-as-DRAFT (doc 14 §1): a version opens as a new file, never
    // mutates the live doc — "back to then", not "merge then into now".
    a.download = `${docId.slice(0, 8)}-${label}-${savedAt}.docx`;
    a.click();
    URL.revokeObjectURL(a.href);
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
        <div>
          <p>This document is open in another tab.</p>
          <button onClick={() => { setTakeover(true); setAttempt((a) => a + 1); }}>Use here instead</button>
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
        <div>
          <p>
            {reason === "code-locked"
              ? "Too many tries — wait a minute, then re-enter the share code."
              : reason === "code-invalid"
                ? "That code didn't match — ask whoever shared the document."
                : "This document needs its share code (sent separately from the link)."}
          </p>
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
          {onNewDocument && (
            <p style={{ marginTop: 8, fontSize: 13 }}>
              Forgot the code?{" "}
              <button data-testid="refused-new-document" onClick={onNewDocument}>Start a new document</button>
            </p>
          )}
        </div>
      );
    }
    if (reason === "idle-timeout" || reason === "session-expired") {
      // THE DEADLINE ARRIVING, not a refusal. The countdown that preceded it
      // explained what was about to happen, so this must read as the ending
      // it announced — "refresh to retry" would be both wrong (there is
      // nothing to retry) and a lie about where the document lives.
      return (
        <div data-testid="session-ended">
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
          {reviveState !== "no-copy" && (
            <>
              <button data-testid="bring-back" disabled={reviveState === "reviving"} onClick={() => void revive()}>
                {reviveState === "reviving" ? "Bringing it back…" : "Bring it back live"}
              </button>{" "}
              {onDisconnect && (
                <button data-testid="ended-keep-local" onClick={leaveSession}>Keep editing on your own</button>
              )}{" "}
            </>
          )}
          {onNewDocument && <button data-testid="ended-new-document" onClick={onNewDocument}>Start fresh</button>}
        </div>
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
    return <p>Connection refused: {reason}. Refresh to retry.</p>;
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div data-testid="toolbar" style={{ display: "flex", gap: 8, alignItems: "center", padding: "4px 8px", flexWrap: "wrap" }}>
        {/* Roster chips (doc 14 §2): everyone in the session, greyed when
            offline; names are self-asserted — the UI says so in the title. */}
        {(session?.roster ?? []).map((r) => (
          <span key={r.clientId} data-testid="roster-chip" data-connected={r.connected} style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
            <span
              title="Names are chosen by participants"
              style={{
                padding: "2px 10px", borderRadius: 999, fontSize: 12, color: "#fff",
                background: r.profile.color, opacity: r.connected ? 1 : 0.35,
              }}
            >
              {r.profile.name}
            </span>
            {isOwner && r.clientId !== clientId && (
              <>
                <button title="Demote to viewer" style={{ fontSize: 10, padding: "0 3px" }}
                  onClick={() => session?.admin({ op: "setRole", clientId: r.clientId, role: "viewer" })}>👁</button>
                <button title="Kick" style={{ fontSize: 10, padding: "0 3px" }}
                  onClick={() => session?.admin({ op: "kick", clientId: r.clientId })}>✕</button>
              </>
            )}
          </span>
        ))}
        <span style={{ flex: 1 }} />
        {session?.arrival ? (
          <span style={{ fontSize: 12, background: "#d1ecf1", padding: "2px 8px", borderRadius: 6 }}>
            You edited offline ({session.arrival.tailLength} change{session.arrival.tailLength === 1 ? "" : "s"}).{" "}
            {session.arrival.mode === "suggest" ? (
              <>
                <button onClick={() => session.reconcile("suggest")}>Add my changes as suggestions</button>{" "}
                <button onClick={() => session.reconcile("draft")}>Keep as draft</button>
              </>
            ) : (
              <>
                Too many to suggest cleanly — <button onClick={() => session.reconcile("draft")}>keep as a draft</button>.
              </>
            )}
          </span>
        ) : (session?.persistErrors ?? 0) > 0 ? (
          // DATA-LOSS WARNING, deliberately ahead of the other banners. This
          // browser's stored bundle IS the durable copy (doc 12 §4) — the
          // server keeps nothing — so a failed write means the document may
          // exist nowhere but this tab's memory. Quota exceeded, private
          // mode, or storage blocked all land here. The download button
          // beside this banner is the escape hatch, which is why the copy
          // points at it rather than just apologising.
          <span data-testid="persist-banner" style={{ fontSize: 12, background: "#f8d7da", padding: "2px 8px", borderRadius: 6 }}>
            This document may not be saved in this browser — download a copy.
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
        {/* The escape hatch (doc 12 §4): your copy is as durable as this
            browser; one click makes it a file. */}
        <button data-testid="download" onClick={download} disabled={!session?.ready}>Download .docx</button>
        {onDisconnect && (
          <button
            data-testid="disconnect"
            title="Leave this session and keep editing your copy locally"
            onClick={leaveSession}
          >
            Disconnect
          </button>
        )}
        <button onClick={() => void saveVersion()} disabled={!session?.ready}>Save version</button>
        <button onClick={() => setShowActivity((v) => !v)}>{showActivity ? "Hide" : "Show"} activity</button>
        {isOwner && (
          <button
            data-testid="readonly-toggle"
            title="Owner: block all editors (you keep writing)"
            onClick={() => { const on = !readOnly; setReadOnly(on); session?.admin({ op: "readOnly", on }); }}
          >
            {readOnly ? "🔒 Read-only ON" : "🔓 Make read-only"}
          </button>
        )}
      </div>
      {versions.length > 0 && (
        <div style={{ display: "flex", gap: 6, padding: "0 8px 4px", fontSize: 12, flexWrap: "wrap" }}>
          {versions.map((v) => (
            <button key={v.savedAt} title="Opens as a new file — never merges into the live doc"
              onClick={() => void downloadVersion(v.savedAt, v.label)}>
              ⏱ {v.label} · {new Date(v.savedAt).toLocaleTimeString()}
            </button>
          ))}
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
          refusedContent={refusedContent}
          editable
        />
      </div>
      {session?.sessionWarning && (
        <SessionCountdown warning={session.sessionWarning} onDownload={download} />
      )}
      {/* Dev perf menu (Ctrl+Shift+P / ?perf=1) — inert until opened. */}
      <PerfHud clientId={clientId} docStats={docStats} />
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
