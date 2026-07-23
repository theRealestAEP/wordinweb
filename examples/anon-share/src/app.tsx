import { useMemo, useState, type ReactNode } from "react";
import { CollabEditor, IndexedDbBundleStore, type CollabSession, type DocBundle } from "wordinweb/collab";
import { reviveEncrypted } from "./e2ee-flows";

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

export function App({ url, httpBase, docId, clientId, name, docKey }: {
  url: string;
  httpBase: string;
  docId: string;
  clientId: string;
  name: string;
  /** E2EE mode (doc 13): present iff the link carried `#k=` — the mode is
   * the LINK's fact; the editor hard-refuses a contradicting welcome. */
  docKey?: string;
}) {
  const store = useMemo(() => new IndexedDbBundleStore(), []);
  const [session, setSession] = useState<CollabSession | null>(null);
  const [takeover, setTakeover] = useState(false);
  const [reviveState, setReviveState] = useState<"idle" | "reviving" | "no-copy">("idle");
  // Share code (doc 13 §7): entered on a code-required/invalid refusal;
  // remounting with it retries (and in E2EE mode mixes into key derivation).
  const [shareCode, setShareCode] = useState<string | undefined>(undefined);
  const [codeDraft, setCodeDraft] = useState("");
  // Remount key: bumps to retry the connection after takeover/revival.
  const [attempt, setAttempt] = useState(0);

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
        body: JSON.stringify({ docx: b64(bundle.confirmedBytes), sidecar: bundle.confirmedSidecar }),
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
            value={codeDraft}
            onChange={(e) => setCodeDraft(e.target.value)}
            placeholder="6-digit share code"
            inputMode="numeric"
            maxLength={12}
          />{" "}
          <button
            disabled={!codeDraft || reason === "code-locked"}
            onClick={() => { setShareCode(codeDraft); setAttempt((a) => a + 1); }}
          >
            Join
          </button>
        </div>
      );
    }
    return <p>Connection refused: {reason}. Refresh to retry.</p>;
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "4px 8px", flexWrap: "wrap" }}>
        {/* Roster chips (doc 14 §2): everyone in the session, greyed when
            offline; names are self-asserted — the UI says so in the title. */}
        {(session?.roster ?? []).map((r) => (
          <span
            key={r.clientId}
            title="Names are chosen by participants"
            style={{
              padding: "2px 10px", borderRadius: 999, fontSize: 12, color: "#fff",
              background: r.profile.color, opacity: r.connected ? 1 : 0.35,
            }}
          >
            {r.profile.name}
          </span>
        ))}
        <span style={{ flex: 1 }} />
        {session?.epochChanged && (
          <span style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            Restored by another participant — your offline copy is saved as a draft.
          </span>
        )}
        {/* The escape hatch (doc 12 §4): your copy is as durable as this
            browser; one click makes it a file. */}
        <button onClick={download} disabled={!session?.ready}>Download .docx</button>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <CollabEditor
          key={attempt}
          url={url}
          docId={docId}
          clientId={clientId}
          store={store}
          takeover={takeover}
          docKey={docKey}
          shareCode={shareCode}
          profile={{ name, color: "" /* server assigns a palette color */ }}
          onSession={setSession}
          refusedContent={refusedContent}
          editable
        />
      </div>
    </div>
  );
}
