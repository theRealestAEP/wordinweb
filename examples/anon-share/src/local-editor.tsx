import { useEffect, useRef, useState } from "react";
import { DocxView, DocxToolbar, type DocxViewApi } from "wordinweb";

/**
 * The demo's LANDING is the normal single-user editor: a local, editable
 * document with no server and no collab. It seeds from the zero-custody
 * server's blank template (`GET /blank`) purely to have bytes to edit — the
 * document lives entirely in this browser until the user chooses to share it.
 *
 * Going collaborative is the one and only go-live path (doc 13 §6): it takes
 * the CURRENT edited bytes (`api.save()`), seals them client-side under a fresh
 * doc key, and PUTs the encrypted genesis checkpoint. Collaborative always
 * means encrypted — there is no plaintext-vs-encrypted choice in the UI; the
 * key rides the `#k=` fragment the caller writes into the URL.
 *
 * `initialBytes` is the RETURN path: leaving a session drops back here with the
 * document you were just editing, not a fresh blank. Keeping your copy when you
 * leave the room is the zero-custody claim made tangible — the server forgets
 * you, your document does not.
 */
export function LocalEditor({
  httpBase,
  initialBytes,
  onGoLive,
}: {
  httpBase: string;
  /** Bytes to open instead of the blank template (returning from a session). */
  initialBytes?: Uint8Array;
  /** Seal + PUT the given local bytes, then switch the app into collab mode. */
  onGoLive: (bytes: Uint8Array, shareCode?: string) => Promise<void>;
}) {
  // Blank template to start editing (bytes only — no session is created until
  // go-live). Fetched once; a failure surfaces inline. Skipped entirely when
  // the caller handed us a document to reopen.
  const [blank, setBlank] = useState<Uint8Array | null>(initialBytes ?? null);
  const [loadError, setLoadError] = useState(false);
  const [api, setApi] = useState<DocxViewApi | null>(null);
  const [going, setGoing] = useState(false);
  const [code, setCode] = useState("");
  const [modalOpen, setModalOpen] = useState(false);
  const codeRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (initialBytes) return; // reopening a document — nothing to fetch
    let alive = true;
    void fetch(`${httpBase}/blank`)
      .then((r) => {
        if (!r.ok) throw new Error(`blank ${r.status}`);
        return r.arrayBuffer();
      })
      .then((buf) => {
        if (alive) setBlank(new Uint8Array(buf));
      })
      .catch(() => {
        if (alive) setLoadError(true);
      });
    return () => {
      alive = false;
    };
  }, [httpBase, initialBytes]);

  // Focus the code field when the dialog opens, and let Escape dismiss it —
  // a dialog you cannot leave by keyboard is a dialog that traps people.
  useEffect(() => {
    if (!modalOpen) return;
    codeRef.current?.focus();
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && !going) setModalOpen(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [modalOpen, going]);

  const startCollab = () => {
    // A code is REQUIRED (owner's call): a shared link is a capability, and
    // this is the second factor that keeps a link on its own from being one.
    // The button is disabled without it; this guard covers the Enter key.
    if (!api || !code.trim()) return;
    setGoing(true);
    // Read the CURRENT edited document straight off the imperative API — the
    // exact bytes we seal are what the collaborative session opens to.
    const bytes = api.save();
    void onGoLive(bytes, code.trim()).catch(() => {
      setGoing(false);
      setModalOpen(false);
    });
  };

  if (loadError) {
    return (
      <div style={{ padding: 24 }} data-testid="local-editor-error">
        Couldn’t reach the server at {httpBase} to start a document. Start the zero-custody
        server, then reload.
      </div>
    );
  }
  if (!blank) return <div style={{ padding: 24 }}>Loading editor…</div>;

  return (
    <div data-testid="local-editor" style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div
        data-testid="local-header"
        style={{ display: "flex", gap: 10, alignItems: "center", padding: "4px 8px", flexWrap: "wrap" }}
      >
        <span className="brand-lockup">
          <span className="brand-mark">W</span>
          <span className="brand-copy">
            <strong style={{ color: "#202124" }}>WordInWeb</strong>
            <span className="tag" style={{ color: "#5f6368" }}>Editing locally — nothing has left this browser</span>
          </span>
        </span>
        {/* THE primary action on this screen, and the only one that changes
            what the app is — so it is sized like it and placed ahead of the
            spacer rather than tucked at the far right with the utilities.
            Everything else here is secondary to "share this document". */}
        <button
          data-testid="make-collaborative"
          disabled={!api}
          onClick={() => setModalOpen(true)}
          className="cta"
        >
          Make collaborative
        </button>
        <span style={{ flex: 1 }} />
      </div>
      <DocxToolbar api={api} mode="advanced" />
      <div style={{ flex: 1, minHeight: 0 }}>
        <DocxView source={blank} editable onReady={setApi} />
      </div>

      {modalOpen && (
        <div
          className="modal-scrim"
          data-testid="collab-modal"
          onMouseDown={(e) => { if (e.target === e.currentTarget && !going) setModalOpen(false); }}
        >
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="collab-modal-title">
            <h2 id="collab-modal-title">Start collaborating</h2>
            <p>
              This document is encrypted in your browser and shared by link. The server orders
              edits without ever holding the key.
            </p>
            <label htmlFor="share-code-input">Share code</label>
            <input
              id="share-code-input"
              ref={codeRef}
              data-testid="share-code"
              value={code}
              placeholder="e.g. redwood"
              // Long enough for a passphrase ("redwood-canyon-42" is 17 and
              // was being truncated at 12). Nothing downstream has a length or
              // alphabet opinion — the code is arbitrary text fed to
              // stretchShareCode (PBKDF2-SHA256, 600k, per-doc salt) — so the
              // cap exists only to bound the field, and the JOIN field carries
              // the same one.
              maxLength={64}
              disabled={going}
              onChange={(e) => setCode(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && !going) startCollab(); }}
            />
            <p style={{ margin: "8px 0 0" }}>
              Everyone you share the link with needs this code as well. Send it separately —
              a different app, or in person — so that a link on its own is never enough to
              open the document. It doesn’t change what the server can see: the key stays in
              your browser and never reaches it either way.
            </p>
            <div className="row">
              <button className="ghost" data-testid="collab-cancel" disabled={going} onClick={() => setModalOpen(false)}>
                Cancel
              </button>
              {/* REQUIRED, and the empty field is the whole enforcement: no
                  minimum, no strength rules, no suggestions. Whatever they
                  type is the code. */}
              <button data-testid="start-collab" disabled={going || !code.trim()} onClick={startCollab}>
                {going ? "Starting…" : "Start Collab"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
