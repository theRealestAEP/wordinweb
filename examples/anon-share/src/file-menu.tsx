import { useEffect, useRef, useState } from "react";
import type { StoredDocSummary } from "wordinweb/collab";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/**
 * Display name for a stored entry. NEVER the docId — the docId is the room
 * capability (doc 11): anyone holding it can open the session, so it must
 * not appear anywhere a screenshot could leak it. Kind + date + size is
 * what distinguishes entries instead.
 */
export function savedDocName(s: StoredDocSummary): string {
  if (s.title) return s.title;
  switch (s.kind) {
    case "local": return "Local document";
    case "live": return "Shared document (saved copy)";
    case "version": return s.label ? `Version: ${s.label}` : "Version (unnamed)";
    case "draft": return "Offline draft";
    case "superseded": return "Superseded copy";
    default: return "Stored data";
  }
}

export function fmtSize(bytes: number): string {
  if (bytes >= 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
  if (bytes >= 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

function fmtDate(ms: number): string {
  return new Date(ms).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}

/**
 * The demo's File menu: New / Open / Download / Print behind one button.
 *
 * DUMB BY DESIGN — every item is a callback the parent owns. The two screens
 * that mount it (the local editor and the collaborative one) mean different
 * things by "new document" and cannot share an implementation, so this
 * component decides only what a menu decides: what is on it, whether it is
 * open, and how you leave it.
 *
 * An item whose handler is absent is not rendered. The ONE exception is Open,
 * which the collaborative screen shows greyed out rather than hides, so the
 * menu holds still between the two screens.
 */
export function FileMenu({
  onNew,
  onOpen,
  onDownload,
  onPrint,
  disabled,
  openDisabled,
  listSaved,
  onOpenSaved,
  canRejoinSaved,
  onRejoinSaved,
  savedOpenDisabled,
  savedOpenTitle,
  onDownloadSaved,
  onDeleteSaved,
  openRequest,
}: {
  onNew?: () => void;
  /** Receives the picked file's bytes and its name. */
  onOpen?: (bytes: Uint8Array, filename: string) => void;
  onDownload?: () => void;
  onPrint?: () => void;
  disabled?: boolean;
  /** Render Open greyed out. Wins over `onOpen` if both are passed. */
  openDisabled?: boolean;
  /** Enumerate this browser's saved documents/versions/drafts (metadata
   * only). When present the menu shows a "Saved in this browser" section,
   * re-listed every time the menu opens so it reflects storage as it is. */
  listSaved?: () => Promise<StoredDocSummary[]>;
  /** Open a saved entry (replaces the current document). */
  onOpenSaved?: (s: StoredDocSummary) => void;
  /** True when this browser retained the room URL for a live saved entry. */
  canRejoinSaved?: (s: StoredDocSummary) => boolean;
  /** Rejoin the collaborative room behind a saved entry. */
  onRejoinSaved?: (s: StoredDocSummary) => void;
  /** Grey saved entries' open action out — the collaborative screen: opening
   * a stored copy there would replace the shared document and silently fork
   * the room (the same rule as `openDisabled`). Download stays live. */
  savedOpenDisabled?: boolean;
  /** Explains what opening a saved copy does on the current screen. */
  savedOpenTitle?: string;
  /** Download a saved entry as .docx without opening it. */
  onDownloadSaved?: (s: StoredDocSummary) => void;
  /** Delete a saved entry. The menu makes this two-step (confirm in place)
   * so it is never a one-click accident beside "open". */
  onDeleteSaved?: (s: StoredDocSummary) => void | Promise<void>;
  /** Bump to a new non-zero value to open the menu from outside — the
   * storage-full banner's "Manage saved copies" route to the delete list. */
  openRequest?: number;
}) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);
  /** null = not loaded yet (or loading); [] = loaded, nothing there. */
  const [saved, setSaved] = useState<StoredDocSummary[] | null>(null);
  const [savedError, setSavedError] = useState(false);
  /** Key of the entry whose Delete is one click from firing. */
  const [confirmDelete, setConfirmDelete] = useState<string | null>(null);
  /** Origin storage usage vs quota, where the browser reports one. This is
   * the only durable copy of the user's work (zero custody) — how full it is
   * must not be a mystery until it breaks. Null where estimate() is absent
   * or fails; the listing's own summed sizes stand in then. */
  const [quota, setQuota] = useState<{ used: number; total: number } | null>(null);

  // (Re-)list on every open: storage may have changed since last time, and a
  // stale listing here would promise documents that are already gone.
  const refreshSaved = () => {
    if (!listSaved) return;
    setSavedError(false);
    listSaved().then(
      (all) => setSaved([...all].sort((a, b) => b.savedAt - a.savedAt)),
      () => { setSaved(null); setSavedError(true); },
    );
    try {
      void navigator.storage?.estimate?.().then(
        (est) => {
          if (est?.quota && Number.isFinite(est.quota)) setQuota({ used: est.usage ?? 0, total: est.quota });
        },
        () => {},
      );
    } catch {
      // estimate() unavailable — the summed listing sizes stand in below.
    }
  };
  useEffect(() => {
    if (!open) { setConfirmDelete(null); return; }
    refreshSaved();
  }, [open]);

  // External open request (the storage-full banner). Bumping the counter
  // opens the menu so the delete list is one click from the warning.
  useEffect(() => {
    if (openRequest) setOpen(true);
  }, [openRequest]);

  // Outside mousedown and Escape both close. mousedown rather than click so the
  // menu is gone before whatever was clicked underneath reacts.
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    window.addEventListener("mousedown", onDown);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("mousedown", onDown);
      window.removeEventListener("keydown", onKey);
    };
  }, [open]);

  /** Every item closes the menu, then acts. */
  const item = (run: () => void) => () => {
    setOpen(false);
    run();
  };

  const pick = async () => {
    const input = fileRef.current;
    const file = input?.files?.[0];
    // Cleared BEFORE the await: without it the same file picked twice fires no
    // change event the second time, so re-opening a document silently does
    // nothing.
    if (input) input.value = "";
    if (!file) return;
    const buf = await file.arrayBuffer();
    onOpen?.(new Uint8Array(buf), file.name);
  };

  // `openDisabled` is the gate, not the missing handler: a screen that passes
  // both gets the disabled item, never a live Open. The two are contradictory
  // props and the safe reading of the contradiction is the one that cannot
  // fork a document.
  const blocked = !!openDisabled;
  const showOpen = blocked || !!onOpen;

  return (
    <div className="filemenu" ref={rootRef}>
      <button
        data-testid="file-menu"
        className="filemenu-trigger"
        aria-haspopup="menu"
        aria-expanded={open}
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        File<span aria-hidden="true" className="filemenu-caret">⌄</span>
      </button>
      {onOpen && !blocked && (
        <input
          ref={fileRef}
          type="file"
          accept={`.docx,${DOCX_MIME}`}
          hidden
          onChange={() => void pick()}
        />
      )}
      {open && (
        <div className="filemenu-panel" data-testid="file-menu-panel" role="menu">
          {onNew && (
            <button role="menuitem" data-testid="file-new" onClick={item(onNew)}>
              New document
            </button>
          )}
          {showOpen && (
            <>
              <button
                role="menuitem"
                data-testid="file-open"
                disabled={blocked}
                onClick={blocked ? undefined : item(() => fileRef.current?.click())}
              >
                Open .docx…
              </button>
            </>
          )}
          {onDownload && (
            <button role="menuitem" data-testid="file-download" onClick={item(onDownload)}>
              Download .docx
            </button>
          )}
          {onPrint && (
            <button role="menuitem" data-testid="file-print" onClick={item(onPrint)}>
              Print
            </button>
          )}
          {listSaved && (
            <div className="filemenu-saved" data-testid="file-saved-section">
              {/* "In this browser" is load-bearing copy: zero custody means
                  these exist nowhere else — the server stores nothing. */}
              <div className="filemenu-heading">Saved in this browser</div>
              {/* Usage against quota where the browser reports one; otherwise
                  the listing's own summed sizes. estimate() may lie, so this
                  is information, never a gate. */}
              {quota ? (
                <div className="filemenu-note" data-testid="file-saved-quota">
                  {fmtSize(quota.used)} used of {fmtSize(quota.total)} this browser allows the site
                </div>
              ) : saved && saved.length > 0 ? (
                <div className="filemenu-note" data-testid="file-saved-quota">
                  Saved copies use {fmtSize(saved.reduce((n, s) => n + s.byteLength, 0))} in this browser
                </div>
              ) : null}
              {savedError ? (
                <div className="filemenu-note" data-testid="file-saved-error">
                  Browser storage is unavailable. Try again outside private browsing.
                </div>
              ) : saved === null ? (
                <div className="filemenu-note">Loading…</div>
              ) : saved.length === 0 ? (
                <div className="filemenu-note" data-testid="file-saved-empty">Nothing saved in this browser yet.</div>
              ) : (
                saved.map((s) => {
                  const rejoinable = Boolean(onRejoinSaved && canRejoinSaved?.(s));
                  const openDisabled = savedOpenDisabled || !onOpenSaved;
                  return (
                  <div className="filemenu-saveditem" key={s.key} data-testid="file-saved-entry" data-kind={s.kind}>
                    <button
                      role="menuitem"
                      data-testid="file-saved-open"
                      disabled={openDisabled}
                      title={
                        savedOpenDisabled
                          ? "Opening would replace the shared document for you alone and split you from the session. Leave the session first, or download a copy."
                          : savedOpenTitle ?? "Open this saved copy (replaces the current document)"
                      }
                      onClick={openDisabled ? undefined : item(() => onOpenSaved!(s))}
                    >
                      <span className="filemenu-savedname">{savedDocName(s)}</span>
                      <span className="filemenu-savedmeta">{fmtDate(s.savedAt)} · {fmtSize(s.byteLength)}</span>
                    </button>
                    {onDownloadSaved && (
                      <button
                        className="filemenu-mini"
                        data-testid="file-saved-download"
                        title="Download this saved copy as .docx"
                        onClick={item(() => onDownloadSaved(s))}
                      >
                        ⬇
                      </button>
                    )}
                    {rejoinable && (
                      <button
                        className="filemenu-mini"
                        data-testid="file-saved-rejoin"
                        title="Rejoin this collaborative room"
                        onClick={item(() => onRejoinSaved!(s))}
                      >
                        Rejoin
                      </button>
                    )}
                    {onDeleteSaved && (
                      confirmDelete === s.key ? (
                        <button
                          className="filemenu-mini danger"
                          data-testid="file-saved-delete-confirm"
                          title="Delete this saved browser copy."
                          onClick={() => {
                            setConfirmDelete(null);
                            void Promise.resolve(onDeleteSaved(s)).then(refreshSaved, () => setSavedError(true));
                          }}
                        >
                          Really delete?
                        </button>
                      ) : (
                        <button
                          className="filemenu-mini"
                          data-testid="file-saved-delete"
                          title="Delete this saved copy (asks once more)"
                          onClick={() => setConfirmDelete(s.key)}
                        >
                          ✕
                        </button>
                      )
                    )}
                  </div>
                  );
                })
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
