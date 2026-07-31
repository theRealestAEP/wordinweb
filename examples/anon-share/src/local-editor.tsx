import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { DocxView, DocxToolbar, type DocxViewApi } from "wordinweb";
import { IndexedDbBundleStore, type BundleStore, type DocBundle, type StoredDocSummary } from "wordinweb/collab";
import { FileMenu, fmtSize, savedDocName } from "./file-menu";
import { useStartupReclaim } from "./startup-reclaim";
import { type GoLivePhase } from "./e2ee-flows";

/** The single autosave slot the local editor writes (key convention:
 * `local:` marks documents that never went live — parseBundleKey knows it). */
export const LOCAL_AUTOSAVE_KEY = "local:autosave";
/** An archived local document (File > New banks the current one here). */
const localDocKey = (ts: number) => `local:doc-${ts}`;

/** Wrap plain local bytes in the bundle shape the shared store persists.
 * The collab-only fields are inert placeholders — a local document has no
 * epoch, no pending queue and no lineage until it goes live. */
function localBundle(key: string, bytes: Uint8Array): DocBundle {
  return {
    docId: key,
    genesisId: "local",
    confirmedSeq: 0,
    confirmedBytes: bytes,
    confirmedSidecar: { next: 1, entries: [] },
    pending: [],
    clientSeq: 0,
    savedAt: Date.now(),
    lineage: [],
  };
}

/**
 * Autosave cadence. A full `api.save()` serialises the whole document — on a
 * large one that is megabytes — so unlike the live bundle's 1s throttle this
 * runs on a 10s interval gated by a dirty flag, with a best-effort flush on
 * pagehide/hidden/unmount (the BundlePersister pattern: the interval is the
 * durability guarantee, the flush narrows the tail). Injectable for tests.
 */
const AUTOSAVE_MS = 10_000;

/**
 * IndexedDB can BLOCK without failing (an old tab holding the pre-upgrade
 * schema open, a hung storage layer) — a promise that neither resolves nor
 * rejects. Racing a generous deadline turns that silence into the same
 * failure story a rejection gets: restore falls back to the blank template,
 * a write raises the banner. 5s is far above any healthy IndexedDB op.
 */
function withDeadline<T>(p: Promise<T>, ms = 5000): Promise<T> {
  return Promise.race([
    p,
    new Promise<T>((_, reject) => setTimeout(() => reject(new Error("storage timeout")), ms)),
  ]);
}

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
  canRejoinSaved,
  onRejoinSaved,
  onOpenSavedRoom,
  store: storeProp,
  autosaveMs = AUTOSAVE_MS,
}: {
  httpBase: string;
  /** Bytes to open instead of the blank template (returning from a session). */
  initialBytes?: Uint8Array;
  /** Seal + PUT the given local bytes, then switch the app into collab mode.
   * `onPhase` reports pipeline progress for the overlay below — on a 500-page
   * document the seal + upload is real seconds of work. */
  onGoLive: (bytes: Uint8Array, shareCode?: string, onPhase?: (phase: GoLivePhase) => void) => Promise<void>;
  /** Identify saved entries whose collaborative room URL remains available. */
  canRejoinSaved?: (s: StoredDocSummary) => boolean;
  /** Rejoin the collaborative room for a saved entry. */
  onRejoinSaved?: (s: StoredDocSummary) => void;
  /** Open a room-linked saved entry in the shared offline workspace. */
  onOpenSavedRoom?: (s: StoredDocSummary) => void;
  /** Where local work persists (shared with the collab screen). Defaults to
   * this browser's IndexedDB; tests inject the in-memory store. */
  store?: BundleStore;
  /** Autosave interval override (tests). */
  autosaveMs?: number;
}) {
  const store = useMemo(() => storeProp ?? new IndexedDbBundleStore(), [storeProp]);
  // Blank template to start editing (bytes only — no session is created until
  // go-live). Fetched once; a failure surfaces inline. Skipped entirely when
  // the caller handed us a document to reopen, or when the autosave slot has
  // one from a previous visit — refreshing must not lose local work.
  const [blank, setBlank] = useState<Uint8Array | null>(initialBytes ?? null);
  const [loadError, setLoadError] = useState(false);
  const [api, setApi] = useState<DocxViewApi | null>(null);
  const [going, setGoing] = useState(false);
  const [code, setCode] = useState("");
  const [modalOpen, setModalOpen] = useState(false);
  /** Name of the file the user opened, so Download gives it back unchanged. */
  const [openedName, setOpenedName] = useState<string | null>(null);
  /** Name of the file currently being parsed and laid out, or null. */
  const [loading, setLoading] = useState<string | null>(null);
  const codeRef = useRef<HTMLInputElement | null>(null);
  /**
   * The autosave failed — said out loud, with the escape hatch beside it.
   * This browser's storage is the ONLY durable copy (zero custody), so a
   * swallowed quota error here would let someone believe work is saved when
   * it exists nowhere but this tab's memory.
   */
  const [persistError, setPersistError] = useState<string | null>(null);
  /** Startup reclaim (startup-reclaim.ts): copies leaked before retention
   * existed are removed on load, and what was removed is reported here. */
  const reclaim = useStartupReclaim(store);
  /** What the store holds, measured when a persist write fails — the banner
   * must say what is taking the space, not just that space ran out. */
  const [storedSummary, setStoredSummary] = useState<{ count: number; bytes: number } | null>(null);
  /** Bumped by the banner's "Manage saved copies" to open the File menu. */
  const [savedMenuRequest, setSavedMenuRequest] = useState(0);
  useEffect(() => {
    if (!persistError) return;
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
  }, [persistError, store]);

  useEffect(() => {
    if (initialBytes) return; // reopening a document — nothing to fetch
    let alive = true;
    // The autosave slot first: a refresh reopens the document being edited.
    // A failed read (blocked storage) falls through to the blank template —
    // nothing was lost, and the first autosave WRITE failing is what raises
    // the banner.
    void withDeadline(store.get(LOCAL_AUTOSAVE_KEY))
      .catch(() => null)
      .then((saved) => {
        if (alive && saved) {
          // The same overlay Open shows: restoring a big autosaved document
          // parses and lays out the whole thing synchronously, and a cold
          // load that freezes for seconds with no feedback reads as a broken
          // app (measured: ~5.6s to first page on a ~170-page document —
          // e2e/coldload.spec.ts). Cleared by DocxView's onLoad/onError.
          setLoading("your document");
          // Two frames so the spinner PAINTS before the parse starts — the
          // openDocument trap (fea7e44): one rAF still beats the paint.
          requestAnimationFrame(() => requestAnimationFrame(() => {
            if (alive) setBlank(saved.confirmedBytes);
          }));
          return;
        }
        return fetch(`${httpBase}/blank`)
          .then((r) => {
            if (!r.ok) throw new Error(`blank ${r.status}`);
            return r.arrayBuffer();
          })
          .then((buf) => {
            if (alive) setBlank(new Uint8Array(buf));
          });
      })
      .catch(() => {
        if (alive) setLoadError(true);
      });
    return () => {
      alive = false;
    };
  }, [httpBase, initialBytes, store]);

  /**
   * AUTOSAVE. Dirty is set by user input (captured at the root, so toolbar
   * clicks count too) and by every document (re)load; the interval writes
   * only when dirty, so an idle tab costs nothing. A failed write re-arms
   * the flag — the next tick retries — and raises the banner.
   */
  const dirtyRef = useRef(false);
  const apiRef = useRef<DocxViewApi | null>(null);
  const markDirty = useCallback(() => { dirtyRef.current = true; }, []);
  const saveNow = useCallback(async () => {
    const a = apiRef.current;
    if (!a || !dirtyRef.current) return;
    dirtyRef.current = false;
    try {
      await withDeadline(store.put(localBundle(LOCAL_AUTOSAVE_KEY, a.save())));
      setPersistError(null);
    } catch {
      dirtyRef.current = true;
      setPersistError("Browser storage is full or unavailable. Download a copy.");
    }
  }, [store]);
  useEffect(() => {
    const id = setInterval(() => void saveNow(), autosaveMs);
    const flush = () => void saveNow();
    const onVis = () => { if (document.visibilityState === "hidden") flush(); };
    window.addEventListener("pagehide", flush);
    document.addEventListener("visibilitychange", onVis);
    return () => {
      clearInterval(id);
      window.removeEventListener("pagehide", flush);
      document.removeEventListener("visibilitychange", onVis);
      flush(); // leaving the screen (go-live, unmount) lands the last edits
    };
  }, [saveNow, autosaveMs]);

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

  /**
   * FILE MENU — New. Fetches its own blank rather than clearing state and
   * letting the effect above re-run: that effect is keyed on `[httpBase,
   * initialBytes]`, neither of which changes here, and it returns early
   * forever once the caller handed us a document to reopen. Leaning on it
   * would make New a no-op on exactly the path people reach it from — back
   * from a session, or after opening a file.
   */
  const newDocument = () => {
    // Bank the CURRENT document under a dated key before blanking — "New"
    // must never be how work disappears. The archive is what File > Saved
    // lists as a historical local document. If the archive write fails the
    // blanking is ABORTED: proceeding would let the next autosave overwrite
    // the slot, and the banner + Download are the only honest exits then.
    const archive = apiRef.current
      ? store.put(localBundle(localDocKey(Date.now()), apiRef.current.save()))
      : Promise.resolve();
    void archive
      .then(() => {
        dirtyRef.current = false;
        return store.delete(LOCAL_AUTOSAVE_KEY).catch(() => {
          /* slot cleanup only — the archive already holds the document */
        });
      })
      .catch(() => {
        setPersistError("Download this document before starting a new one. Browser storage is full or unavailable.");
        throw new Error("archive-failed");
      })
      .then(() => fetch(`${httpBase}/blank`))
      .then((r) => {
        if (!r.ok) throw new Error(`blank ${r.status}`);
        return r.arrayBuffer();
      })
      .then((buf) => {
        setOpenedName(null);
        setBlank(new Uint8Array(buf));
      })
      .catch((e: Error) => {
        if (e.message !== "archive-failed") setLoadError(true);
      });
  };

  /** Open a saved entry from the File menu. Room-linked copies stay on the
   * room-aware offline path; ordinary local copies replace this document. */
  const openSaved = (s: StoredDocSummary) => {
    if (onOpenSavedRoom && canRejoinSaved?.(s)) {
      onOpenSavedRoom(s);
      return;
    }
    void store
      .get(s.key)
      .then((b) => {
        if (!b) throw new Error("gone");
        openDocument(b.confirmedBytes, `${savedDocName(s)}.docx`);
      })
      .catch(() => setPersistError("Couldn’t read that saved copy from this browser’s storage."));
  };

  /** Download a saved entry without opening it. */
  const downloadSaved = (s: StoredDocSummary) => {
    void store
      .get(s.key)
      .then((b) => {
        if (!b) throw new Error("gone");
        const blob = new Blob([b.confirmedBytes as BlobPart], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = `${savedDocName(s)}.docx`;
        a.click();
        URL.revokeObjectURL(a.href);
      })
      .catch(() => setPersistError("Couldn’t read that saved copy from this browser’s storage."));
  };

  /**
   * FILE MENU — Open. `setBlank` alone re-sources the view: DocxView's load
   * effect lists `source` in its deps and its document cache is keyed by the
   * source's IDENTITY (`docCacheRef.current?.source === source`), so a fresh
   * Uint8Array parses fresh and fires `onReady` with an api over the new
   * document. No remount key is needed.
   */
  const openDocument = (bytes: Uint8Array, filename: string) => {
    setOpenedName(filename);
    // Shown until DocxView reports the document laid out. Parsing and
    // paginating a large .docx takes real seconds, and a blank stage in the
    // meantime reads as a broken app rather than a busy one.
    setLoading(filename);
    // THE SPINNER MUST PAINT BEFORE THE PARSE STARTS, and it does not get to
    // if both state updates land in one batch: React commits them together,
    // DocxView's effect then parses and lays out the document SYNCHRONOUSLY,
    // and the browser reaches its first paint only after all of that — so the
    // overlay is mounted and unmounted without ever being on screen. That is
    // exactly what "still no spinner when loading a big document" was.
    //
    // Two frames, not one: a single requestAnimationFrame callback runs BEFORE
    // the paint it is scheduled alongside, so the work would still beat the
    // pixels. The second frame is what guarantees the spinner is actually up.
    requestAnimationFrame(() => requestAnimationFrame(() => setBlank(bytes)));
  };

  /** FILE MENU — Download. The edited bytes, under the name they came in as. */
  const downloadDocument = () => {
    if (!api) return;
    const blob = new Blob([api.save() as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = openedName ?? "document.docx";
    a.click();
    URL.revokeObjectURL(a.href);
  };

  /** Go-live progress text over the document, or null. Same overlay pattern
   * as `loading` above — the honest alternative to a frozen tab while a big
   * document serialises, seals and uploads. */
  const [goPhase, setGoPhase] = useState<string | null>(null);
  /** The last go-live failure, shown in the reopened modal. */
  const [goError, setGoError] = useState<string | null>(null);

  const phaseCopy: Record<GoLivePhase, string> = {
    encrypt: "Encrypting…",
    upload: "Uploading…",
  };

  const startCollab = () => {
    // A code is REQUIRED (owner's call): a shared link is a capability, and
    // this is the second factor that keeps a link on its own from being one.
    // The button is disabled without it; this guard covers the Enter key.
    if (!api || !code.trim()) return;
    setGoing(true);
    setGoError(null);
    setModalOpen(false);
    setGoPhase("Saving your document…");
    // THE OVERLAY MUST PAINT BEFORE THE SERIALISE. `api.save()` on a 500-page
    // document is hundreds of synchronous milliseconds; run it in the same
    // tick as the state updates above and the browser's first paint happens
    // only after all of it — the exact openDocument trap (fea7e44). Two
    // frames, not one: a single rAF callback still runs before the paint.
    requestAnimationFrame(() => requestAnimationFrame(() => {
      // Read the CURRENT edited document straight off the imperative API — the
      // exact bytes we seal are what the collaborative session opens to.
      const t0 = performance.now();
      const bytes = api.save();
      // Same grep-able shape as scripts/bench-golive.mjs, so field reports of
      // a slow go-live come with the serialise number attached.
      // eslint-disable-next-line no-console
      console.log(`STRESS-METRIC golive-serialize docxBytes=${bytes.byteLength} saveMs=${(performance.now() - t0).toFixed(2)}`);
      void onGoLive(bytes, code.trim(), (p) => setGoPhase(phaseCopy[p])).catch((err: unknown) => {
        // Refused (too large, server down): back to the editor with the
        // reason on screen — landing on a dead "session" screen or silently
        // resetting the button are both lies about what happened.
        setGoing(false);
        setGoPhase(null);
        setGoError(err instanceof Error ? err.message : "Couldn’t start collaboration. Try again.");
        setModalOpen(true);
      });
    }));
  };

  if (loadError) {
    return (
      <div style={{ padding: 24 }} data-testid="local-editor-error">
        Couldn’t connect to {httpBase}. Check the collaboration server, then reload.
      </div>
    );
  }
  if (!blank) {
    // A big autosaved document is about to parse + paginate synchronously
    // (~5.6s to first page on ~170 pages — e2e/coldload.spec.ts), and the
    // frame painted HERE is what stays on screen through that freeze. The
    // styled spinner, not bare text, is that frame whenever a restore is in
    // flight. `position: relative` anchors the overlay's inset:0.
    return (
      <div style={{ position: "relative", height: "100%", padding: 24 }}>
        {loading ? (
          <div className="document-loading" data-testid="doc-loading" role="status" aria-live="polite" aria-busy="true">
            <span className="document-loading-spinner" aria-hidden="true" />
            <strong>Opening {loading}…</strong>
            <span>Large documents may take a few seconds.</span>
          </div>
        ) : (
          "Opening editor…"
        )}
      </div>
    );
  }

  return (
    <div
      data-testid="local-editor"
      style={{ display: "flex", flexDirection: "column", height: "100%" }}
      // Dirty tracking for autosave — capture phase so the editor and the
      // toolbar both count, without either knowing about persistence.
      onKeyDownCapture={markDirty}
      onMouseUpCapture={markDirty}
      onPasteCapture={markDirty}
      onDropCapture={markDirty}
    >
      {/* Same product header as the collaborative screen. It was a light-text
          variant here, which made the two halves of one app look like two
          different products depending on whether you had shared yet. */}
      <div className="app-header">
        <span className="brand-lockup">
          <span className="brand-mark">W</span>
          <span className="brand-copy">
            <strong>WordInWeb<span className="collab">Collaborative!</span></strong>
            <span className="tag">Editing locally in this browser</span>
          </span>
        </span>
        <nav className="app-links" aria-label="Project links">
          <a className="other" href="https://word-in-web.com" target="_blank" rel="noreferrer">WordInWeb</a>
          <a href="https://github.com/theRealestAEP/wordinweb" target="_blank" rel="noreferrer">GitHub</a>
          <a href="https://www.aepick.me/blog/collaborative-docx-editing/" target="_blank" rel="noreferrer">Blog</a>
          <a href="https://word-in-web.com/report/" target="_blank" rel="noreferrer">Parity report</a>
        </nav>
      </div>
      <div
        data-testid="local-header"
        className="sessionbar"
        style={{ display: "flex", gap: 10, alignItems: "center", padding: "7px 14px", flexWrap: "wrap" }}
      >
        {/* File first, the way every editor puts it first — and every item is
            live here. There is no session yet, so nothing can fork: opening a
            file replaces a document only this browser has ever seen. */}
        <FileMenu
          onNew={newDocument}
          onOpen={openDocument}
          onDownload={downloadDocument}
          onPrint={() => api?.print()}
          disabled={!api}
          // Saved documents (versions, drafts, archived local docs) live in
          // this browser's storage; the slot currently on screen is excluded
          // — "open the document you are already editing" is not an item.
          listSaved={() => store.list().then((all) => all.filter((s) => s.key !== LOCAL_AUTOSAVE_KEY))}
          onOpenSaved={openSaved}
          canRejoinSaved={canRejoinSaved}
          onRejoinSaved={onRejoinSaved}
          onDownloadSaved={downloadSaved}
          onDeleteSaved={(s) => store.delete(s.key)}
          openRequest={savedMenuRequest}
        />
        <button
          data-testid="saved-documents"
          disabled={!api}
          title="Browse documents, versions, and drafts saved in this browser"
          onClick={() => setSavedMenuRequest((n) => n + 1)}
        >
          Saved documents
        </button>
        <span className="bar-sep" aria-hidden="true" />
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
        {persistError && (
          // Same shape as the collab screen's persist banner, same reason: a
          // failed write means the only durable copy stopped updating, and
          // the download button belongs NEXT TO that sentence, not in a menu.
          <span data-testid="local-persist-banner" style={{ fontSize: 12, background: "#f8d7da", padding: "2px 8px", borderRadius: 6 }}>
            {persistError}
            {/* WHAT is taking the space, and the route to act on it. Someone
                stuck in this state needs an action, not a diagnosis: the
                button opens File > Saved, where delete lives (two-step). */}
            {storedSummary && storedSummary.count > 0 && (
              <span data-testid="local-stored-summary">
                {" "}{storedSummary.count} saved cop{storedSummary.count === 1 ? "y uses" : "ies use"} {fmtSize(storedSummary.bytes)} in
                this browser. Delete old copies to free space.
              </span>
            )}{" "}
            <button data-testid="local-persist-download" onClick={downloadDocument} disabled={!api}>
              Download .docx
            </button>{" "}
            <button data-testid="local-manage-saved" onClick={() => setSavedMenuRequest((n) => n + 1)}>
              Manage saved copies
            </button>
          </span>
        )}
        {reclaim.notice && (
          // Startup reclaim's report — stored copies were deleted (or could
          // not be), so it is said on screen and stays until dismissed.
          <span data-testid="reclaim-notice" style={{ fontSize: 12, background: "#fff3cd", padding: "2px 8px", borderRadius: 6 }}>
            {reclaim.notice}{" "}
            <button className="ghost" data-testid="reclaim-dismiss" onClick={reclaim.dismiss}>Dismiss</button>
          </span>
        )}
        <span style={{ flex: 1 }} />
      </div>
      <DocxToolbar api={api} mode="advanced" />
      {/* `position: relative` anchors the loading overlay's inset:0, and
          `minHeight: 0` is what lets this flex child shrink below its content
          so DocxView's container actually scrolls — which is what keeps it
          mounting only the visible pages instead of the whole document. */}
      <div style={{ flex: 1, minHeight: 0, position: "relative" }}>
        {loading && (
          <div className="document-loading" data-testid="doc-loading" role="status" aria-live="polite" aria-busy="true">
            <span className="document-loading-spinner" aria-hidden="true" />
            <strong>Opening {loading}…</strong>
            <span>Large documents may take a few seconds.</span>
          </div>
        )}
        {goPhase && (
          <div className="document-loading" data-testid="golive-progress" role="status" aria-live="polite" aria-busy="true">
            <span className="document-loading-spinner" aria-hidden="true" />
            <strong>Starting collaboration…</strong>
            <span data-testid="golive-phase">{goPhase}</span>
          </div>
        )}
        {/* onError as well as onLoad: a document that fails to parse would
            otherwise leave the spinner up forever, which is a worse lie than
            no spinner. Both paths clear it; the error surfaces through the
            view's own handling. */}
        <DocxView
          source={blank}
          editable
          // height:100% is LOAD-BEARING for virtualization: DocxView's outer
          // wrapper only gets a height when the caller passes one. Without it
          // the wrapper (and the overflow:auto container inside) grow to the
          // full document height, the container never scrolls, and the
          // virtualizer's viewport equals the whole document — every page
          // stays mounted and each keystroke pays O(total pages) in DOM scans
          // (measured: p50 52ms at 500 pages, 10ms once bounded).
          style={{ height: "100%" }}
          onReady={(a) => {
            apiRef.current = a;
            setApi(a);
            // A (re)loaded document counts as unsaved until the next autosave
            // writes it — covers the restored slot, an opened file, and New.
            dirtyRef.current = true;
          }}
          onLoad={() => setLoading(null)}
          onError={() => setLoading(null)}
        />
      </div>

      {modalOpen && (
        <div
          className="modal-scrim"
          data-testid="collab-modal"
          onMouseDown={(e) => { if (e.target === e.currentTarget && !going) setModalOpen(false); }}
        >
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="collab-modal-title">
            <h2 id="collab-modal-title">Share this document</h2>
            <p>
              WordInWeb encrypts this document in your browser. The server only relays encrypted edits.
            </p>
            {goError && (
              <p data-testid="golive-error" style={{ background: "#f8d7da", padding: "6px 10px", borderRadius: 6 }}>
                {goError}
              </p>
            )}
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
              Send this code separately from the link. Everyone needs both to open the document.
            </p>
            <div className="row">
              <button className="ghost" data-testid="collab-cancel" disabled={going} onClick={() => setModalOpen(false)}>
                Cancel
              </button>
              {/* REQUIRED, and the empty field is the whole enforcement: no
                  minimum, no strength rules, no suggestions. Whatever they
                  type is the code. */}
              <button data-testid="start-collab" disabled={going || !code.trim()} onClick={startCollab}>
                {going ? "Starting…" : "Start collaboration"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
