import { useEffect, useRef, useState } from "react";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

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
 * which the collaborative screen has to refuse rather than hide: replacing the
 * document wholesale emits no intent describing the replacement, so the local
 * copy would silently fork from every peer. A missing item teaches nothing
 * about that; a disabled one carrying the reason does.
 */
export function FileMenu({
  onNew,
  onOpen,
  onDownload,
  onPrint,
  disabled,
  openBlockedReason,
}: {
  onNew?: () => void;
  /** Receives the picked file's bytes and its name. */
  onOpen?: (bytes: Uint8Array, filename: string) => void;
  onDownload?: () => void;
  onPrint?: () => void;
  disabled?: boolean;
  /** When set, Open renders DISABLED with this string as the reason. */
  openBlockedReason?: string;
}) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);

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

  // The REASON is the gate, not the missing handler: a screen that passes both
  // gets the refusal, never a live Open. The two are contradictory props and
  // the safe reading of the contradiction is the one that cannot fork a
  // document.
  const blocked = !!openBlockedReason;
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
        File<span aria-hidden="true" className="filemenu-caret">▾</span>
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
                title={openBlockedReason}
                onClick={blocked ? undefined : item(() => fileRef.current?.click())}
              >
                Open .docx…
              </button>
              {blocked && <p className="filemenu-note">{openBlockedReason}</p>}
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
        </div>
      )}
    </div>
  );
}
