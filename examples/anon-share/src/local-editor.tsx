import { useEffect, useState } from "react";
import { DocxView, DocxToolbar, type DocxViewApi } from "wordinweb";

/**
 * The demo's LANDING is the normal single-user editor: a local, editable
 * document with no server and no collab. It seeds from the zero-custody
 * server's blank template (`GET /blank`) purely to have bytes to edit — the
 * document lives entirely in this browser until the user chooses to share it.
 *
 * "Make collaborative" is the one and only go-live path (doc 13 §6): it takes
 * the CURRENT edited bytes (`api.save()`), seals them client-side under a fresh
 * doc key, and PUTs the encrypted genesis checkpoint. Collaborative always
 * means encrypted — there is no plaintext-vs-encrypted choice in the UI; the
 * key rides the `#k=` fragment the caller writes into the URL.
 */
export function LocalEditor({
  httpBase,
  onGoLive,
}: {
  httpBase: string;
  /** Seal + PUT the given local bytes, then switch the app into collab mode. */
  onGoLive: (bytes: Uint8Array, shareCode?: string) => Promise<void>;
}) {
  // Blank template to start editing (bytes only — no session is created until
  // "Make collaborative"). Fetched once; a failure surfaces inline.
  const [blank, setBlank] = useState<Uint8Array | null>(null);
  const [loadError, setLoadError] = useState(false);
  const [api, setApi] = useState<DocxViewApi | null>(null);
  const [going, setGoing] = useState(false);
  const [code, setCode] = useState("");

  useEffect(() => {
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
  }, [httpBase]);

  const goLive = () => {
    if (!api) return;
    setGoing(true);
    // Read the CURRENT edited document straight off the imperative API — the
    // exact bytes we seal are what the collaborative session opens to.
    const bytes = api.save();
    void onGoLive(bytes, code.trim() || undefined).catch(() => setGoing(false));
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
        <b>wordinweb</b>
        <span className="hint">Editing locally — nothing leaves your browser until you share it.</span>
        <span style={{ flex: 1 }} />
        <input
          data-testid="share-code"
          value={code}
          placeholder="share code (optional)"
          maxLength={12}
          title="Optional code told to people separately — a leaked link alone can't enter or decrypt"
          onChange={(e) => setCode(e.target.value)}
          style={{ font: "12px system-ui", padding: "4px 8px", border: "1px solid #dadce0", borderRadius: 6 }}
        />
        <button
          data-testid="make-collaborative"
          disabled={going || !api}
          onClick={goLive}
          style={{ font: "13px system-ui", padding: "5px 12px", border: "1px solid #dadce0", background: "#1a73e8", color: "#fff", borderRadius: 6, cursor: "pointer" }}
        >
          {going ? "Going live…" : "Make collaborative"}
        </button>
      </div>
      <DocxToolbar api={api} mode="advanced" />
      <div style={{ flex: 1, minHeight: 0 }}>
        <DocxView source={blank} editable onReady={setApi} />
      </div>
    </div>
  );
}
