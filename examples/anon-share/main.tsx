import { useState } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./src/app";
import { LocalEditor } from "./src/local-editor";
import { goLiveEncrypted } from "./src/e2ee-flows";
import { docKeyFromFragment } from "wordinweb/collab";

/**
 * Dev harness for the zero-custody demo (Vite).
 *
 * The demo OPENS in the normal single-user editor (a local, editable document —
 * no server, no collab). "Make collaborative" seals the CURRENT local document
 * and goes live encrypted (doc 13 §6): the browser mints id + key, seals the
 * bytes client-side, PUTs the genesis checkpoint, and the URL becomes the share
 * link (`?doc=<id>#k=<key>`). The key rides the FRAGMENT — shareable, never
 * sent. A URL that already carries `?doc=` is the "someone shared me a link"
 * path: join that collaborative session directly.
 *
 *   ZERO_CUSTODY=1 node packages/server/dist/cli.js   (:1234, HTTP+WS)
 *
 * When the session ends the server forgets it; this browser's IndexedDB bundle
 * revives it. Override the server with ?server=host:port if not on
 * localhost:1234.
 */
const params = new URLSearchParams(location.search);
const HOSTPORT = params.get("server") ?? "localhost:1234";
const WS = `ws://${HOSTPORT}`;
const HTTP = `http://${HOSTPORT}`;

/**
 * Identity (doc 12 §7): ONE persistent clientId per browser profile — the
 * hub enforces a single live connection per (doc, clientId), so a second
 * same-profile tab gets the "use here instead" takeover flow instead of
 * silently colliding clientSeq counters. Incognito/other browsers mint
 * their own id and join as ordinary extra participants.
 */
function persistent(key: string, mint: () => string): string {
  let v = localStorage.getItem(key);
  if (!v) {
    v = mint();
    localStorage.setItem(key, v);
  }
  return v;
}
const clientId = persistent("wordinweb-client-id", () => {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  return `c_${[...b].map((x) => x.toString(16).padStart(2, "0")).join("")}`;
});
// Display name (doc 14 §2): self-asserted, stable per browser, server-sanitized.
const name = persistent("wordinweb-name", () => "");

function Harness() {
  const [docId, setDocId] = useState<string | null>(() => params.get("doc"));
  // E2EE (doc 13 §1): the key rides the fragment — browsers never send it.
  const [docKey, setDocKey] = useState<string | null>(() => docKeyFromFragment(location.hash));

  // Landing = the normal single-user editor (local, no server, no collab).
  // "Make collaborative" seals the CURRENT local bytes and goes live.
  if (!docId) {
    const goLive = async (bytes: Uint8Array, shareCode?: string) => {
      const { docId: id, docKey: key, ownerToken } = await goLiveEncrypted(HTTP, bytes, shareCode);
      // Owner capability (doc 14 §2.5): kept out of the shared link, in this
      // browser's storage — same key the collab App reads back.
      if (ownerToken) localStorage.setItem(`wordinweb-owner-${id}`, ownerToken);
      const url = new URL(location.href);
      url.searchParams.set("doc", id);
      url.hash = `k=${key}`;
      history.replaceState(null, "", url.toString());
      // Switch into collaborative mode: the session already holds our sealed
      // genesis checkpoint, so App's join opens to exactly this document.
      setDocKey(key);
      setDocId(id);
    };
    return <LocalEditor httpBase={HTTP} onGoLive={goLive} />;
  }

  // Collaborative mode: either we just went live, or the URL carried `?doc=`
  // (someone shared the link). App joins the live session over WS.
  const shareUrl = location.href;
  return (
    <div>
      <header>
        <b>wordinweb collab</b>
        <span className="pill">doc {docId.slice(0, 10)}…</span>
        {docKey ? <span className="pill" style={{ background: "#1a7f37", color: "#fff" }}>encrypted</span> : null}
        <span className="hint">Share this URL / open on another device:</span>
        <input readOnly value={shareUrl} onFocus={(e) => e.currentTarget.select()} />
      </header>
      <div id="root-editor">
        <App
          url={WS}
          httpBase={HTTP}
          docId={docId}
          clientId={clientId}
          name={name}
          docKey={docKey ?? undefined}
          ownerToken={localStorage.getItem(`wordinweb-owner-${docId}`) ?? undefined}
        />
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")!).render(<Harness />);
