import { useState } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./src/app";
import { LocalEditor } from "./src/local-editor";
import { goLiveEncrypted } from "./src/e2ee-flows";
import { docKeyFromFragment, IndexedDbBundleStore } from "wordinweb/collab";
import { installWsRegistry } from "./src/perf/ws-tap";

// Perf HUD plumbing (Ctrl+Shift+P / ?perf=1): registers collab sockets so the
// HUD can attach when opened. Must run BEFORE the session builds its socket;
// registration alone adds no listeners and no per-frame cost.
installWsRegistry();

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
 * revives it. Override the server with ?server=host:port (dev and e2e do).
 */
const params = new URLSearchParams(location.search);
/**
 * Where the collab server lives.
 *
 * DEFAULT: the SAME ORIGIN this page was served from. In the deployed shape a
 * reverse proxy serves this bundle at / and proxies the server's routes
 * (/docs*, /blank, /healthz, and the WebSocket upgrade) on that same origin —
 * so there is no cross-origin request to configure and no CORS story at all.
 * It also follows the page's scheme: an https page speaks wss, which matters
 * more than convenience here, because the E2EE flows need WebCrypto and
 * `crypto.subtle` is unavailable outside a SECURE CONTEXT. A page served over
 * plain http on anything but localhost cannot encrypt, so TLS is a
 * requirement of the feature, not a deployment nicety.
 *
 * OVERRIDE: `?server=host:port` keeps the old behavior verbatim for local dev
 * and the e2e suite, which serve the app from Vite on one port and the server
 * on another. Explicit override always wins and stays plain ws/http, since
 * those runs are localhost (a secure context by definition).
 */
// Resolution order: explicit ?server= (dev/e2e, always wins) → the dev
// stack's injected env (scripts/dev.mjs sets VITE_COLLAB_SERVER to its API
// port — the two-port dev split makes same-origin WRONG there: GET /blank
// would hit Vite and return index.html, "invalid zip data") → same-origin
// (production: one Caddy origin serving app + proxying the API).
// import.meta.env.VITE_COLLAB_SERVER is compile-time: absent from a
// production `vite build` unless deliberately set, so the prod default
// stays same-origin.
const override = params.get("server") || (import.meta.env.VITE_COLLAB_SERVER as string | undefined);
const secure = location.protocol === "https:";
const WS = override ? `ws://${override}` : `${secure ? "wss" : "ws"}://${location.host}`;
const HTTP = override ? `http://${override}` : `${secure ? "https" : "http"}://${location.host}`;

/**
 * Storage helper — `store` is localStorage (per profile, shared across tabs)
 * or sessionStorage (per tab, survives reload but NOT a new tab).
 */
function persist(store: Storage, key: string, mint: () => string): string {
  let v = store.getItem(key);
  if (!v) {
    v = mint();
    store.setItem(key, v);
  }
  return v;
}

/**
 * Identity (doc 12 §7): the hub enforces a single live connection per
 * (doc, clientId) so two sockets can never share clientSeq counters and
 * poison dedup. The clientId therefore lives in **sessionStorage**, which is
 * per-tab: it survives a reload (so resume/reconnect keeps the same identity)
 * but a *second tab* mints its own id and joins as an ordinary extra
 * participant. Using localStorage here was a footgun — every same-profile tab
 * shared one clientId, so opening the doc in a second tab just took over the
 * first (the natural way to try the demo silently broke it). Different
 * browsers/profiles/incognito already mint their own id, unchanged.
 */
const clientId = persist(sessionStorage, "wordinweb-client-id", () => {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  return `c_${[...b].map((x) => x.toString(16).padStart(2, "0")).join("")}`;
});
// Display name (doc 14 §2): self-asserted, stable per browser (localStorage,
// shared across tabs is fine — it's just a label), server-sanitized.
const storedName = persist(localStorage, "wordinweb-name", () => "");

// ONE store for both screens: local autosaves, collab bundles, versions and
// drafts all live in the same IndexedDB, so the File menu can list them all
// and switching between local and collaborative reuses one connection.
const bundleStore = new IndexedDbBundleStore();

function Harness() {
  const [docId, setDocId] = useState<string | null>(() => params.get("doc"));
  const [name, setName] = useState(storedName);
  // Bytes carried back out of a session by Disconnect, so the local editor
  // reopens the document you were editing rather than a blank one.
  const [localBytes, setLocalBytes] = useState<Uint8Array | undefined>(undefined);
  const [copied, setCopied] = useState(false);
  // E2EE (doc 13 §1): the key rides the fragment — browsers never send it.
  const [docKey, setDocKey] = useState<string | null>(() => docKeyFromFragment(location.hash));

  // Abandon the current collaborative doc and return to the local editor to
  // start a fresh one. The escape hatch for "I forgot the share code" — you
  // can't get into the old (code-locked) session, but you can always make a
  // new one. Clears ?doc= and the #k= key while keeping ?server=.
  const newDocument = () => {
    const u = new URL(location.href);
    u.searchParams.delete("doc");
    u.hash = "";
    history.replaceState(null, "", u.toString());
    setDocKey(null);
    setDocId(null);
  };

  // Landing = the normal single-user editor (local, no server, no collab).
  // "Make collaborative" seals the CURRENT local bytes and goes live.
  if (!docId) {
    const goLive = async (bytes: Uint8Array, shareCode?: string, onPhase?: (phase: "encrypt" | "upload") => void) => {
      const { docId: id, docKey: key, ownerToken } = await goLiveEncrypted(HTTP, bytes, shareCode, onPhase);
      // Owner capability (doc 14 §2.5): kept out of the shared link, in this
      // browser's storage — same key the collab App reads back.
      if (ownerToken) localStorage.setItem(`wordinweb-owner-${id}`, ownerToken);
      // Persist the creator's OWN share code on THIS device (never in the link)
      // so they aren't re-prompted for a code they just set, and a reload can't
      // lock them out of their own document. Joiners on other devices still
      // supply it out-of-band (the code is a second factor beyond the link).
      if (shareCode) localStorage.setItem(`wordinweb-code-${id}`, shareCode);
      const url = new URL(location.href);
      url.searchParams.set("doc", id);
      url.hash = `k=${key}`;
      history.replaceState(null, "", url.toString());
      // Switch into collaborative mode: the session already holds our sealed
      // genesis checkpoint, so App's join opens to exactly this document.
      setDocKey(key);
      setDocId(id);
    };
    return <LocalEditor httpBase={HTTP} initialBytes={localBytes} onGoLive={goLive} store={bundleStore} />;
  }

  // Collaborative mode: either we just went live, or the URL carried `?doc=`
  // (someone shared the link). App joins the live session over WS.
  //
  // THE SHARE LINK IS NEVER RENDERED. It carries the decryption key in its
  // fragment, so putting it on screen puts it in every screenshot, screen
  // share, and shoulder-glance. Copying it to the clipboard is the whole
  // interaction — the string only ever exists in the DOM for as long as
  // navigator.clipboard needs it.
  const copyShareLink = () => {
    void navigator.clipboard.writeText(location.href).then(
      () => {
        setCopied(true);
        setTimeout(() => setCopied(false), 1600);
      },
      () => setCopied(false),
    );
  };

  // Leave the room, keep the copy. The session's current bytes come back with
  // us into the local editor and the socket is closed, so the server sees a
  // departure rather than a phantom participant. Nothing durable is lost —
  // that is the zero-custody contract, made into a button.
  const disconnect = (bytes: Uint8Array | null) => {
    if (bytes) setLocalBytes(bytes);
    const u = new URL(location.href);
    u.searchParams.delete("doc");
    u.hash = "";
    history.replaceState(null, "", u.toString());
    setDocKey(null);
    setDocId(null);
  };

  return (
    // THE HEIGHT CHAIN IS LOAD-BEARING (see index.html): the local branch
    // mounts directly under the height:100% #root, but this collab branch
    // wraps App in two divs — without heights here, App's `height: 100%`
    // resolves against auto, the DocxView container never scrolls, and the
    // virtualizer mounts EVERY page (p50 52 ms/keystroke at 500 pages).
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div className="app-header">
        <span className="brand-lockup">
          <span className="brand-mark">W</span>
          <span className="brand-copy">
            <strong>WordInWeb<span className="collab">Collaborative!</span></strong>
            <span className="tag">End-to-end encrypted · non-custodial · nothing stored on the server</span>
          </span>
        </span>
        <nav className="app-links" aria-label="Project links">
          <a className="other" href="https://word-in-web.com" target="_blank" rel="noreferrer">WordInWeb</a>
          <a href="https://github.com/theRealestAEP/wordinweb" target="_blank" rel="noreferrer">GitHub</a>
          <a href="https://www.aepick.me/blog" target="_blank" rel="noreferrer">Blog</a>
          <a href="https://word-in-web.com/report/" target="_blank" rel="noreferrer">Parity report</a>
        </nav>
      </div>
      {/* The alias/share/new-document controls used to live in their own bar
          here. They are session controls like download and disconnect, so they
          now sit in App's single session row instead of a second strip of
          chrome saying much the same thing one line higher. */}
      <div id="root-editor" style={{ flex: 1, minHeight: 0 }}>
        <App
          url={WS}
          httpBase={HTTP}
          docId={docId}
          clientId={clientId}
          name={name}
          store={bundleStore}
          onDisconnect={disconnect}
          docKey={docKey ?? undefined}
          ownerToken={localStorage.getItem(`wordinweb-owner-${docId}`) ?? undefined}
          // Creator's own code (this device) — seeds App so it joins without a
          // re-prompt; also carries a joiner's entered code across reloads.
          initialShareCode={localStorage.getItem(`wordinweb-code-${docId}`) ?? undefined}
          onNewDocument={newDocument}
          onNameChange={(v) => { setName(v); localStorage.setItem("wordinweb-name", v); }}
          onShareLink={copyShareLink}
          shareCopied={copied}
        />
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")!).render(<Harness />);
