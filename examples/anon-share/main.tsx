import { createElement, useState } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./src/app";

/**
 * Dev harness for the zero-custody demo (Vite). The browser URL carries the
 * magic-link capability (`?doc=<id>`, plan doc 11); the server is the
 * zero-custody one:
 *
 *   ZERO_CUSTODY=1 node packages/server/dist/cli.js   (:1234, HTTP+WS)
 *
 * "New document" is go-live (doc 12 §3): POST /docs {blank:true} seeds an
 * in-RAM session and the URL becomes the share link. When the session ends
 * the server forgets it; this browser's IndexedDB bundle revives it.
 * Override the server with ?server=host:port if not on localhost:1234.
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
  const [creating, setCreating] = useState(false);

  if (!docId) {
    return createElement(
      "div",
      { style: { maxWidth: 560, margin: "12vh auto", background: "#fff", padding: 28, borderRadius: 12, border: "1px solid #dadce0" } },
      createElement("h2", { style: { marginTop: 0 } }, "wordinweb — zero-custody collab demo"),
      createElement(
        "p",
        { className: "hint" },
        "Create a document and share its URL. The server hosts the live session only — ",
        "your browser keeps the document, and any participant can bring the same link back to life later.",
      ),
      createElement(
        "button",
        {
          disabled: creating,
          onClick: () => {
            setCreating(true);
            // Go-live (doc 12 §3): the server mints the unguessable docId —
            // the URL IS the capability (doc 11).
            void fetch(`${HTTP}/docs`, {
              method: "POST",
              headers: { "content-type": "application/json" },
              body: JSON.stringify({ blank: true }),
            })
              .then((r) => r.json())
              .then(({ docId: id }: { docId: string }) => {
                const url = new URL(location.href);
                url.searchParams.set("doc", id);
                history.replaceState(null, "", url.toString());
                setDocId(id);
              })
              .catch(() => setCreating(false));
          },
        },
        creating ? "Creating…" : "New document",
      ),
      createElement("p", { className: "hint", style: { marginTop: 20 } }, `Server: ${HTTP} (zero-custody — sessions are not persisted)`),
    );
  }

  const shareUrl = location.href;
  return createElement(
    "div",
    null,
    createElement(
      "header",
      null,
      createElement("b", null, "wordinweb collab"),
      createElement("span", { className: "pill" }, `doc ${docId.slice(0, 10)}…`),
      createElement("span", { className: "hint" }, "Share this URL / open on another device:"),
      createElement("input", { readOnly: true, value: shareUrl, onFocus: (e: { target: { select(): void } }) => e.target.select() }),
    ),
    createElement(
      "div",
      { id: "root-editor" },
      createElement(App, { url: WS, httpBase: HTTP, docId, clientId, name }),
    ),
  );
}

createRoot(document.getElementById("root")!).render(createElement(Harness));
