import { createElement, useState } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./src/app";

/**
 * Dev harness for the anon-share demo (Vite). Wires the browser URL to the
 * magic-link model from plan doc 11: the `?doc=<id>` query param IS the
 * capability. No id in the URL → mint an unguessable one and adopt it (a
 * "New document"); share the URL and open it in a second tab to collaborate.
 *
 * The collab server is the zero-config dev server:
 *   node packages/server/dist/cli.js        (ws://localhost:1234, auth-off)
 * Override with ?server=ws://host:port if you ran it on another port.
 */
const params = new URLSearchParams(location.search);
const SERVER = params.get("server") ?? "ws://localhost:1234";

function mintDocId(): string {
  // 128-bit unguessable id — the magic-link capability (browser crypto).
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  return [...b].map((x) => x.toString(16).padStart(2, "0")).join("");
}
// A per-tab client id so two tabs are two distinct participants.
const clientId = `tab-${mintDocId().slice(0, 8)}`;

function Harness() {
  const [docId, setDocId] = useState<string | null>(() => params.get("doc"));

  if (!docId) {
    // Landing: create a magic-link document.
    return createElement(
      "div",
      { style: { maxWidth: 560, margin: "12vh auto", background: "#fff", padding: 28, borderRadius: 12, border: "1px solid #dadce0" } },
      createElement("h2", { style: { marginTop: 0 } }, "wordinweb collaborative demo"),
      createElement("p", { className: "hint" }, "Create a document, then share its URL (or just open it in a second tab) to edit together in real time."),
      createElement(
        "button",
        {
          onClick: () => {
            const id = mintDocId();
            const url = new URL(location.href);
            url.searchParams.set("doc", id);
            history.replaceState(null, "", url.toString());
            setDocId(id);
          },
        },
        "New document",
      ),
      createElement("p", { className: "hint", style: { marginTop: 20 } }, `Server: ${SERVER}`),
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
      createElement("span", { className: "pill" }, `doc ${docId.slice(0, 8)}…`),
      createElement("span", { className: "pill" }, clientId),
      createElement("span", { className: "hint" }, "Share this URL / open in another tab:"),
      createElement("input", { readOnly: true, value: shareUrl, onFocus: (e: any) => e.target.select() }),
      createElement("span", { className: "hint" }, `· ${SERVER}`),
    ),
    createElement(
      "div",
      { id: "root-editor" },
      createElement(App, { url: SERVER, docId, clientId }),
    ),
  );
}

createRoot(document.getElementById("root")!).render(createElement(Harness));
