# wordinweb collab demo — anon-share

The public showcase from the plan (`internal/collab-plan/11-threat-model.md`),
two modes:

- **Magic link** — click *New document* → get an unguessable URL → share it
  (Hacker-News-post style) → a small group edits with the full feature set.
  The URL is the capability (an unguessable ≥128-bit id).
- **Party doc** — join a rotating shared document with a crowd of strangers,
  text + presence only. Deliberately minimal surface, so this
  highest-exposure anonymous mode is **structurally XSS-free** (no
  authored-URL, media, or import sink is enabled).

## What's implemented and tested

The server-side demo logic lives in `@wordinweb/server` (`src/demo.ts`) and is
unit-tested (`test/demo.test.ts`):

- `makeDocId(randomBytes)` — unguessable 128-bit doc ids (the magic-link
  capability).
- `PartyPool` — round-robin assignment of visitors onto a fixed set of shared
  docs (bounded storage/abuse surface).
- `DEMO_INTENT_ALLOWLIST` / `intentAllowedInDemo(mode, kind)` — the explicit
  enabled-intent allowlist per mode (party = text-only; magic-link = the full
  implemented set; **no authored-URL intents** in either until the
  scheme-allowlist gate).
- `RateLimiter` — token-bucket abuse limits (per-IP doc creation,
  per-connection intents).

The collab loop these drive (`@wordinweb/collab` + `@wordinweb/server`) is
tested end to end headlessly (multi-client convergence, presence,
persistence). The client binding (`bindEditor`, `useCollab`) is tested via the
in-process loopback.

## Running it (two terminals)

From the **repo root**, build the packages once (and re-run after any package
change):

```
npm run build            # builds core + wordinweb (react)
# collab/server are built via their own tsconfig if needed:
( cd packages/collab && npm run build ) && ( cd packages/server && npm run build )
```

Then, in this directory (`examples/anon-share/`):

```
# terminal 1 — the collab dev server (ws://localhost:1234, auth-off, ephemeral)
npm run server

# terminal 2 — the Vite app (opens a browser tab)
npm run dev
```

Vite prints a `http://localhost:5817/` (or next free port) URL. Open it, click
**New document** — the page adopts an unguessable `?doc=<id>` URL (the
magic-link capability). **Copy that URL into a second browser tab** (or send it
to a friend) and type in either: edits and cursors sync live between tabs.

Override the server with `?server=ws://host:port` if you started it elsewhere
(e.g. `PORT=3000 npm run server`).

`src/app.tsx` shows the wiring: `<CollabEditor url docId clientId editable />`
composes `useCollab({ url, docId, clientId })` → the live `DocxView`. The
harness (`main.tsx` + `index.html` + `vite.config.ts`) is dev-only and not
shipped to npm. The headless test suite covers the protocol, convergence, and
binding logic; the DOM rendering runs here in the browser.

## Security posture (doc 11)

Uploads and document import are disabled (docs start blank). Rich paste is
disabled (plain text only). Demo runs on an isolated signing key and isolated
capacity, behind the rate limits above, with a strict CSP. Hyperlink
authoring is off until the scheme-allowlist launch gate (already implemented
in core `url-safety.ts`, but the authoring intent is not yet enabled).
