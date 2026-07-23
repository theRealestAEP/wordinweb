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

## Running it (needs a browser + `ws`)

```
npm i ws                       # optional peer dep of @wordinweb/server
npx wordinweb-collab-server    # ws://localhost:1234, ephemeral (not durable)
# then serve this app (Vite) and open two tabs on the same magic link
```

`src/app.tsx` shows the wiring: `useCollab({ url, docId, clientId })` →
`<DocxView collab={session} editable />`. The frontend and the visual editor
integration are browser-verified (the headless test suite covers the
protocol, convergence, and binding logic; the DOM rendering and the
DocxEditor↔EditorBridge adapter run only in a browser).

## Security posture (doc 11)

Uploads and document import are disabled (docs start blank). Rich paste is
disabled (plain text only). Demo runs on an isolated signing key and isolated
capacity, behind the rate limits above, with a strict CSP. Hyperlink
authoring is off until the scheme-allowlist launch gate (already implemented
in core `url-safety.ts`, but the authoring intent is not yet enabled).
