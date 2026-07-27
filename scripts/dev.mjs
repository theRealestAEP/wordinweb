#!/usr/bin/env node
/**
 * Single-command local launch for the zero-custody collab demo (plan
 * doc 12). Builds the workspace packages, then runs BOTH halves of the
 * stack — the zero-custody server (HTTP seed + WebSocket on :1234) and the
 * Vite demo app (:5817) — under one process. Ctrl+C stops both.
 *
 *   npm run demo            # build + serve, opens the browser
 *   npm run demo -- --no-build   # skip the build (packages already built)
 *   PORT=… VITE_PORT=… npm run demo
 *
 * Playwright's webServer reuses this via `--no-open` (see playwright.config).
 */
import { spawn } from "node:child_process";
import { once } from "node:events";
import { setTimeout as delay } from "node:timers/promises";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = new Set(process.argv.slice(2));
const SERVER_PORT = process.env.PORT ?? "1234";
const VITE_PORT = process.env.VITE_PORT ?? "5817";
const OPEN = !args.has("--no-open");

const children = [];
function run(cmd, cmdArgs, opts = {}) {
  const child = spawn(cmd, cmdArgs, { cwd: ROOT, stdio: "inherit", ...opts });
  children.push(child);
  return child;
}
function shutdown(code = 0) {
  for (const c of children) c.kill("SIGTERM");
  process.exit(code);
}
for (const sig of ["SIGINT", "SIGTERM"]) process.on(sig, () => shutdown(0));

async function main() {
  if (!args.has("--no-build")) {
    console.log("[dev] building workspace packages…");
    const build = run("npm", ["run", "build:packages"]);
    const [code] = await once(build, "exit");
    if (code !== 0) shutdown(code);
  }

  console.log(`[dev] zero-custody server → http://localhost:${SERVER_PORT}  (obs: GET /stats, structured events on stderr)`);
  // WW_OBS on for the demo: the server prints one structured JSON line per
  // connection lifecycle event (conn-open / refused / kick / conn-close) to
  // stderr, and exposes GET /stats — so a "why aren't two windows syncing?"
  // question is answerable from the terminal (are both connecting to the SAME
  // docId? any refusal? see observability.ts — no doc content is ever logged).
  const server = run("node", ["packages/server/dist/cli.js"], {
    env: { ...process.env, ZERO_CUSTODY: "1", PORT: SERVER_PORT, WW_OBS: "1" },
  });
  server.on("exit", (c) => shutdown(c ?? 0));

  // Give the server its listen tick before Vite proxies/opens.
  await delay(400);

  console.log(`[dev] demo app → http://localhost:${VITE_PORT}`);
  // `--force` re-runs optimizeDeps every boot: the demo consumes the
  // workspace packages' freshly-built dist, and Vite's dep cache would
  // otherwise serve a STALE pre-bundle after a rebuild (a footgun that
  // masked package fixes until the cache was cleared).
  // --strictPort: FAIL LOUDLY if VITE_PORT is taken instead of silently
  // drifting to :5174 — a drift left a stale, still-running Vite serving the
  // OLD pre-bundle while the fresh one sat on another port, so edits looked
  // "not live" because the tab was on the stale bundle. Better to error than
  // to serve stale.
  const viteArgs = ["vite", "--force", "--strictPort", "--port", VITE_PORT, ...(OPEN ? ["--open"] : [])];
  const vite = run("npx", viteArgs, {
    cwd: resolve(ROOT, "examples/anon-share"),
    // VITE_COLLAB_SERVER: tells the app where THIS stack's API lives. The
    // app's production default is SAME-ORIGIN (one Caddy origin), which is
    // wrong under the two-port dev split (Vite :5817 / API :1234) — without
    // this, GET /blank hits Vite, returns index.html, and the local editor
    // shows "invalid zip data". The e2e board never caught that because
    // every spec passes ?server=; the bare-URL dev flow is pinned now.
    env: { ...process.env, VITE_COLLAB_SERVER: `localhost:${SERVER_PORT}` },
  });
  vite.on("exit", (c) => shutdown(c ?? 0));
}

main().catch((e) => {
  console.error(e);
  shutdown(1);
});
