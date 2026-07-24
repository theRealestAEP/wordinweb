#!/usr/bin/env node
/**
 * Single-command local launch for the zero-custody collab demo (plan
 * doc 12). Builds the workspace packages, then runs BOTH halves of the
 * stack — the zero-custody server (HTTP seed + WebSocket on :1234) and the
 * Vite demo app (:5173) — under one process. Ctrl+C stops both.
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
const VITE_PORT = process.env.VITE_PORT ?? "5173";
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

  console.log(`[dev] zero-custody server → http://localhost:${SERVER_PORT}`);
  const server = run("node", ["packages/server/dist/cli.js"], {
    env: { ...process.env, ZERO_CUSTODY: "1", PORT: SERVER_PORT },
  });
  server.on("exit", (c) => shutdown(c ?? 0));

  // Give the server its listen tick before Vite proxies/opens.
  await delay(400);

  console.log(`[dev] demo app → http://localhost:${VITE_PORT}`);
  // `--force` re-runs optimizeDeps every boot: the demo consumes the
  // workspace packages' freshly-built dist, and Vite's dep cache would
  // otherwise serve a STALE pre-bundle after a rebuild (a footgun that
  // masked package fixes until the cache was cleared).
  const viteArgs = ["vite", "--force", "--port", VITE_PORT, ...(OPEN ? ["--open"] : [])];
  const vite = run("npx", viteArgs, {
    cwd: resolve(ROOT, "examples/anon-share"),
  });
  vite.on("exit", (c) => shutdown(c ?? 0));
}

main().catch((e) => {
  console.error(e);
  shutdown(1);
});
