import { defineConfig, devices } from "@playwright/test";

/**
 * End-to-end browser tests for the zero-custody collab demo (plan doc 09
 * H-tier, extended to the browser). These drive the REAL stack — the Vite
 * app over a real WebSocket + HTTP seed against the zero-custody server —
 * so they catch the integration-layer bugs unit/loopback tests miss (the
 * class that broke the demo historically: double-apply, self-transform,
 * dead-typing). Deterministic: fixed ports, `webServer` boots the stack and
 * waits for readiness, no external network.
 *
 * Run: `npm run e2e` (headless) / `npm run e2e:headed` (watch it).
 * Prereq: packages built (`npm run build:packages`) — the webServer command
 * builds them, so a cold `npm run e2e` works from scratch.
 */
export default defineConfig({
  testDir: "./e2e",
  fullyParallel: false, // the demo server is a single shared process
  workers: 1,
  retries: 0,
  timeout: 30_000,
  expect: { timeout: 10_000 },
  reporter: [["list"]],
  use: {
    // Dedicated ports (5173/1234 collide with common dev servers); the
    // demo learns the server port via `?server=` (threaded in the spec).
    baseURL: "http://localhost:5399",
    trace: "retain-on-failure",
    // The demo mints/persists a per-browser clientId in localStorage; a
    // fresh context per test keeps scenarios isolated and deterministic.
  },
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"] } }],
  webServer: {
    command: "node scripts/dev.mjs --no-open",
    url: "http://localhost:5399",
    reuseExistingServer: false, // always our fresh stack, never a stray server
    timeout: 120_000,
    env: { VITE_PORT: "5399", PORT: "1399" },
    stdout: "pipe",
    stderr: "pipe",
  },
});
