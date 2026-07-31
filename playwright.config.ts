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
    env: {
      VITE_PORT: "5399",
      PORT: "1399",
      // PER-IP LIMITS, RAISED FOR THE BOARD — not disabled.
      //
      // The board is the party with the strange access pattern: it seeds a
      // fresh room per scenario (24+ rooms in ~2 minutes, all from
      // 127.0.0.1), which is a machine pattern no human produces. Production
      // defaults (10 seeds/min, 25 live docs, 50 conns per address) stay
      // meaningful precisely because the harness declares its own exemption
      // here rather than weakening them for everyone.
      //
      // NON-ZERO ON PURPOSE: 0 DISABLES a limit, which would skip the guard's
      // code path entirely and cost us the only end-to-end exercise it gets.
      // Large-but-live keeps the checks running (a throw or a leak in the
      // guard still fails the board) while staying out of the way.
      //
      // Rooms linger for WW_EMPTY_ROOM_TTL_MS (60s) after their last
      // participant leaves, so a run's rooms ACCUMULATE against the live-doc
      // cap even as specs finish — which is why maxDocs is raised too, not
      // just the seed rate.
      //
      // Loopback is deliberately NOT exempted in the server itself: on a host
      // running a proxy without WW_TRUST_PROXY set, every real client already
      // appears as 127.0.0.1, so a blanket loopback exemption would silently
      // switch off per-IP protection in exactly the misconfiguration the
      // trust-proxy default is designed to fail safe against.
      WW_IP_SEED_PER_MIN: "500",
      WW_IP_MAX_DOCS: "500",
      WW_IP_MAX_CONNS: "500",
    },
    stdout: "pipe",
    stderr: "pipe",
  },
});
