#!/usr/bin/env node
/**
 * Fail if any package's `dist/` is older than its `src/`.
 *
 * WHY THIS EXISTS. `packages/server/vitest.config.ts` deliberately has no
 * source aliases, so server tests resolve `@wordinweb/collab` and
 * `@wordinweb/core` to their BUILT output — which is the right call, because it
 * is what actually ships. The cost is that a stale build tests the wrong code
 * SILENTLY, and passing tests are the worst possible disguise for that.
 *
 * It happened three times in one day:
 *  - the envelope-padding suite passed in a worktree and failed after merging,
 *    purely because the merging checkout's collab dist predated the change;
 *  - the remote-repaint benchmark failed on a correct merge because
 *    `packages/react` had not been rebuilt (the ad-hoc build command named only
 *    core and collab, not the `build:packages` script that names all four);
 *  - an agent ran four mutation tests green against a checkout it had not built.
 *
 * The fix is not more discipline. It is making the failure loud, before the
 * tests get a chance to report a comforting lie.
 *
 * Deliberately a timestamp comparison, not a content hash: it needs to be fast
 * enough to sit in front of every test run, and "src touched after dist" is
 * exactly the condition that matters. False positives (a touched-but-unchanged
 * file) cost one rebuild; a false negative costs a wrong green.
 */
import { readdirSync, statSync, existsSync } from "node:fs";
import { join } from "node:path";

const ROOT = new URL("..", import.meta.url).pathname;
const PACKAGES = ["core", "collab", "server", "react"];

/** Newest mtime under `dir`, or 0 when it does not exist. */
function newestMtime(dir) {
  if (!existsSync(dir)) return 0;
  let newest = 0;
  const walk = (d) => {
    for (const entry of readdirSync(d, { withFileTypes: true })) {
      // node_modules under a package would swamp the walk and says nothing
      // about whether our own sources moved.
      if (entry.name === "node_modules") continue;
      const p = join(d, entry.name);
      if (entry.isDirectory()) walk(p);
      else newest = Math.max(newest, statSync(p).mtimeMs);
    }
  };
  walk(dir);
  return newest;
}

const stale = [];
for (const pkg of PACKAGES) {
  const base = join(ROOT, "packages", pkg);
  const src = join(base, "src");
  const dist = join(base, "dist");
  if (!existsSync(src)) continue;
  const srcAt = newestMtime(src);
  // No dist at all is stale only if something consumes it; report it and let
  // the rebuild sort it out rather than guessing which packages are needed.
  const distAt = newestMtime(dist);
  if (srcAt > distAt) {
    stale.push({ pkg, seconds: Math.round((srcAt - distAt) / 1000), built: distAt > 0 });
  }
}

if (stale.length) {
  console.error("\nStale build — tests would run against the WRONG code.\n");
  for (const { pkg, seconds, built } of stale) {
    console.error(
      built
        ? `  packages/${pkg}: src is ${seconds}s newer than dist`
        : `  packages/${pkg}: never built`,
    );
  }
  console.error("\nRebuild everything, then re-run:\n\n  npm run build:packages\n");
  console.error("Naming individual packages is how this gets missed — build:packages");
  console.error("covers core, collab, server and react together.\n");
  process.exit(1);
}
