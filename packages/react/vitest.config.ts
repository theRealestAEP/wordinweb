import { defineConfig } from "vitest/config";
import { fileURLToPath } from "node:url";

const src = (p: string) => fileURLToPath(new URL(p, import.meta.url));

export default defineConfig({
  // Resolve the workspace packages to their SOURCE, not their built dist.
  //
  // These tests import @wordinweb/* through node_modules, which resolves to
  // dist — so a suite could pass against a STALE build and "verified" meant
  // "verified against whatever was compiled last". That bit three separate
  // investigations in one session: a fix looked ineffective, a mutation test
  // looked like it proved nothing, and a real regression looked green. The
  // failure is always silent and always points the wrong way.
  //
  // Sibling of the same fix in examples/anon-share/tsconfig.json, which does
  // this for the typecheck layer. Between them, neither layer can be verified
  // against code that is no longer on disk.
  //
  // ORDER MATTERS: the subpath entries must precede their package prefix so
  // "@wordinweb/collab/client" is not swallowed by "@wordinweb/collab".
  resolve: {
    alias: [
      { find: "@wordinweb/collab/client", replacement: src("../collab/src/client.ts") },
      { find: "@wordinweb/collab/server", replacement: src("../collab/src/server.ts") },
      { find: "@wordinweb/core", replacement: src("../core/src/index.ts") },
      { find: "@wordinweb/server", replacement: src("../server/src/index.ts") },
      // The DEMO app (examples/anon-share) imports this package by its public
      // name, which node_modules resolves to dist. Tests that mount demo
      // components need the same source-not-dist guarantee as everything else.
      { find: "wordinweb/collab", replacement: src("./src/collab.tsx") },
      { find: /^wordinweb$/, replacement: src("./src/index.tsx") },
    ],
  },
  test: {
    environment: "jsdom",
    include: ["test/**/*.test.{ts,tsx}"],
    setupFiles: ["./test/setup.ts"],
    testTimeout: 15000,
  },
});
