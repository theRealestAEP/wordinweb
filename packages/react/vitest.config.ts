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
  // ORDER MATTERS: the subpath entries must precede their package prefix so
  // "@wordinweb/collab/client" is not swallowed by "@wordinweb/collab".
  resolve: {
    alias: [
      { find: "@wordinweb/collab/client", replacement: src("../collab/src/client.ts") },
      { find: "@wordinweb/collab/server", replacement: src("../collab/src/server.ts") },
      { find: "@wordinweb/core", replacement: src("../core/src/index.ts") },
      { find: "@wordinweb/server", replacement: src("../server/src/index.ts") },
    ],
  },
  test: {
    environment: "jsdom",
    include: ["test/**/*.test.{ts,tsx}"],
    setupFiles: ["./test/setup.ts"],
    testTimeout: 15000,
  },
});
