import { defineConfig } from "vitest/config";
import { fileURLToPath } from "node:url";

const source = (path: string) => fileURLToPath(new URL(path, import.meta.url));

export default defineConfig({
  resolve: {
    alias: [
      { find: "@wordinweb/collab/client", replacement: source("../../packages/collab/src/client.ts") },
      { find: "@wordinweb/collab/server", replacement: source("../../packages/collab/src/server.ts") },
      { find: "@wordinweb/core", replacement: source("../../packages/core/src/index.ts") },
      { find: "@wordinweb/server", replacement: source("../../packages/server/src/index.ts") },
      { find: "wordinweb/collab", replacement: source("../../packages/react/src/collab.tsx") },
      { find: /^wordinweb$/, replacement: source("../../packages/react/src/index.tsx") },
    ],
  },
  test: {
    environment: "jsdom",
    include: ["test/**/*.test.{ts,tsx}"],
    setupFiles: ["./test/setup.ts"],
    testTimeout: 15_000,
  },
});
