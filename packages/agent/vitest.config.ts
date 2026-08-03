import { defineConfig } from "vitest/config";
import { fileURLToPath } from "node:url";

export default defineConfig({
  resolve: {
    alias: {
      "@wordinweb/core": fileURLToPath(new URL("../core/src/index.ts", import.meta.url)),
      "@wordinweb/collab/client": fileURLToPath(new URL("../collab/src/client.ts", import.meta.url)),
      "@wordinweb/collab/server": fileURLToPath(new URL("../collab/src/server.ts", import.meta.url)),
    },
  },
  test: {
    environment: "node",
  },
});
