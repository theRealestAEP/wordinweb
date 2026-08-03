import { readFileSync, writeFileSync } from "node:fs";
import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/cli.ts"],
  format: ["esm"],
  dts: true,
  sourcemap: true,
  clean: true,
  platform: "neutral",
  noExternal: ["@wordinweb/core", "@wordinweb/collab/client", "@wordinweb/collab/server"],
  onSuccess: async () => {
    const cli = "dist/cli.js";
    const source = readFileSync(cli, "utf8");
    if (!source.startsWith("#!")) writeFileSync(cli, `#!/usr/bin/env node\n${source}`);
  },
});
