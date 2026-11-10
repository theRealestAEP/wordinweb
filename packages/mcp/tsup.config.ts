import { readFileSync, writeFileSync } from "node:fs";
import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/server.ts", "src/cli.ts"],
  format: ["esm"],
  dts: true,
  sourcemap: true,
  clean: true,
  platform: "node",
  target: "node20",
  onSuccess: async () => {
    const cli = "dist/cli.js";
    const source = readFileSync(cli, "utf8");
    if (!source.startsWith("#!")) writeFileSync(cli, `#!/usr/bin/env node\n${source}`);
  },
});
