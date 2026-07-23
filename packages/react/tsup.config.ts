import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.tsx", "src/collab.tsx"],
  format: ["esm"],
  dts: true,
  sourcemap: true,
  clean: true,
  platform: "browser",
  // Bundle core into whichever entry uses it; bundle collab only into the
  // collab entry, so the main `wordinweb` bundle never includes the collab
  // engine (plan doc 07 tree-shaking: unreachable beats shakeable).
  noExternal: ["@wordinweb/core", "@wordinweb/collab"],
});
