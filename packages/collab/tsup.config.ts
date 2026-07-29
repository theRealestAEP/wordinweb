import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/client.ts", "src/server.ts"],
  format: ["esm"],
  // `@wordinweb/core` is private and never published, so it is inlined here —
  // JS via `noExternal`, and .d.ts via the `paths` entry in tsconfig.json.
  // The .d.ts half is the one that bites: the dts rollup resolves separately
  // from the JS bundle and inlines a workspace package only when `paths` maps
  // its EXACT specifier to source (see packages/react/tsup.config.ts for the
  // full story). Pinned by scripts/check-dts-selfcontained.mjs at prepack.
  dts: true,
  sourcemap: true,
  clean: true,
  // Neutral: `./client` runs in browsers, `./server` in node; nothing here
  // touches platform APIs at module scope. Core's own runtime deps (cfb,
  // fflate, wmf; emf-converter is tree-shaken away with the paint path) stay
  // EXTERNAL — inlining core means adopting them, so they are declared as
  // dependencies instead of bundled (cfb is CJS and reaches for `fs`, which
  // does not survive bundling into ESM).
  platform: "neutral",
  noExternal: ["@wordinweb/core"],
});
