import { readFileSync, writeFileSync } from "node:fs";
import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/cli.ts"],
  format: ["esm"],
  // `@wordinweb/core` and `@wordinweb/collab` are private and never published,
  // so both are inlined — JS via `noExternal`, .d.ts via `paths` in
  // tsconfig.json (see packages/react/tsup.config.ts for why both halves are
  // needed). Core's runtime deps (cfb, fflate, wmf; emf-converter is
  // tree-shaken away with the paint path) stay external and are declared as
  // dependencies — inlining core means adopting them. `ws` is an optional
  // peer, only ever loaded through a dynamic `import(spec)` with a variable
  // specifier, which esbuild leaves untouched.
  // Pinned by scripts/check-dts-selfcontained.mjs at prepack.
  dts: true,
  sourcemap: true,
  clean: true,
  platform: "node",
  noExternal: ["@wordinweb/core", "@wordinweb/collab"],
  // The npm `bin` needs a shebang, but it cannot live in src/cli.ts: vitest's
  // rollup transform fails on `#!` in an imported module, and react's tests
  // reach cli.ts through index.ts re-exports. So prepend it to the built file.
  onSuccess: async () => {
    const cli = "dist/cli.js";
    const src = readFileSync(cli, "utf8");
    if (!src.startsWith("#!")) writeFileSync(cli, "#!/usr/bin/env node\n" + src);
  },
});
