import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.tsx", "src/collab.tsx"],
  format: ["esm"],
  // `noExternal` bundles the JS, but the .d.ts rollup resolves SEPARATELY,
  // and it inlines a workspace package only when tsconfig `paths` maps that
  // exact specifier to source. `@wordinweb/core` was mapped and inlined;
  // `@wordinweb/collab/client` was not, so the declarations shipped a bare
  // import of a private:true package that is never published — consumers'
  // tsc then either hard-errors (skipLibCheck:false) or, worse because it is
  // silent, degrades every re-exported type to `any`. The fix is the paths
  // entry in tsconfig.json, not a tsup option: `dts.resolve` looks like the
  // knob but tsup turns every `paths` key into an ignore-list for it.
  // Pinned by scripts/check-dts-selfcontained.mjs.
  dts: true,
  sourcemap: true,
  clean: true,
  platform: "browser",
  // Bundle core into whichever entry uses it; bundle collab only into the
  // collab entry, so the main `wordinweb` bundle never includes the collab
  // engine (plan doc 07 tree-shaking: unreachable beats shakeable).
  noExternal: ["@wordinweb/core", "@wordinweb/collab"],
});
