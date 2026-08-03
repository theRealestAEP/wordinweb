import { defineConfig } from "vite";
import { fileURLToPath } from "node:url";

/**
 * Dev-only bundler for the anon-share demo. Uses Vite's built-in esbuild JSX
 * (automatic runtime) so no extra React plugin is required. The workspace
 * packages (`wordinweb`, `wordinweb/collab`, `@wordinweb/collab/client`,
 * `@wordinweb/core`) resolve through the repo's node_modules symlinks to their
 * built `dist/` — run `npm run build` at the repository root first. This
 * private example is the local run and test harness.
 */
export default defineConfig({
  root: fileURLToPath(new URL(".", import.meta.url)),
  esbuild: { jsx: "automatic" },
  server: {
    port: 5817,
    open: true,
    proxy: { "/agent-invites": "http://localhost:1234" },
  },
  // Pre-bundle the workspace ESM so their internal bare imports resolve.
  optimizeDeps: { include: ["wordinweb", "wordinweb/collab", "@wordinweb/collab/client", "@wordinweb/core", "react", "react-dom", "react-dom/client"] },
});
