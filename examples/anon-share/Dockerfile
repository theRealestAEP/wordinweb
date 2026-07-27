# The anon-share demo, built to STATIC FILES for the deployed stack.
#
# Two deliberate properties:
#  - The output is plain files. There is no app server, no SSR, nothing to
#    exploit at runtime — Caddy serves the bundle and proxies the collab
#    server's routes on the same origin (see compose.yml / Caddyfile).
#  - The bundle is built ONCE at image build time against the workspace
#    packages' dist, so a deployed frontend and its server come from the same
#    tree. Version skew between the two is what ENGINE_VERSION exists to
#    refuse; not shipping skew in the first place is better.
FROM node:22-alpine AS build
WORKDIR /app
COPY package.json package-lock.json ./
COPY packages/core/package.json packages/core/
COPY packages/collab/package.json packages/collab/
COPY packages/server/package.json packages/server/
COPY packages/react/package.json packages/react/
COPY examples/anon-share/package.json examples/anon-share/
# The lockfile is DISCARDED here, and only here. It is generated on the
# developer's machine, so it records that platform's optional rollup binary
# and not linux-musl's; the react package's tsup build then dies with "Cannot
# find module @rollup/rollup-linux-arm64-musl" (npm/cli#4828). `npm install`
# alone is not enough — npm keeps the stale optional tree — so the lockfile
# goes too, which is the remedy npm's own error message prescribes.
#
# THE TRADE, stated because it is real: this stage resolves dependencies fresh
# rather than reproducibly. It is acceptable HERE because the stage emits
# static assets that are verified by the e2e board, and unacceptable in the
# SERVER image, which still uses `npm ci` — it compiles with tsc and never
# touches rollup, so it keeps full lockfile pinning.
RUN rm -f package-lock.json && npm install --no-audit --no-fund
COPY tsconfig.base.json ./
COPY packages/ packages/
COPY examples/anon-share/ examples/anon-share/
# The demo imports the workspace packages' BUILT output, so build them first.
RUN npm run build:packages
# Invoked directly, NOT via `npm -w`: the demo is deliberately outside the
# workspaces list (root package.json declares `packages/*` only), so a
# workspace-scoped build fails with "No workspaces found". Verified by running
# it — this Dockerfile originally used the -w form and would not have built.
RUN cd examples/anon-share && npx vite build

# Nothing but the built assets — the runtime image is Caddy's (compose.yml
# mounts this stage's output), so this stage exists to be copied FROM.
FROM scratch AS dist
COPY --from=build /app/examples/anon-share/dist /dist
