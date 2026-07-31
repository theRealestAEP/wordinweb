# wordinweb collab server — the runnable sequencer host.
# Tag policy (plan doc 07): the image tag MUST equal the npm version so a
# deployed server and its client bundle stay version-locked (the handshake
# refuses mismatched clients).
FROM node:22-alpine AS build
WORKDIR /app
# Build the workspace packages the server depends on.
COPY package.json package-lock.json ./
COPY packages/core/package.json packages/core/
COPY packages/collab/package.json packages/collab/
COPY packages/server/package.json packages/server/
RUN npm ci
# ROLLUP'S NATIVE BINARY, installed explicitly because npm will not.
#
# The packages build with tsup, which loads rollup at startup, and rollup picks
# a native binary through platform-specific OPTIONAL dependencies. A lockfile
# records only the ones resolved on the machine that generated it — ours holds
# @rollup/rollup-darwin-arm64 and nothing else — so inside this Alpine image
# the build died on "Cannot find module @rollup/rollup-linux-arm64-musl".
#
# Two things that do NOT fix it, both tried:
#  - `npm install` instead of `npm ci`: npm does not add the missing platform
#    package while a lockfile is present. That IS the bug (npm/cli#4828).
#  - skipping the type build: tsup loads rollup whether or not it emits .d.ts.
#
# So install it directly. The arch is derived rather than hardcoded, because
# an x86 host would need a different package and a wrong pin would fail here
# with the same confusing message. `--no-save` keeps the lockfile untouched.
RUN ARCH="$(uname -m | sed 's/aarch64/arm64/; s/x86_64/x64/')" && \
    npm i --no-save --no-audit --no-fund "@rollup/rollup-linux-${ARCH}-musl"
COPY tsconfig.base.json ./
COPY packages/core packages/core
COPY packages/collab packages/collab
COPY packages/server packages/server
# Each package's OWN build script, not a hand-rolled tsc invocation. The
# previous form (`npx -w @wordinweb/collab tsc -p packages/collab/tsconfig.json`)
# ran with the workspace as cwd, so the repo-relative config path resolved to
# packages/collab/packages/collab/tsconfig.json and the build died with TS5058.
# It had rotted unnoticed because nothing built this image; `npm run build -w`
# is the same path `build:packages` uses locally, so the two cannot drift.
RUN npm run build -w @wordinweb/core -w @wordinweb/collab -w @wordinweb/server

FROM node:22-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
# Install the server package's declared production dependency tree.
COPY package.json package-lock.json ./
COPY packages/core/package.json packages/core/
COPY packages/collab/package.json packages/collab/
COPY packages/server/package.json packages/server/
# Install only the server workspace's production dependencies.
RUN npm ci --omit=dev --ignore-scripts --workspace @wordinweb/server --include-workspace-root=false
COPY --from=build /app/packages/core/dist /app/packages/core/dist
COPY --from=build /app/packages/collab/dist /app/packages/collab/dist
COPY --from=build /app/packages/server/dist /app/packages/server/dist
ENV PORT=1234
EXPOSE 1234
# The zero-config dev/demo server (blank docs, in-memory ephemeral storage).
# A production deployment instead embeds CollabHub with real storage/JWT/media
# and its own entrypoint.
CMD ["node", "/app/packages/server/dist/cli.js"]
