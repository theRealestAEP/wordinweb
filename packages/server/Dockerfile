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
COPY tsconfig.base.json ./
COPY packages/core packages/core
COPY packages/collab packages/collab
COPY packages/server packages/server
RUN npm run build -w @wordinweb/core \
 && npx -w @wordinweb/collab tsc -p packages/collab/tsconfig.json \
 && npx -w @wordinweb/server tsc -p packages/server/tsconfig.json

FROM node:22-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
# ws is an optional peer dep of @wordinweb/server; the dev/demo entrypoint needs it.
RUN npm init -y >/dev/null 2>&1 && npm i ws@^8 >/dev/null 2>&1
COPY --from=build /app/packages/core/dist /app/packages/core/dist
COPY --from=build /app/packages/core/package.json /app/packages/core/
COPY --from=build /app/packages/collab/dist /app/packages/collab/dist
COPY --from=build /app/packages/collab/package.json /app/packages/collab/
COPY --from=build /app/packages/server/dist /app/packages/server/dist
COPY --from=build /app/packages/server/package.json /app/packages/server/
ENV PORT=1234
EXPOSE 1234
# The zero-config dev/demo server (blank docs, in-memory ephemeral storage).
# A production deployment instead embeds CollabHub with real storage/JWT/media
# and its own entrypoint.
CMD ["node", "/app/packages/server/dist/cli.js"]
