# Cellestra for a NAS: builds the web client, then runs the dependency-free Node server.
FROM node:22-alpine AS build
WORKDIR /src
RUN corepack enable
COPY pnpm-lock.yaml pnpm-workspace.yaml package.json .npmrc ./
COPY packages/core/package.json packages/core/
COPY packages/charts/package.json packages/charts/
COPY packages/ui/package.json packages/ui/
COPY apps/web/package.json apps/web/
RUN pnpm install --frozen-lockfile --filter @cellestra/web...
COPY tsconfig.base.json ./
COPY packages packages
COPY apps/web apps/web
RUN VITE_CELLESTRA_NAS=1 pnpm --filter @cellestra/web build

FROM node:22-alpine
ENV NODE_ENV=production PORT=8787 CELLESTRA_STATIC=/app/web CELLESTRA_ROOTS="Documents=/data"
WORKDIR /app
COPY apps/server/package.json server/
COPY apps/server/src server/src
COPY --from=build /src/apps/web/dist web
USER node
EXPOSE 8787
VOLUME ["/data"]
HEALTHCHECK --interval=30s --timeout=3s CMD wget -qO- http://127.0.0.1:8787/api/info >/dev/null || exit 1
CMD ["node", "server/src/main.ts"]
