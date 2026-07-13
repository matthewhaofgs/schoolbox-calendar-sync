# syntax=docker/dockerfile:1

FROM node:22-bookworm-slim AS dependencies
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci

FROM dependencies AS builder
ENV NEXT_TELEMETRY_DISABLED=1
COPY . .
RUN npm run build

FROM node:22-bookworm-slim AS runner
ENV NODE_ENV=production \
    RELAY_HOST=0.0.0.0 \
    PORT=3000 \
    NEXT_TELEMETRY_DISABLED=1 \
    DATABASE_PATH=/app/data/relay.sqlite
WORKDIR /app

RUN groupadd --system --gid 10001 relay \
  && useradd --system --uid 10001 --gid relay --home-dir /app relay

COPY --from=builder --chown=relay:relay /app/.next/standalone ./
COPY --from=builder --chown=relay:relay /app/.next/static ./.next/static
COPY --from=builder --chown=relay:relay /app/scripts ./scripts
RUN mkdir -p /app/data /app/.next/cache && chown -R relay:relay /app/data /app/.next

USER relay
EXPOSE 3000
VOLUME ["/app/data"]
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD ["node", "-e", "fetch('http://127.0.0.1:3000/api/health').then(r=>{if(!r.ok)process.exit(1)}).catch(()=>process.exit(1))"]
CMD ["node", "scripts/production-runner.mjs"]
