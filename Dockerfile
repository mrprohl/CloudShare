# --- Build the add-in ---
FROM node:22-alpine AS addin
WORKDIR /build
COPY addin/package*.json ./
RUN npm install
COPY addin/ ./
RUN npm run build

# --- Build the backend ---
FROM node:22-alpine AS server
WORKDIR /build
COPY server/package*.json ./
RUN npm install
COPY server/ ./
RUN npm run build && npm prune --omit=dev

# --- Bitwarden CLI ---
# Must match the Vaultwarden version in use – test before updating. Changing it here → GitHub rebuilds.
# Matched to Vaultwarden 1.37.5 / web vault 2026.10.1.
# Installed as a local project (not npm -g) so vulnerable transitive dependencies can be pinned to
# fixed versions via "overrides" (the Trivy scan in the workflow fails on fixed HIGH/CRITICAL).
FROM node:22-alpine AS bw
ARG BW_CLI_VERSION=2026.10.1
WORKDIR /opt/bw
RUN printf '{"private":true,"dependencies":{"@bitwarden/cli":"%s"},"overrides":{"node-forge":"^1.4.1","multer":"^2.5.0","tmp":"^0.2.6"}}' \
      "${BW_CLI_VERSION}" > package.json \
 && npm install --omit=dev --no-audit --no-fund \
 && node_modules/.bin/bw --version

# --- Runtime ---
FROM node:22-alpine
# npm/corepack are not needed at runtime and only add attack surface (and scanner findings).
RUN apk add --no-cache tini \
 && rm -rf /usr/local/lib/node_modules/npm /usr/local/lib/node_modules/corepack \
           /usr/local/bin/npm /usr/local/bin/npx /usr/local/bin/corepack
COPY --from=bw /opt/bw /opt/bw
RUN ln -s /opt/bw/node_modules/.bin/bw /usr/local/bin/bw \
 && test -x "$(readlink -f /usr/local/bin/bw)"

ENV NODE_ENV=production \
    BITWARDENCLI_APPDATA_DIR=/home/node/.bw \
    DATA_DIR=/data

WORKDIR /app
COPY --from=server /build/package.json ./
COPY --from=server /build/node_modules ./node_modules
COPY --from=server /build/dist ./dist
COPY --from=addin /build/dist ./public.tpl
COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod 755 /entrypoint.sh \
 && mkdir -p /app/public /home/node/.bw /data \
 && chown node:node /app/public /home/node/.bw /data

USER node
EXPOSE 3000
ENTRYPOINT ["/sbin/tini", "--", "/entrypoint.sh"]
