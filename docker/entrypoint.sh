#!/bin/sh
set -eu

: "${PUBLIC_BASE_URL:?PUBLIC_BASE_URL is missing}"
: "${ENTRA_CLIENT_ID:?ENTRA_CLIENT_ID is missing}"
# VAULTWARDEN_ENABLED=false: no Vaultwarden: the password goes in the password mail itself (server/src/vaultwarden.ts).
VAULTWARDEN_ENABLED="${VAULTWARDEN_ENABLED:-true}"
if [ "$VAULTWARDEN_ENABLED" != "false" ]; then
  : "${VAULTWARDEN_URL:?VAULTWARDEN_URL is missing}"
  : "${BW_CLIENTID:?BW_CLIENTID is missing}"
  : "${BW_CLIENTSECRET:?BW_CLIENTSECRET is missing}"
fi

PUBLIC_BASE_URL="${PUBLIC_BASE_URL%/}"
PUBLIC_HOST="${PUBLIC_BASE_URL#*://}"

# Adapt the add-in files to the environment (idempotent; the originals live in /app/public.tpl).
rm -rf /app/public/*
cp -r /app/public.tpl/. /app/public/
for file in /app/public/*.html /app/public/*.js /app/public/manifest.xml; do
  sed -i \
    -e "s|%%PUBLIC_BASE_URL%%|${PUBLIC_BASE_URL}|g" \
    -e "s|%%PUBLIC_HOST%%|${PUBLIC_HOST}|g" \
    -e "s|%%ENTRA_CLIENT_ID%%|${ENTRA_CLIENT_ID}|g" \
    "$file"
done

if [ "$VAULTWARDEN_ENABLED" != "false" ]; then
  # Log the Bitwarden CLI in with the Vaultwarden service account (API key from BW_CLIENTID/BW_CLIENTSECRET).
  bw logout >/dev/null 2>&1 || true
  bw config server "$VAULTWARDEN_URL" >/dev/null
  bw login --apikey >/dev/null

  # Vault Management API on localhost only; the backend unlocks the vault at startup.
  bw serve --hostname 127.0.0.1 --port 8087 &
fi

exec node /app/dist/index.js
