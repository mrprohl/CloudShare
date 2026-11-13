#!/usr/bin/env bash
# ============================================================================
# CloudShare – setup wizard for a Docker host with a central Caddy (see README)
# Creates the internal network cloudshare-proxy (<PROXY_NET>.0/24, central Caddy = <PROXY_NET>.2), writes env/
# (directory 700, files 600): versions.env, cloudshare.env – and caddy/cloudshare.caddy from the template.
# Enter accepts the default shown in [brackets]. Secrets are prompted without echo.
# Location: /opt/cloudshare (fixed, as in docker-compose.yml)
# Other network: PROXY_NET=10.99.0 ./setup.sh   (first three octets of a free /24)
# ============================================================================
set -euo pipefail
cd "$(dirname "$0")"
[ "$(pwd)" = /opt/cloudshare ] || echo "Note: package is in $(pwd), Compose expects /opt/cloudshare." >&2
PROXY_NET=${PROXY_NET:-172.30.13}

# Internal network shared with the central Caddy; containers get .128 and up, so .2 stays free for Caddy
if command -v docker >/dev/null; then
  docker network inspect cloudshare-proxy >/dev/null 2>&1 || {
    docker network create --internal --subnet "$PROXY_NET.0/24" --ip-range "$PROXY_NET.128/25" cloudshare-proxy >/dev/null \
      && echo "Network cloudshare-proxy ($PROXY_NET.0/24, internal) created."; }
fi
chmod -R a+rX caddy

init_versions() {   # create env/versions.env from the template (an existing one stays unchanged)
  umask 077; mkdir -p env; chmod 700 env
  if [ -f env/versions.env ]; then echo "env/versions.env exists – unchanged."
  else sed "s/^PROXY_IP=.*/PROXY_IP=$PROXY_NET.2/" versions.env.example > env/versions.env
       echo "env/versions.env created (template)."; fi
  chmod 600 env/versions.env
}
init_versions
if [ -f env/cloudshare.env ]; then
  echo "Configuration in env/ exists – to regenerate, back it up and delete it first."; exit 0
fi

# ask VAR "Text" "default" – visible input, Enter = default
ask()        { local answer; read -rp "$2${3:+ [$3]}: " answer; printf -v "$1" '%s' "${answer:-$3}"; }
# ask_secret VAR "Text" – hidden input, must not be empty
ask_secret() { local answer=""; while [ -z "$answer" ]; do read -rsp "$2: " answer; echo; done; printf -v "$1" '%s' "$answer"; }
# Value for an env file: literal in '…'; if it contains ' itself, then "…" with escapes ($ → $$ against interpolation)
env_value() {
  case "$1" in
    *"'"*) local escaped=${1//\\/\\\\}; escaped=${escaped//\"/\\\"}; escaped=${escaped//\$/\$\$}; printf '"%s"' "$escaped" ;;
    *)     printf "'%s'" "$1" ;;
  esac
}

echo "== CloudShare setup =="
echo "-- General --"
ask PUBLIC_BASE_URL         "Public address"                              "https://cloudshare.example.com"
ask ATTACHMENT_THRESHOLD_MB "CloudShare from attachment size (MB)"        "1"
ask ATTACHMENT_POLICY       "Large attachments: enforce or suggest"       "enforce"
ask SHARE_EXPIRE_DAYS       "Link valid, default (days)"                  "14"
ask SHARE_MAX_EXPIRE_DAYS   "Link valid, maximum (days)"                  "90"
echo "Mail card footer – placeholders from the M365 profile: {displayName} {givenName} {surname} {jobTitle}"
echo '  {department} {company} {email} {phone} {mobile} {office}; \n = new line; "-" = no footer'
ask MAIL_FOOTER             "Footer template"                             '{displayName} · {jobTitle}\n{company} · Tel. {phone} · {email}'
[ "$MAIL_FOOTER" = "-" ] && MAIL_FOOTER=""
echo; echo "-- Microsoft Entra ID (app registration \"CloudShare\") --"
ask ENTRA_TENANT_ID     "Directory (tenant) ID"     ""
ask ENTRA_CLIENT_ID     "Application (client) ID"   ""
ask_secret ENTRA_CLIENT_SECRET "Client secret (value)"
echo; echo "-- Nextcloud --"
ask NEXTCLOUD_URL    "Nextcloud address"        "https://cloud.example.com"
ask NEXTCLOUD_LOCAL  "Nextcloud behind the same central Caddy (y/n)" "y"
ask NEXTCLOUD_FOLDER "Folder for attachments"   "Outlook-Anhaenge"
ask NEXTCLOUD_REQUEST_FOLDER "Folder for upload links" "Outlook-Dateianforderungen"
echo; echo "-- Vaultwarden service account (own user, e.g. cloudshare@<your domain>) --"
ask VAULTWARDEN_URL       "Vaultwarden address"        "https://vault.example.com"
ask BW_CLIENTID           "API key client_id"          ""
ask_secret BW_CLIENTSECRET       "API key client_secret"
ask_secret BW_PASSWORD           "Master password"
ask SEND_MAX_ACCESS_COUNT "Password link: accesses (default)" "3"
ask SEND_MAX_ACCESS_LIMIT "Password link: accesses (maximum)" "20"
ask SEND_EXPIRE_DAYS      "Password link: valid (days)"     "7"

for var in PUBLIC_BASE_URL ENTRA_TENANT_ID ENTRA_CLIENT_ID NEXTCLOUD_URL VAULTWARDEN_URL BW_CLIENTID; do
  [ -n "${!var}" ] || { echo "$var must not be empty." >&2; exit 1; }
done
case "$ATTACHMENT_POLICY" in enforce|suggest) ;; *) echo "ATTACHMENT_POLICY must be enforce or suggest." >&2; exit 1;; esac
case "$NEXTCLOUD_LOCAL" in y|n) ;; *) echo "Answer y or n for the Nextcloud question." >&2; exit 1;; esac
for var in PUBLIC_BASE_URL NEXTCLOUD_URL VAULTWARDEN_URL; do
  case "${!var}" in https://*) ;; *) echo "$var must start with https://." >&2; exit 1;; esac
done
case "$BW_CLIENTID" in user.*) ;; *) echo "BW_CLIENTID starts with user." >&2; exit 1;; esac
for var in ATTACHMENT_THRESHOLD_MB SHARE_EXPIRE_DAYS SHARE_MAX_EXPIRE_DAYS SEND_MAX_ACCESS_COUNT SEND_MAX_ACCESS_LIMIT SEND_EXPIRE_DAYS; do
  [[ "${!var}" =~ ^[0-9]+(\.[0-9]+)?$ ]] || { echo "$var must be a number." >&2; exit 1; }
done
# shellcheck disable=SC2034  # written to env/cloudshare.env indirectly via ${!var} below
CREDENTIAL_KEY=$(openssl rand -base64 32)

umask 077
{
  echo "# written by setup.sh on $(date '+%Y-%m-%d %H:%M') – contains secrets"
  for var in PUBLIC_BASE_URL ATTACHMENT_THRESHOLD_MB ATTACHMENT_POLICY SHARE_EXPIRE_DAYS SHARE_MAX_EXPIRE_DAYS \
           MAIL_FOOTER ENTRA_TENANT_ID ENTRA_CLIENT_ID ENTRA_CLIENT_SECRET \
           NEXTCLOUD_URL NEXTCLOUD_FOLDER NEXTCLOUD_REQUEST_FOLDER \
           CREDENTIAL_KEY VAULTWARDEN_URL BW_CLIENTID BW_CLIENTSECRET BW_PASSWORD \
           SEND_MAX_ACCESS_COUNT SEND_MAX_ACCESS_LIMIT SEND_EXPIRE_DAYS; do
    echo "$var=$(env_value "${!var}")"
  done
} > env/cloudshare.env
chmod 700 env; chmod 600 env/*
chmod 755 setup.sh

# Nextcloud behind the same Caddy → reach it internally via extra_hosts (env/versions.env)
NEXTCLOUD_HOST=${NEXTCLOUD_URL#https://}; NEXTCLOUD_HOST=${NEXTCLOUD_HOST%%/*}
if [ "$NEXTCLOUD_LOCAL" = y ]; then sed -i "s|^#\?NEXTCLOUD_HOST=.*|NEXTCLOUD_HOST=$NEXTCLOUD_HOST|" env/versions.env
else sed -i "s|^NEXTCLOUD_HOST=|#NEXTCLOUD_HOST=|" env/versions.env; fi

# Site block for the central Caddy from the template
PUBLIC_HOST=${PUBLIC_BASE_URL#https://}; PUBLIC_HOST=${PUBLIC_HOST%%/*}
if [ -f caddy/cloudshare.caddy ]; then echo "caddy/cloudshare.caddy exists – unchanged."
else sed "s/cloudshare\.example\.com/$PUBLIC_HOST/g" caddy/cloudshare.caddy.example > caddy/cloudshare.caddy
     chmod 644 caddy/cloudshare.caddy; echo "caddy/cloudshare.caddy written for $PUBLIC_HOST."; fi

echo; echo "env/ written:"; ls -l env | sed 's/^/  /'
echo "Store CREDENTIAL_KEY in your password manager (without it all users have to reconnect):"
echo "  grep CREDENTIAL_KEY env/cloudshare.env"
echo "Next: add cloudshare-proxy and /opt/cloudshare/caddy to the central Caddy (README), then deploy the stack."
