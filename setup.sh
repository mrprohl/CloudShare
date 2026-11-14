#!/usr/bin/env bash
# ============================================================================
# CloudShare – setup wizard
# Writes env/ (directory 700, files 600): versions.env (Compose variables incl. CLOUDSHARE_HOST for Caddy)
# and cloudshare.env (app settings and secrets).
# Enter accepts the default shown in [brackets]. Secrets are prompted without echo.
# Location: /opt/cloudshare (fixed, as in docker-compose.yml)
# ============================================================================
set -euo pipefail
cd "$(dirname "$0")"
[ "$(pwd)" = /opt/cloudshare ] || echo "Note: package is in $(pwd), Compose expects /opt/cloudshare." >&2
chmod -R a+rX caddy

init_versions() {   # create env/versions.env from the template (an existing one stays unchanged)
  umask 077; mkdir -p env; chmod 700 env
  if [ -f env/versions.env ]; then echo "env/versions.env exists – unchanged."
  else cp versions.env.example env/versions.env; echo "env/versions.env created (template)."; fi
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
ask LANGUAGE                "Language (de, en – files in locales/)"       "de"
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
ask NEXTCLOUD_FOLDER "Folder for attachments"   "CloudShare"
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
[[ "$LANGUAGE" =~ ^[a-z]{2,3}(-[a-z0-9]{2,8})?$ ]] || { echo "LANGUAGE must be a language code like de or en." >&2; exit 1; }
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
  for var in PUBLIC_BASE_URL LANGUAGE ATTACHMENT_THRESHOLD_MB ATTACHMENT_POLICY SHARE_EXPIRE_DAYS SHARE_MAX_EXPIRE_DAYS \
           MAIL_FOOTER ENTRA_TENANT_ID ENTRA_CLIENT_ID ENTRA_CLIENT_SECRET \
           NEXTCLOUD_URL NEXTCLOUD_FOLDER NEXTCLOUD_REQUEST_FOLDER \
           CREDENTIAL_KEY VAULTWARDEN_URL BW_CLIENTID BW_CLIENTSECRET BW_PASSWORD \
           SEND_MAX_ACCESS_COUNT SEND_MAX_ACCESS_LIMIT SEND_EXPIRE_DAYS; do
    echo "$var=$(env_value "${!var}")"
  done
} > env/cloudshare.env
chmod 700 env; chmod 600 env/*.env
# Optional files for the container (/config, read-only): templates.json and locales/<language>.json
mkdir -p env/config/locales; chmod 755 env/config env/config/locales
chmod 755 setup.sh

# Host name for Caddy (only on a new env/versions.env from the template)
PUBLIC_HOST=${PUBLIC_BASE_URL#https://}; PUBLIC_HOST=${PUBLIC_HOST%%/*}
sed -i "s|^CLOUDSHARE_HOST=cloudshare.example.com$|CLOUDSHARE_HOST=$PUBLIC_HOST|" env/versions.env
echo "CLOUDSHARE_HOST in env/versions.env: $(grep ^CLOUDSHARE_HOST= env/versions.env)"

echo; echo "env/ written:"; ls -l env | sed 's/^/  /'
echo "Store CREDENTIAL_KEY in your password manager (without it all users have to reconnect):"
echo "  grep CREDENTIAL_KEY env/cloudshare.env"
echo "Next: DNS for $PUBLIC_HOST → this server, ports 80/443 open, then deploy the stack (README)."
