#!/usr/bin/env bash
# ============================================================================
# CloudShare – CI checks of the deployment files and the running stack (GitHub Actions, Ubuntu runner).
#   ci/stack.sh validate                 setup.sh (non-interactive), docker compose config, caddy validate
#   ci/stack.sh smoke <image> <tag>      additionally starts the real stack (Caddy + CloudShare) and tests it
#                                        from outside: TLS, redirects, headers, add-in files, API, hardening
# Uses dummy settings; the Bitwarden CLI is replaced by a mock (ci/fake-bw), nothing leaves the runner except
# the Caddy image pull. The repository is linked to /opt/cloudshare because the compose files use that path.
# ============================================================================
set -euo pipefail

repo=$(cd "$(dirname "$0")/.." && pwd)
base=/opt/cloudshare
failures=0

ok()   { echo "  ✓ $*"; }
fail() { echo "  ✗ $*" >&2; failures=$((failures + 1)); }
step() { echo; echo "== $* =="; }

link_repo() {
  if [ ! -e "$base" ]; then sudo ln -s "$repo" "$base"; fi
  [ "$(cd "$base" && pwd -P)" = "$(cd "$repo" && pwd -P)" ] || { echo "$base is not this repository" >&2; exit 1; }
}

# Runs the real setup wizard with dummy answers (one per prompt, in order) – also tests setup.sh itself.
run_setup() {
  step "setup.sh (non-interactive, dummy values)"
  # Never on a real installation: env/ holds the secrets and would be replaced by dummy values.
  if [ -d "$base/env" ] && [ "${CI:-}" != true ]; then
    echo "$base/env exists – refusing to overwrite it outside CI (run this on a test machine only)." >&2; exit 1
  fi
  rm -rf "$base/env"
  "$base/setup.sh" >/dev/null <<'EOF'
https://localhost
1
enforce
14
90
{displayName} · {jobTitle}\n{company}
00000000-0000-0000-0000-000000000001
00000000-0000-0000-0000-000000000002
ci-dummy-secret
https://cloud.example.com
Outlook-Anhaenge
Outlook-Dateianforderungen
https://vault.example.com
user.ci-dummy
ci-dummy-client-secret
ci-dummy-password
3
20
7
EOF
  [ "$(stat -c %a "$base/env")" = 700 ] && ok "env/ is 700" || fail "env/ is not 700"
  for file in versions.env cloudshare.env; do
    [ "$(stat -c %a "$base/env/$file")" = 600 ] && ok "env/$file is 600" || fail "env/$file is not 600"
  done
  grep -q '^CLOUDSHARE_HOST=localhost$' "$base/env/versions.env" && ok "CLOUDSHARE_HOST taken from the address" \
    || fail "CLOUDSHARE_HOST not set in env/versions.env"
  grep -q '^CREDENTIAL_KEY=.\{44,\}' "$base/env/cloudshare.env" && ok "CREDENTIAL_KEY generated" \
    || fail "CREDENTIAL_KEY missing"
}

compose() {
  docker compose -p cloudshare --env-file "$base/env/versions.env" \
    -f "$base/compose.yml" -f "$base/ci/compose.smoke.yml" "$@"
}

validate() {
  step "Compose files"
  docker compose -f "$base/docker-compose.yml" config -q && ok "docker-compose.yml (stub with include)" \
    || fail "docker-compose.yml invalid"
  compose config -q && ok "compose.yml with smoke-test override" || fail "compose.yml invalid"

  step "Caddyfile"
  local caddy_version
  caddy_version=$(sed -n 's/^CADDY_VERSION=//p' "$base/env/versions.env")
  docker run --rm -e CLOUDSHARE_HOST=cloudshare.example.com \
    -v "$base/caddy/Caddyfile:/etc/caddy/Caddyfile:ro" "caddy:${caddy_version:-2-alpine}" \
    caddy validate --config /etc/caddy/Caddyfile --adapter caddyfile >/dev/null 2>&1 \
    && ok "caddy validate" || fail "caddy validate failed"
}

wait_healthy() {
  local status
  for _ in $(seq 1 60); do
    status=$(docker inspect -f '{{.State.Health.Status}}' cloudshare 2>/dev/null || echo missing)
    [ "$status" = healthy ] && return 0
    [ "$status" = unhealthy ] && return 1
    sleep 3
  done
  return 1
}

# expect_status <expected> <description> <curl args…>
expect_status() {
  local expected=$1 description=$2 status; shift 2
  status=$(curl -sk -o /dev/null -w '%{http_code}' "$@" || true)
  [ "$status" = "$expected" ] && ok "$description ($status)" || fail "$description: expected $expected, got $status"
}

inspect() { docker inspect -f "$2" "$1"; }

# EXIT trap of the smoke test: on any failure show the container logs; always remove the stack afterwards.
cleanup() {
  local rc=$?
  if [ "$rc" -ne 0 ] || [ "$failures" -gt 0 ]; then compose logs --no-color | tail -n 80; fi
  compose down -v >/dev/null 2>&1 || true
}

smoke() {
  local image=$1 tag=$2
  printf 'CLOUDSHARE_IMAGE=%s\nCLOUDSHARE_VERSION=%s\n' "$image" "$tag" >> "$base/env/versions.env"
  trap cleanup EXIT

  step "Start stack ($image:$tag + Caddy)"
  compose up -d --quiet-pull
  if wait_healthy; then ok "cloudshare healthy"; else fail "cloudshare not healthy"; return; fi
  docker logs cloudshare 2>&1 | grep -q 'Vaultwarden vault unlocked' && ok "entrypoint ran, vault unlocked (mock)" \
    || fail "no 'Vaultwarden vault unlocked' in the log"

  step "Caddy: TLS, redirect, headers"
  local headers location
  location=$(curl -s -o /dev/null -w '%{redirect_url}' http://localhost/api/config || true)
  [ "$location" = "https://localhost/api/config" ] && ok "HTTP redirects to HTTPS" || fail "HTTP redirect: '$location'"
  headers=$(curl -sk -D - -o /dev/null https://localhost/taskpane.html)
  grep -qi '^strict-transport-security:' <<<"$headers" && ok "HSTS set" || fail "HSTS missing"
  grep -qi '^x-content-type-options: nosniff' <<<"$headers" && ok "nosniff set" || fail "X-Content-Type-Options missing"
  grep -qi '^x-frame-options:' <<<"$headers" && fail "X-Frame-Options set (breaks the Outlook taskpane)" \
    || ok "no X-Frame-Options (taskpane may be framed)"
  grep -qi '^server:' <<<"$headers" && fail "Server header exposed" || ok "no Server header"

  step "Add-in files and API"
  expect_status 200 "taskpane.html" https://localhost/taskpane.html
  expect_status 200 "commands.html" https://localhost/commands.html
  expect_status 200 "icon-80.png" https://localhost/assets/icon-80.png
  local manifest config
  manifest=$(curl -sk https://localhost/manifest.xml)
  grep -q 'https://localhost/taskpane.html' <<<"$manifest" && ok "manifest.xml points to this host" \
    || fail "manifest.xml without https://localhost/taskpane.html"
  grep -q '00000000-0000-0000-0000-000000000002' <<<"$manifest" && ok "manifest.xml contains the Entra client ID" \
    || fail "Entra client ID missing in manifest.xml"
  if curl -sk https://localhost/manifest.xml https://localhost/taskpane.html https://localhost/commands.js | grep -q '%%'; then
    fail "unreplaced %%placeholders%% in the add-in files"
  else ok "all placeholders replaced"; fi
  config=$(curl -sk https://localhost/api/config)
  jq -e '.attachmentPolicy == "enforce" and .thresholdMb == 1 and .sendMaxAccessCount == 3' <<<"$config" >/dev/null \
    && ok "/api/config reflects the settings" || fail "/api/config: $config"
  expect_status 401 "/api/shares without login" https://localhost/api/shares
  expect_status 401 "/api/nextcloud/status with an invalid token" -H 'Authorization: Bearer invalid' \
    https://localhost/api/nextcloud/status
  head -c 13000000 /dev/zero > /tmp/cloudshare-13mb.bin
  expect_status 413 "request body over 12 MB rejected by Caddy" -X PUT --data-binary @/tmp/cloudshare-13mb.bin \
    -H 'Content-Type: application/octet-stream' https://localhost/api/uploads/x/chunks/0

  step "Hardening"
  [ "$(inspect cloudshare '{{.HostConfig.ReadonlyRootfs}}')" = true ] && ok "cloudshare read-only" || fail "cloudshare writable"
  [ "$(inspect cloudshare '{{.Config.User}}')" = 1000:1000 ] && ok "cloudshare runs as 1000:1000" || fail "cloudshare user"
  [ "$(inspect cloudshare '{{.HostConfig.CapDrop}}')" = "[ALL]" ] && ok "cloudshare drops all capabilities" \
    || fail "cloudshare capabilities"
  [ "$(inspect cloudshare '{{len .HostConfig.PortBindings}}')" = 0 ] && ok "cloudshare publishes no ports" \
    || fail "cloudshare publishes ports"
  inspect cloudshare '{{.HostConfig.SecurityOpt}}' | grep -q 'no-new-privileges:true' && ok "no-new-privileges" \
    || fail "no-new-privileges missing"
  [ "$(inspect cloudshare-caddy '{{.HostConfig.ReadonlyRootfs}}')" = true ] && ok "caddy read-only" || fail "caddy writable"
  [ "$(inspect cloudshare-caddy '{{.HostConfig.CapAdd}}')" = "[NET_BIND_SERVICE]" ] \
    && ok "caddy only gets NET_BIND_SERVICE" || fail "caddy capabilities"
  [ "$(docker network inspect -f '{{.Internal}}' cloudshare_internal)" = true ] \
    && ok "network internal has no outbound access" || fail "network internal is not internal"
}

case "${1:-}" in
  validate) link_repo; run_setup; validate ;;
  smoke)    [ $# -eq 3 ] || { echo "usage: $0 smoke <image> <tag>" >&2; exit 2; }
            link_repo; run_setup; validate; smoke "$2" "$3" ;;
  *)        echo "usage: $0 validate | smoke <image> <tag>" >&2; exit 2 ;;
esac

echo
if [ "$failures" -gt 0 ]; then echo "$failures check(s) failed." >&2; exit 1; fi
echo "All checks passed."
