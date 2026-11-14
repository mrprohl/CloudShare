#!/usr/bin/env bash
# ============================================================================
# CloudShare – CI checks of the deployment files and the running stack (GitHub Actions, Ubuntu runner).
#   ci/stack.sh validate                 setup.sh (non-interactive), docker compose config, caddy validate
#   ci/stack.sh smoke <image> <tag>      starts the real stack (Caddy + CloudShare) and tests it
#                                        from outside: TLS, redirects, headers, add-in files, API, hardening
#   ci/stack.sh upgrade <previous-image> <image> <tag>
#                                        data written by the previous image (connections, password jobs) must
#                                        still be readable by the new one, which must start on it
#   ci/stack.sh load <image> <tag>       k6: parallel users upload large files through Caddy (Nextcloud and
#                                        Microsoft token keys mocked); errors, latency, memory, completeness
# Uses dummy settings; the Bitwarden CLI is replaced by a mock (ci/fake-bw), nothing leaves the runner except
# the Caddy image pull. The repository is linked to /opt/cloudshare because the compose files use that path.
# ============================================================================
set -euo pipefail

repo=$(cd "$(dirname "$0")/.." && pwd)
base=/opt/cloudshare
failures=0
# Additional compose files (e.g. the load-test mocks)
extra=()

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
de

enforce
14
90
{displayName} · {jobTitle}\n{company}
00000000-0000-0000-0000-000000000001
00000000-0000-0000-0000-000000000002
ci-dummy-secret
https://cloud.example.com
CloudShare
CloudShare-Dateianforderungen
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
    -f "$base/compose.yml" -f "$base/ci/compose.smoke.yml" "${extra[@]}" "$@"
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

# EXIT trap of the stack tests: on any failure show the container logs; always remove the stack afterwards.
cleanup() {
  local rc=$?
  if [ "$rc" -ne 0 ] || [ "$failures" -gt 0 ]; then compose logs --no-color | tail -n 80; fi
  compose down -v >/dev/null 2>&1 || true
}

# set_image <image> <tag> – which CloudShare image the stack runs (env/versions.env)
set_image() {
  sed -i '/^CLOUDSHARE_IMAGE=/d; /^CLOUDSHARE_VERSION=/d' "$base/env/versions.env"
  printf 'CLOUDSHARE_IMAGE=%s\nCLOUDSHARE_VERSION=%s\n' "$1" "$2" >> "$base/env/versions.env"
}

smoke() {
  local image=$1 tag=$2
  set_image "$image" "$tag"
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
  local pane
  pane=$(curl -sk https://localhost/taskpane.html)
  if grep -q '<html lang="de">' <<<"$pane" && grep -q 'Als Link einfügen' <<<"$pane" \
    && curl -sk https://localhost/commands.js | grep -q 'alert.button'; then
    ok "texts of LANGUAGE=de in the add-in files"
  else fail "language texts missing in the add-in files"; fi
  config=$(curl -sk https://localhost/api/config)
  jq -e '.attachmentPolicy == "enforce" and .thresholdMb == 5 and .sendMaxAccessCount == 3' <<<"$config" >/dev/null \
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
  local caps
  caps=$(inspect cloudshare-caddy '{{.HostConfig.CapDrop}} {{.HostConfig.CapAdd}}')
  # Docker may report capabilities with or without the CAP_ prefix
  [[ "$caps" =~ ^\[ALL\]\ \[(CAP_)?NET_BIND_SERVICE\]$ ]] \
    && ok "caddy drops all capabilities except NET_BIND_SERVICE" || fail "caddy capabilities: $caps"
  [ "$(docker network inspect -f '{{.Internal}}' cloudshare_internal)" = true ] \
    && ok "network internal has no outbound access" || fail "network internal is not internal"
}

# Runs node with a script from ci/ inside an image on the stack's data volume, with the app's settings.
node_in() {
  local image=$1 script=$2
  ( set -a; . "$base/env/cloudshare.env"; set +a
    docker run --rm --user 1000:1000 --read-only --network none \
      -e DATA_DIR=/data -e PUBLIC_BASE_URL -e ENTRA_TENANT_ID -e ENTRA_CLIENT_ID -e ENTRA_CLIENT_SECRET \
      -e NEXTCLOUD_URL -e CREDENTIAL_KEY -e BW_PASSWORD \
      -v cloudshare_data:/data -v "$base/ci:/ci:ro" --entrypoint node "$image" "$script" )
}

upgrade() {
  local previous=$1 image=$2 tag=$3
  step "Upgrade test: data of $previous → $image:$tag"
  if ! docker pull -q "$previous" >/dev/null 2>&1; then
    echo "  – previous image $previous not available, upgrade test skipped"; return
  fi
  trap cleanup EXIT
  compose down -v >/dev/null 2>&1 || true
  docker volume create --label com.docker.compose.project=cloudshare --label com.docker.compose.volume=data \
    cloudshare_data >/dev/null

  if node_in "$previous" /ci/upgrade/seed.mjs >/dev/null; then ok "previous version wrote connections and password jobs"
  else fail "previous version could not write its data"; return; fi

  echo "  new version reads them:"
  if node_in "$image:$tag" /ci/upgrade/verify.mjs; then ok "all data readable and unchanged"
  else fail "data of the previous version is not readable"; fi

  set_image "$image" "$tag"
  compose up -d --quiet-pull
  if wait_healthy; then ok "new version starts on the existing data"; else fail "new version not healthy"; return; fi
  docker logs cloudshare 2>&1 | grep -q 'could not be read' && fail "startup warns about unreadable data" \
    || ok "no warnings about unreadable data at startup"
  stat_mode=$(docker run --rm -v cloudshare_data:/data alpine:3 stat -c '%a %u' /data/nextcloud-credentials.json)
  [ "$stat_mode" = "600 1000" ] && ok "credentials file stays 600, owned by 1000" || fail "credentials file: $stat_mode"
}

load() {
  local image=$1 tag=$2 users=${LOAD_USERS:-10} files=${LOAD_FILES:-2} mib=${LOAD_FILE_MIB:-45}
  extra=(-f "$base/ci/compose.load.yml")
  set_image "$image" "$tag"
  trap cleanup EXIT

  step "Load test: $users users × $files files × $mib MiB through Caddy (Nextcloud and token keys mocked)"
  compose up -d --quiet-pull
  if wait_healthy; then ok "stack with mocks healthy"; else fail "cloudshare not healthy"; return; fi
  for _ in $(seq 1 20); do curl -sf http://127.0.0.1:8099/stats >/dev/null && break; sleep 1; done

  # Memory of the CloudShare container, sampled every second while k6 runs (MiB).
  local samples=/tmp/cloudshare-mem.txt
  : > "$samples"
  ( while true; do
      docker stats --no-stream --format '{{.MemUsage}}' cloudshare 2>/dev/null | awk '{
        v = $1; unit = v; gsub(/[0-9.]/, "", unit); sub(/[A-Za-z]+$/, "", v)
        print (unit == "GiB" ? v * 1024 : unit == "KiB" ? v / 1024 : unit == "B" ? v / 1048576 : v) }' >> "$samples"
      sleep 1
    done ) &
  local sampler=$!

  local k6_status=0
  docker run --rm --network host -e USERS="$users" -e FILES_PER_USER="$files" -e FILE_MIB="$mib" \
    -v "$base/ci/load:/scripts:ro" grafana/k6:2.3.0 run --no-color /scripts/upload.js || k6_status=$?
  kill "$sampler" 2>/dev/null || true

  step "Load test results"
  [ "$k6_status" -eq 0 ] && ok "k6 thresholds met (no failed request, p95 chunk/complete < 3 s)" \
    || fail "k6 thresholds not met (exit $k6_status)"
  local stats expected_bytes=$((users * files * mib * 1024 * 1024))
  stats=$(curl -s http://127.0.0.1:8099/stats)
  jq -e --argjson b "$expected_bytes" '.chunkBytes == $b' <<<"$stats" >/dev/null \
    && ok "every byte arrived at Nextcloud ($((expected_bytes / 1048576)) MiB)" || fail "bytes at Nextcloud: $stats"
  jq -e --argjson n "$((users * files))" '.uploadsFinished == $n and .shares == $n' <<<"$stats" >/dev/null \
    && ok "$((users * files)) uploads completed and shared" || fail "uploads/shares: $stats"

  local peak limit
  peak=$(sort -n "$samples" | tail -n 1 | cut -d. -f1)
  limit=$(( $(inspect cloudshare '{{.HostConfig.Memory}}') / 1048576 ))
  [ -n "$peak" ] && [ "$peak" -lt $((limit * 8 / 10)) ] && ok "memory peak ${peak} MiB (< 80 % of ${limit} MiB)" \
    || fail "memory peak ${peak:-?} MiB of ${limit} MiB"
  [ "$(inspect cloudshare '{{.State.OOMKilled}} {{.RestartCount}}')" = "false 0" ] \
    && ok "no out-of-memory kill, no restart" || fail "container was killed or restarted"
  wait_healthy && ok "still healthy after the load" || fail "unhealthy after the load"
}

case "${1:-}" in
  validate) link_repo; run_setup; validate ;;
  smoke)    [ $# -eq 3 ] || { echo "usage: $0 smoke <image> <tag>" >&2; exit 2; }
            link_repo; run_setup; validate; smoke "$2" "$3" ;;
  upgrade)  [ $# -eq 4 ] || { echo "usage: $0 upgrade <previous-image> <image> <tag>" >&2; exit 2; }
            link_repo; run_setup >/dev/null; upgrade "$2" "$3" "$4" ;;
  load)     [ $# -eq 3 ] || { echo "usage: $0 load <image> <tag>" >&2; exit 2; }
            link_repo; run_setup >/dev/null; load "$2" "$3" ;;
  *)        echo "usage: $0 validate | smoke <image> <tag> | upgrade <previous-image> <image> <tag> | load <image> <tag>" >&2
            exit 2 ;;
esac

echo
if [ "$failures" -gt 0 ]; then echo "$failures check(s) failed." >&2; exit 1; fi
echo "All checks passed."
