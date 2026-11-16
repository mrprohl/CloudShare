#!/usr/bin/env bash
# Claude Code Stop hook: before Claude finishes, check that changed code came with changed documentation.
# Looks at the working tree (uncommitted and untracked) and the commits of the branch since public/main (or main).
# Code: addin/src, server/src, addin/manifest.xml. Documentation: README.md, docs/, CLAUDE.md, cloudshare.env.example,
# locales/ (help texts). A changed pane (taskpane html/css/ts) also asks for the README pictures. Blocks once per stop
# (stop_hook_active lets the second attempt through), so Claude either updates the docs or says why none are needed.
set -u
input=$(cat)
case "$input" in *'"stop_hook_active":true'*|*'"stop_hook_active": true'*) exit 0 ;; esac

cd "$(git rev-parse --show-toplevel 2>/dev/null)" || exit 0

base=""
for ref in refs/remotes/public/main refs/remotes/origin/main; do
  if git rev-parse --verify -q "$ref" >/dev/null; then base=$(git merge-base HEAD "$ref" 2>/dev/null) && break; fi
done
changed=$( { git status --porcelain --untracked-files=all | sed -E 's/^.. //; s/.* -> //'; [ -n "$base" ] && git diff --name-only "$base" HEAD; } | sort -u)
[ -z "$changed" ] && exit 0

code=$(grep -E '^(addin/src/|server/src/|addin/manifest\.xml$)' <<<"$changed" || true)
docs=$(grep -E '^(README\.md|CLAUDE\.md|cloudshare\.env\.example|docs/|locales/)' <<<"$changed" || true)
pane=$(grep -E '^addin/src/taskpane/.*\.(html|css|ts)$' <<<"$changed" || true)
pictures=$(grep -E '^docs/images/' <<<"$changed" || true)

reasons=()
if [ -n "$code" ] && [ -z "$docs" ]; then
  reasons+=("Code changed without any documentation change: $(tr '\n' ' ' <<<"$code"). Check README.md (features, settings table), docs/architecture.md (flows, data, API table), the help texts help.* in locales/*.json, CLAUDE.md and cloudshare.env.example – update what the change affects, or tell the user why nothing needs updating.")
fi
if [ -n "$pane" ] && [ -z "$pictures" ]; then
  reasons+=("The taskpane changed but no README picture (docs/images): if the change is visible in a README picture, the UI tests will fail – after the CI run take the pictures from the artifact readme-pictures. Mockup pictures (docs/mockups) need re-rendering when the card or a mail changed.")
fi
[ ${#reasons[@]} -eq 0 ] && exit 0

reason=$(printf '%s ' "${reasons[@]}" | sed 's/\\/\\\\/g; s/"/\\"/g')
printf '{"decision":"block","reason":"Documentation check: %s"}\n' "$reason"
