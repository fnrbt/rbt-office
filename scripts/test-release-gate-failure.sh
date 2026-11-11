#!/usr/bin/env bash
# Bounded control-flow test: simulate release prerequisites, fail the real
# --core invocation, and assert that the real publisher never calls npm publish.
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT
mkdir -p "$tmp/repo/.forge" "$tmp/repo/tests" "$tmp/repo/scripts" "$tmp/bin" "$tmp/bundle"
cp "$root/.forge/publish-npm" "$tmp/repo/.forge/publish-npm"
printf '{"version":"1.2.3"}\n' > "$tmp/repo/package.json"
printf 'deadbeef\n' > "$tmp/bundle/commit"
printf '1.2.3\n' > "$tmp/bundle/version"
printf 'package\n' > "$tmp/bundle/rbt-office-1.2.3.tgz"
(cd "$tmp/bundle" && sha256sum commit version rbt-office-1.2.3.tgz > SHA256SUMS)
cat > "$tmp/repo/tests/run-all.sh" <<'EOF'
#!/usr/bin/env bash
[[ "$*" == --core ]] || exit 2
printf 'core reached\n' > "$TEST_TMP/core-reached"
exit 1
EOF
cat > "$tmp/bin/git" <<'EOF'
#!/usr/bin/env bash
case "$*" in
  'rev-parse --show-toplevel') printf '%s/repo\n' "$TEST_TMP" ;;
  'rev-parse HEAD') echo deadbeef ;;
  'status --porcelain --untracked-files=all') : ;;
  *) exit 2 ;;
esac
EOF
cat > "$tmp/bin/node" <<'EOF'
#!/usr/bin/env bash
if [[ "$1" == -p ]]; then echo 1.2.3; fi
EOF
cat > "$tmp/bin/dotnet" <<'EOF'
#!/usr/bin/env bash
exit 0
EOF
cp "$tmp/bin/dotnet" "$tmp/bin/pnpm"
cat > "$tmp/bin/npm" <<'EOF'
#!/usr/bin/env bash
printf '%s\n' "$*" >> "$TEST_TMP/npm-calls"
EOF
chmod +x "$tmp/bin/"* "$tmp/repo/.forge/publish-npm"
export TEST_TMP="$tmp" PATH="$tmp/bin:$PATH" FORGE_VERSION=1.2.3
export FORGE_NPM_REGISTRY=https://invalid.example NPM_CONFIG_USERCONFIG="$tmp/npmrc"
if "$tmp/repo/.forge/publish-npm" --publish "$tmp/bundle" > "$tmp/output" 2>&1; then
    echo 'gate unexpectedly succeeded' >&2; exit 1
fi
[[ -f "$tmp/core-reached" ]] || { cat "$tmp/output" >&2; echo 'gate did not reach the failing core suite' >&2; exit 1; }
if [[ -f "$tmp/npm-calls" ]]; then
    case "$(cat "$tmp/npm-calls")" in *publish*) echo 'npm publish was called after a failing suite' >&2; exit 1 ;; esac
fi
echo 'failing core regression blocked npm publish'
