#!/usr/bin/env bash
# Repo-wide test gate: build the solution, then run every test suite against
# THIS worktree's freshly built dlls. Each suite's `#r` paths are normalized
# to the local build (the .fsx files hard-code the author's absolute paths),
# so the gate is portable across machines and worktrees.
#
# Usage: tests/run-all.sh [--skip-build]
# Suitable as a pre-push hook: ln -s ../../tests/run-all.sh .git/hooks/pre-push
set -u
cd "$(dirname "$0")/.."
ROOT="$(pwd)"
FSVB="$(cd .. && pwd)/fsvb"

if [ "${1:-}" != "--skip-build" ]; then
    echo "building Office.slnx..."
    if ! dotnet build Office.slnx -v quiet >/dev/null; then
        echo "SOLUTION BUILD FAILED" >&2
        exit 1
    fi
fi

failed=0

run_suite () {
    # run_suite <label> <fsx-file>
    local label="$1" f="$2"
    [ -f "$f" ] || { echo "--- $label (missing, skipped) ---"; return; }
    echo "--- $label ---"
    local tmp
    tmp="$(mktemp --suffix=.fsx)"
    # Normalize every hard-coded #r path to this worktree / sibling fsvb.
    sed -E \
        -e "s#\"[^\"]*/office/(src|word-native|spreadsheet-native)/#\"$ROOT/\1/#" \
        -e "s#\"[^\"]*/fsvb/bin/#\"$FSVB/bin/#" \
        -e "s#\"/home/[^/]+/\.nuget/#\"$HOME/.nuget/#" \
        "$f" > "$tmp"
    if ! dotnet fsi --quiet "$tmp"; then
        failed=1
    fi
    rm -f "$tmp"
}

echo "=== formula engine ==="
if ! bash tests/formula/run-all.sh; then failed=1; fi

echo "=== office formats ==="
for f in tests/office-formats/*.fsx; do
    run_suite "$(basename "$f" .fsx)" "$f"
done

echo "=== word ==="
for s in textbuffer layout pdf; do
    run_suite "word-$s" "tests/word/$s.fsx"
done

echo "=== print ==="
run_suite "print-helpers" "tests/print/helpers.fsx"

echo ""
if [ "$failed" -ne 0 ]; then
    echo "OFFICE TESTS: FAILED"
    exit 1
fi
echo "OFFICE TESTS: ALL SUITES PASSED"
