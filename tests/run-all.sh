#!/usr/bin/env bash
# Repo-wide test gate: build the projects under test, then run every test
# suite. Each .fsx references this worktree's build output by relative path.
#
# Usage: tests/run-all.sh [--skip-build]
# Suitable as a pre-push hook: ln -s ../../tests/run-all.sh .git/hooks/pre-push
set -u
cd "$(dirname "$0")/.."
ROOT="$(pwd)"

# Built one by one rather than through Office.slnx: the solution also pulls in
# sibling-repo hosts (widgets, wisp) that these suites do not exercise.
PROJECTS=(
    src/Office.Formats/Rbt.OfficeFormat.fsproj
    src/Word/Rbt.Word.fsproj
    src/Word.Pdf/Rbt.Word.Pdf.fsproj
)

if [ "${1:-}" != "--skip-build" ]; then
    for p in "${PROJECTS[@]}"; do
        echo "building $p..."
        if ! dotnet build "$p" -v quiet >/dev/null; then
            echo "BUILD FAILED: $p" >&2
            exit 1
        fi
    done
fi

failed=0

run_suite () {
    # run_suite <label> <fsx-file>
    local label="$1" f="$2"
    [ -f "$f" ] || { echo "--- $label (missing, skipped) ---"; return; }
    echo "--- $label ---"
    if ! dotnet fsi --quiet "$f"; then
        failed=1
    fi
}

echo "=== formula engine ==="
if ! bash tests/formula/run-all.sh; then failed=1; fi

echo "=== office formats ==="
for f in tests/office-formats/*.fsx; do
    run_suite "$(basename "$f" .fsx)" "$f"
done

echo "=== word ==="
for s in textbuffer layout incremental-layout pdf; do
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
