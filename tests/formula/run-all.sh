#!/usr/bin/env bash
# Run every formula-engine test stage. Each stage script exits non-zero on
# the first failed assertion; this runner aggregates them and fails if any
# stage fails. Builds the Spreadsheet project first and runs each stage
# against THIS worktree's freshly built dll (the `#r` path in each stage is
# normalized to the local build, so the suite tests local changes).
set -u
cd "$(dirname "$0")/../.."

echo "building RedBat.Spreadsheet..."
if ! dotnet build src/Spreadsheet/RedBat.Spreadsheet.fsproj -v quiet >/dev/null; then
    echo "BUILD FAILED" >&2
    exit 1
fi

DLL="$(pwd)/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

stages=(
    stage1-operators
    stage2-functions
    stage3-context
    stage4-perf
    stage5-calc
    stage6-modern
    stage7-streaming
    stage8-arrays
    stage9-lambda
    stage10-extended
    stage11-stats
    stage12-financial
)

failed=0
for s in "${stages[@]}"; do
    f="tests/formula/$s.fsx"
    [ -f "$f" ] || { echo "--- $s (missing, skipped) ---"; continue; }
    echo "--- $s ---"
    tmp="$(mktemp --suffix=.fsx)"
    # Normalize the #r dll reference to THIS worktree's freshly built dll.
    sed -E "s#\"[^\"]*RedBat\.Spreadsheet\.dll\"#\"$DLL\"#" "$f" > "$tmp"
    if ! dotnet fsi "$tmp"; then
        failed=1
    fi
    rm -f "$tmp"
done

echo ""
if [ "$failed" -ne 0 ]; then
    echo "FORMULA TESTS: FAILED"
    exit 1
fi
echo "FORMULA TESTS: ALL STAGES PASSED"
