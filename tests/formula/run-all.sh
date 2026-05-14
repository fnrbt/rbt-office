#!/usr/bin/env bash
# Run every formula-engine test stage. Each stage script exits non-zero on
# the first failed assertion; this runner aggregates them and fails if any
# stage fails. Build the Spreadsheet project first — the scripts #r its dll.
set -u
cd "$(dirname "$0")/../.."

echo "building RedBat.Spreadsheet..."
if ! dotnet build src/Spreadsheet/RedBat.Spreadsheet.fsproj -v quiet >/dev/null; then
    echo "BUILD FAILED" >&2
    exit 1
fi

stages=(
    tests/formula/stage1-operators.fsx
    tests/formula/stage2-functions.fsx
    tests/formula/stage3-context.fsx
    tests/formula/stage4-perf.fsx
    tests/formula/stage5-calc.fsx
)

failed=0
for s in "${stages[@]}"; do
    echo "--- $s ---"
    if ! dotnet fsi "$s"; then
        failed=1
    fi
done

echo ""
if [ "$failed" -ne 0 ]; then
    echo "FORMULA TESTS: FAILED"
    exit 1
fi
echo "FORMULA TESTS: ALL STAGES PASSED"
