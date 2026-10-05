#!/usr/bin/env bash
# Run every formula-engine test stage. Each stage script exits non-zero on
# the first failed assertion; this runner aggregates them and fails if any
# stage fails. Builds the Spreadsheet project first; each stage references
# this worktree's build output by relative path.
set -u
cd "$(dirname "$0")/../.."

# Office.Formats is needed by the vbaProject.bin stage.
for proj in src/Spreadsheet/Rbt.Spreadsheet.fsproj src/Office.Formats/Rbt.OfficeFormat.fsproj; do
    echo "building $proj..."
    if ! dotnet build "$proj" -v quiet >/dev/null; then
        echo "BUILD FAILED: $proj" >&2
        exit 1
    fi
done

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
    stage13-pivot
    stage14-cf
    stage15-validation
    stage16-filter
    stage17-tables
    stage18-arrays2
    stage19-vba
    stage20-vbabin
    stage21-structure
    stage22-semantics
    stage23-functions
    stage24-moves
    stage25-arrays3
    stage26-rules
    stage27-input
    stage28-layout
    stage29-borders
    stage30-overflow
    stage31-edits
    stage32-recalc
    stage33-sheetops
    stage34-cfapi
    stage35-validations
    stage36-names
    stage37-sheets
    stage38-rowcol
    stage39-sortfilter
    stage40-annotations
    stage41-find
    stage42-tables
    stage43-status
    numbertext
)

failed=0
for s in "${stages[@]}"; do
    f="tests/formula/$s.fsx"
    [ -f "$f" ] || { echo "--- $s (missing, skipped) ---"; continue; }
    echo "--- $s ---"
    if ! dotnet fsi --quiet "$f"; then
        failed=1
    fi
done

echo ""
if [ "$failed" -ne 0 ]; then
    echo "FORMULA TESTS: FAILED"
    exit 1
fi
echo "FORMULA TESTS: ALL STAGES PASSED"
