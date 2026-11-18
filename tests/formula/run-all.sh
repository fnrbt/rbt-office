#!/usr/bin/env bash
# Run every mandatory formula stage; missing stages or fixtures fail before
# any build/test. --skip-build explicitly opts out of fresh build coverage.
# --check-coverage is the root runner's build-free preflight.
set -u
cd "$(dirname "$0")/../.."

skip_build=0
check_coverage=0
for option in "$@"; do
    case "$option" in
        --skip-build)
            if [ "$skip_build" -ne 0 ]; then
                echo "Duplicate option: --skip-build" >&2
                exit 2
            fi
            skip_build=1
            ;;
        --check-coverage)
            if [ "$check_coverage" -ne 0 ]; then
                echo "Duplicate option: --check-coverage" >&2
                exit 2
            fi
            check_coverage=1
            ;;
        *)
            echo "Unknown option: $option (use --skip-build or --check-coverage)" >&2
            exit 2
            ;;
    esac
done
if [ "$skip_build" -eq 1 ] && [ "$check_coverage" -eq 1 ]; then
    echo "--check-coverage cannot be combined with --skip-build" >&2
    exit 2
fi

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
    stage44-pagesetup
    stage45-order
    stage46-quicksort
    stage47-x14rules
    stage48-tableentry
    stage49-chrome
    stage50-date1904
    stage51-resultformat
    stage52-charts
    stage53-common
    stage54-tablestyles
    stage55-chartex
    stage56-slicers
    stage57-macro-host
    stage58-function-table
    stage59-math-functions
    stage60-statistics-functions
    stage61-financial-functions
    stage62-text-array-info-functions
    stage63-macro-showcase
    numbertext
)

required=(
    tests/formula/excel-2019-tints.tsv
    tests/test-documents-excel/pandas-test_decimal.xlsx
    tests/test-documents-excel/synthetic-table-styles.xlsx
    tests/test-documents-excel/synthetic-chartex.xlsx
    tests/test-documents-excel/synthetic-slicers.xlsx
    tests/real-world/chart-zoo.xlsx
    tests/real-world/dashboard-sales-analysis.xlsx
    tests/real-world/dashboard-ecommerce-sales.xlsx
)
for s in "${stages[@]}"; do required+=("tests/formula/$s.fsx"); done
for f in "${required[@]}"; do
    if [ ! -f "$f" ]; then
        echo "MISSING MANDATORY FORMULA SUITE/FIXTURE: $f" >&2
        exit 1
    fi
done
if [ "$check_coverage" -eq 1 ]; then
    echo "FORMULA COVERAGE: all mandatory stages and fixtures present"
    exit 0
fi

if [ "$skip_build" -eq 1 ]; then
    echo "USER OPT-OUT: --skip-build; fresh formula build coverage is excluded"
else
    export MSBUILDDISABLENODEREUSE=1
    export DOTNET_PROCESSOR_COUNT="${DOTNET_PROCESSOR_COUNT:-2}"
    # Office.Formats is needed by the vbaProject.bin stage.
    for proj in src/Spreadsheet/Rbt.Spreadsheet.fsproj src/Office.Formats/Rbt.OfficeFormat.fsproj; do
        echo "building $proj..."
        if ! dotnet build "$proj" -v quiet -m:1 -nr:false -p:UseSharedCompilation=false; then
            echo "BUILD FAILED: $proj" >&2
            exit 1
        fi
    done
fi

failed=0
for s in "${stages[@]}"; do
    f="tests/formula/$s.fsx"
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
