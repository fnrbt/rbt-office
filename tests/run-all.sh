#!/usr/bin/env bash
# Repo-wide .NET test gates. Core runs by default; native is an explicit,
# separate profile. Each .fsx references this worktree's Debug build output.
#
# Usage: tests/run-all.sh [--core|--native] [--skip-build]
# --skip-build is a user opt-out of fresh build coverage, not a passing build.
set -u
cd "$(dirname "$0")/.."

profile=core
skip_build=0
profile_seen=0
for option in "$@"; do
    case "$option" in
        --core|--native)
            if [ "$profile_seen" -ne 0 ]; then
                echo "Specify exactly one test profile (--core or --native)" >&2
                exit 2
            fi
            profile="${option#--}"
            profile_seen=1
            ;;
        --skip-build)
            if [ "$skip_build" -ne 0 ]; then
                echo "Duplicate option: --skip-build" >&2
                exit 2
            fi
            skip_build=1
            ;;
        *)
            echo "Unknown option: $option (use --core, --native or --skip-build)" >&2
            exit 2
            ;;
    esac
done

core_projects=(
    src/Office.Formats/Rbt.OfficeFormat.fsproj
    src/Word/Rbt.Word.fsproj
    src/Word.Pdf/Rbt.Word.Pdf.fsproj
    src/Word.Import/Rbt.WordImport.fsproj
    src/Spreadsheet.Pdf/Rbt.Spreadsheet.Pdf.fsproj
)
native_projects=(
    word-wisp/Rbt.WordWisp.fsproj
    spreadsheet-wisp/Rbt.SpreadsheetWisp.fsproj
    tools/pdf-runner/PdfRunner.fsproj
    tools/spreadsheet-pdf-runner/SpreadsheetPdfRunner.fsproj
)

# Pin the complete formats inventory. Globbing would silently omit a deleted
# suite (or pass an unmatched literal as a "missing, skipped" suite).
formats_suites=(
    builders docx-progressive docx-reader docx-source-spans docx-writer
    xlsx-bridge xlsx-carry-over xlsx-date1904 xlsx-dynamic-arrays xlsx-lazy
    xlsx-orphans xlsx-package xlsx-parts xlsx-patched-corpus xlsx-reader
    xlsx-regenerated xlsx-rules xlsx-scan xlsx-writer xlsx-x14 xlsm-vba-project
    xml-patch xml-source zipinterop
)
word_suites=(
    textbuffer editing layout incremental-layout faces-relayout font-set pdf
    import-smoke package-io commands nested-table-split xml-scan header-footer
    table-merge-split table-columns anchored-objects drawing-order text-boxes
    group-shapes roundtrip-layout table-break-borders sections table-insert
    wrap-band list-levels drop-caps default-tab-stop note-layout table-compat
    column-balance section-end macros outline line-boxes marker-ascent endnote-start row-break rfonts-precedence footnote-block border-space
    float-wrap float-next long-word kerning hang-punct table-width embedded-fonts page-top-space float-table
    table-style-props border-across vmerge-content row-continue list-sections
    hyphen-break hyphenation justify-squeeze data-binding
    section-margins inline-break section-top-space line-baseline empty-mark heading-mark empty-lines keep-next-widow cell-line-fit cell-break
)

# These are required corpus anchors and all drawing fixtures that the formats
# writer otherwise conditionally skips. Missing corpus folders cannot reduce
# scan coverage silently; ordinary suite assertions still govern corpus size.
core_fixtures=(
    tests/test-documents/calibre-demo.docx
    tests/test-documents-wild/docx4j-getting-started.docx
    tests/test-documents-business/docx4j-toc.docx
    tests/test-documents-edge-cases/complex-tables.docx
    tests/test-documents-features/test-header-footer.docx
    tests/test-documents-online/file-sample.docx
    tests/test-documents-stress/microsoft-word-sample.docx
    tests/word/roundtrip-saved/unmodelled-content.original.docx
    tests/word/roundtrip-saved/unmodelled-content.pass1.docx
    tests/real-world/chart-zoo.xlsx
    tests/real-world/dashboard-ecommerce-sales.xlsx
    tests/real-world/dashboard-sales-analysis.xlsx
    tests/test-documents-excel/pandas-chartsheet.xlsx
    tests/test-documents-excel/phpsp-27template.xlsx
    tests/test-documents-excel/phpsp-32chartreadwrite.xlsx
    tests/test-documents-excel/poi-47504.xlsx
    tests/test-documents-excel/poi-47668.xlsx
    tests/test-documents-excel/poi-47813.xlsx
    tests/test-documents-excel/poi-49156.xlsx
    tests/test-documents-excel/synthetic-chart-column.xlsx
)
native_fixtures=(
    tests/word/native-painter/document.xml
    tests/word/native-painter/expected.json
    tests/test-documents-excel/synthetic-drawing-extras.xlsx
)

required=()
projects=()
if [ "$profile" = core ]; then
    projects=("${core_projects[@]}")
    required=(tests/formula/run-all.sh tests/print/helpers.fsx)
    for s in "${formats_suites[@]}"; do required+=("tests/office-formats/$s.fsx"); done
    for s in "${word_suites[@]}"; do required+=("tests/word/$s.fsx"); done
    required+=("${core_fixtures[@]}")
else
    projects=("${native_projects[@]}")
    required=(tests/word/native-painter.fsx tests/word/native-embedded-fonts.fsx tests/word/wisp-editorhost.fsx tests/spreadsheet/wisp-editorhost.fsx tests/spreadsheet/native-import.fsx tests/native-pdf.fsx "${native_fixtures[@]}")
    for tool in pdftotext pdfimages; do
        if ! command -v "$tool" >/dev/null; then
            echo "MISSING NATIVE PDF TOOL: $tool (install Poppler utilities)" >&2
            exit 1
        fi
    done
fi

for f in "${required[@]}"; do
    if [ ! -f "$f" ]; then
        echo "MISSING MANDATORY SUITE/FIXTURE: $f" >&2
        exit 1
    fi
done
if [ "$profile" = core ]; then
    # An added formats suite must be explicitly enrolled rather than quietly
    # being left outside the published coverage manifest.
    declare -A known_formats=()
    for s in "${formats_suites[@]}"; do known_formats["$s"]=1; done
    for f in tests/office-formats/*.fsx; do
        s="${f##*/}"
        s="${s%.fsx}"
        if [ -z "${known_formats[$s]:-}" ]; then
            echo "UNREGISTERED OFFICE-FORMATS SUITE: $f" >&2
            exit 1
        fi
    done
fi
if [ "$profile" = core ]; then
    # Check formula stages before building any project or launching a suite.
    bash tests/formula/run-all.sh --check-coverage || exit 1
fi

echo "=== $profile profile ==="
if [ "$profile" = core ]; then
    echo "EXCLUDED PROFILES: native, browser (run these gates separately)"
else
    echo "EXCLUDED PROFILES: core, browser (run these gates separately)"
fi
if [ "$skip_build" -eq 1 ]; then
    echo "USER OPT-OUT: --skip-build; fresh build coverage is excluded"
else
    export MSBUILDDISABLENODEREUSE=1
    export DOTNET_PROCESSOR_COUNT="${DOTNET_PROCESSOR_COUNT:-2}"
    for p in "${projects[@]}"; do
        echo "building $p..."
        if ! dotnet build "$p" -v quiet -m:1 -nr:false -p:UseSharedCompilation=false; then
            echo "BUILD FAILED: $p" >&2
            exit 1
        fi
    done
fi

failed=0

run_suite () {
    # run_suite <label> <fsx-file>; absence was checked before any build.
    local label="$1" f="$2"
    echo "--- $label ---"
    if ! dotnet fsi --quiet "$f"; then
        failed=1
    fi
}

if [ "$profile" = core ]; then
    echo "=== formula engine ==="
    formula_args=()
    if [ "$skip_build" -eq 1 ]; then formula_args+=(--skip-build); fi
    if ! bash tests/formula/run-all.sh "${formula_args[@]}"; then failed=1; fi

    echo "=== office formats ==="
    for s in "${formats_suites[@]}"; do
        run_suite "$s" "tests/office-formats/$s.fsx"
    done

    echo "=== word ==="
    for s in "${word_suites[@]}"; do
        run_suite "word-$s" "tests/word/$s.fsx"
    done

    echo "=== print ==="
    run_suite "print-helpers" "tests/print/helpers.fsx"
else
    if [ ! -f office-wisp/word-render/bin/Debug/net10.0/NotoSans-Regular.ttf ]; then
        echo "MISSING NATIVE PAINTER FONT: office-wisp/word-render/bin/Debug/net10.0/NotoSans-Regular.ttf" >&2
        exit 1
    fi
    run_suite "word-native-painter" "tests/word/native-painter.fsx"
    run_suite "word-native-embedded-fonts" "tests/word/native-embedded-fonts.fsx"
    run_suite "word-wisp-editorhost" "tests/word/wisp-editorhost.fsx"
    run_suite "spreadsheet-wisp-editorhost" "tests/spreadsheet/wisp-editorhost.fsx"
    run_suite "spreadsheet-native-import" "tests/spreadsheet/native-import.fsx"
    run_suite "native-pdf" "tests/native-pdf.fsx"
    echo "--- spreadsheet-wisp-executable ---"
    dotnet spreadsheet-wisp/bin/Debug/net10.0/Rbt.SpreadsheetWisp.dll --selftest || failed=1
fi


echo ""
if [ "$failed" -ne 0 ]; then
    echo "OFFICE $profile TESTS: FAILED"
    exit 1
fi
echo "OFFICE $profile TESTS: ALL MANDATORY SUITES PASSED"
