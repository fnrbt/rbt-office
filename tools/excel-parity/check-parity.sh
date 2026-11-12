#!/usr/bin/env bash
# Excel-parity regression gate. Runs:
#   1. VBA-layout harness across every fixture with an Excel COM TSV in
#      /tmp/xfix/vba/ at 0.5 pt tolerance. Fails on any field mismatch.
#   2. Print-PDF MAE thresholds across fixtures with a saved Excel
#      print PDF reference, fails when body/full MAE exceeds the gate.
#   3. Print-helper tests + full fixture smoke render.
#
# Designed to be runnable from CI: prints a summary, exits nonzero on any
# breach.
#
# Skips silently when no TSVs / reference PDFs are present — useful when
# winvm is unavailable on the CI runner.
set -uo pipefail

cd "$(dirname "$0")/../.."
ROOT=$(pwd)
EXIT_CODE=0

echo "=== Excel parity gate ==="
echo

# 1. VBA-layout harness.
TSV_DIR=${VBA_TSV_DIR:-/tmp/xfix/vba}
OURS_TSV_DIR=${VBA_OURS_TSV_DIR:-tmp/excel-parity-vba}
if compgen -G "$TSV_DIR/*-excel.tsv" >/dev/null; then
  echo "[1/3] VBA-layout harness ($TSV_DIR)"
  mkdir -p "$OURS_TSV_DIR"
  layout_fail=0
  layout_pass=0
  while IFS= read -r tsv; do
    name=$(basename "$tsv" -excel.tsv)
    if [[ ! -f "tests/test-documents-excel/$name.xlsx" ]]; then
      continue
    fi
    ours_tsv="$OURS_TSV_DIR/$name-ours.tsv"
    if ! tools/excel-parity/spreadsheet-vba-layout-tsv.sh \
        "tests/test-documents-excel/$name.xlsx" \
        "$ours_tsv" >/dev/null 2>&1; then
      echo "  FAIL: $name (layout tool errored)"
      layout_fail=$((layout_fail+1))
      EXIT_CODE=1
      continue
    fi
    if python3 tools/excel-parity/compare-vba-layout.py \
        "$TSV_DIR/$name-excel.tsv" \
        "$ours_tsv" \
        --tol 0.5 >/dev/null 2>&1; then
      layout_pass=$((layout_pass+1))
    else
      echo "  FAIL: $name (field mismatches)"
      layout_fail=$((layout_fail+1))
      EXIT_CODE=1
    fi
  done < <(find "$TSV_DIR" -name '*-excel.tsv' | sort)
  echo "  $layout_pass/$((layout_pass + layout_fail)) fixtures clean"
  echo
else
  echo "[1/3] VBA-layout harness — SKIPPED (no TSVs in $TSV_DIR)"
  echo
fi

# 2. Print-PDF MAE thresholds. Only run for fixtures with a saved Excel
#    reference; --reuse-excel skips the winvm extraction.
PDF_REF_DIR=artifacts/pdf-compare-excel
if [[ -d "$PDF_REF_DIR" ]]; then
  echo "[2/3] Print-PDF MAE thresholds"
  print_fixtures=()
  for dir in "$PDF_REF_DIR"/tests_test-documents-excel_*; do
    [[ -f "$dir/excel.pdf" ]] || continue
    name=$(basename "$dir" | sed 's|tests_test-documents-excel_||')
    if [[ -f "tests/test-documents-excel/$name.xlsx" ]]; then
      print_fixtures+=("tests/test-documents-excel/$name.xlsx")
    fi
  done
  if [[ ${#print_fixtures[@]} -gt 0 ]]; then
    # Thresholds calibrated against current renderer output + 10% headroom.
    # Tighten as parity improves; loosen only with explicit justification.
    if ! python3 tools/excel-parity/compare-print-pdfs.py \
        "${print_fixtures[@]}" \
        --reuse-excel \
        --max-body-mae 10 \
        --max-full-mae 10 \
        >/tmp/print-mae.log 2>&1; then
      echo "  FAIL: MAE threshold breached"
      tail -20 /tmp/print-mae.log | sed 's/^/    /'
      EXIT_CODE=1
    else
      echo "  ${#print_fixtures[@]} fixtures within MAE thresholds"
    fi
  else
    echo "  SKIPPED (no Excel print PDFs in $PDF_REF_DIR)"
  fi
  echo
else
  echo "[2/3] Print-PDF MAE thresholds — SKIPPED (no $PDF_REF_DIR)"
  echo
fi

# 3. Print-helper tests + smoke render.
echo "[3/3] Print-helper tests + smoke render"
if dotnet fsi tests/print/helpers.fsx 2>&1 | grep -q 'Print-helper tests: .* passed, 0 failed'; then
  echo "  Print helpers: passed"
else
  echo "  FAIL: Print helper tests"
  EXIT_CODE=1
fi

if [[ "${SKIP_SMOKE:-0}" != "1" ]]; then
  SMOKE_DIR=${SMOKE_DIR:-tmp/parity-smoke}
  mkdir -p "$SMOKE_DIR"
  smoke_fail=0
  smoke_ok=0
  for f in tests/test-documents-excel/*.xlsx; do
    name=$(basename "$f" .xlsx)
    if dotnet run --project tools/spreadsheet-pdf-runner --no-build -- \
        "$f" "$SMOKE_DIR/$name.pdf" >/dev/null 2>&1; then
      smoke_ok=$((smoke_ok+1))
    else
      echo "  FAIL: smoke render $name"
      smoke_fail=$((smoke_fail+1))
      EXIT_CODE=1
    fi
  done
  rm -f "$SMOKE_DIR"/*.pdf
  echo "  Smoke render: $smoke_ok/$((smoke_ok + smoke_fail)) fixtures"
fi

echo
if [[ $EXIT_CODE -eq 0 ]]; then
  echo "=== PASS ==="
else
  echo "=== FAIL ==="
fi
exit $EXIT_CODE
