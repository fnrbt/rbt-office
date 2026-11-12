#!/usr/bin/env bash
# Run pixel-diff over a list of repros and print one row per repro.
#
# Usage:
#   tools/word-parity/measure-parity.sh
#   tools/word-parity/measure-parity.sh table-grid-dense-row-pitch heading-cambria-font-wrap
#   REPRO_DIR=tests/repros/word tools/word-parity/measure-parity.sh
#
# For each repro <name>:
# - Word reference PDF rendered through winvm (cached under artifacts/word/)
# - Our PDF rendered through tools/pdf-runner
# - pdftoppm both at 144 DPI
# - Pixel-diff PDF↔PDF over the whole page, threshold 32, both raw and ink-XOR
#
# Re-run any time as a regression guard after sweeping layout changes.

set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

repro_dir="${REPRO_DIR:-tests/repros/word}"
out_dir="${OUT_DIR:-artifacts/parity-measure}"
mkdir -p "$out_dir"

repros=("$@")
if [ ${#repros[@]} -eq 0 ]; then
  # Default: every .docx in tests/repros/word.
  mapfile -t repros < <(find "$repro_dir" -maxdepth 1 -name '*.docx' -printf '%f\n' | sed 's/\.docx$//' | sort)
fi

printf "%-44s %-22s %-22s\n" "repro" "raw (threshold 32)" "ink-XOR (% union)"
printf "%-44s %-22s %-22s\n" "----" "----" "----"

for name in "${repros[@]}"; do
  docx="$repro_dir/$name.docx"
  if [ ! -f "$docx" ]; then
    printf "%-44s %s\n" "$name" "MISSING docx"
    continue
  fi

  word_pdf="$out_dir/$name-word.pdf"
  word_png="$out_dir/$name-word-1.png"
  ours_pdf="$out_dir/$name-ours.pdf"
  ours_png="$out_dir/$name-ours-1.png"

  if [ ! -f "$word_png" ]; then
    tools/word-parity/word-print-pdf.sh "$docx" "$word_pdf" >/dev/null 2>&1
    pdftoppm -r 144 "$word_pdf" "$out_dir/$name-word" -png -f 1 -l 1 >/dev/null 2>&1
  fi

  dotnet run --project tools/pdf-runner/PdfRunner.fsproj -- "$docx" "$ours_pdf" >/dev/null 2>&1
  pdftoppm -r 144 "$ours_pdf" "$out_dir/$name-ours" -png -f 1 -l 1 >/dev/null 2>&1 || true

  if [ ! -f "$ours_png" ] || [ ! -f "$word_png" ]; then
    printf "%-44s %s\n" "$name" "RENDER FAILED"
    continue
  fi

  output=$(python3 tools/word-parity/pixel-diff.py "$word_png" "$ours_png" --threshold 32 2>&1)
  raw=$(echo "$output" | grep "raw pixels differing:" | awk '{print $4, $5}')
  ink=$(echo "$output" | grep "% of union ink:" | awk '{print $5}')
  printf "%-44s %-22s %-22s\n" "$name" "${raw:-?}" "${ink:-?}"
done
