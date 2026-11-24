#!/usr/bin/env bash
# Numeric-fill print gate. Run only after the engine source freeze and a single dist build.
# usage: run-print-numfill-gate.sh DIST EXCEL_ROOT OUT
#   DIST         built dist (scripts/build-dist.sh)
#   EXCEL_ROOT   /var/tmp/numfill-oracle (train, holdout-georgia, holdout-tahoma)
#   OUT          gate output directory
set -euo pipefail
repo="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
dist=$(realpath "$1")
excel=${2:-/var/tmp/numfill-oracle}
out=$(realpath -m "$3")
mkdir -p "$out"
export PLAYWRIGHT="$repo/tools/browser-runner/node_modules/playwright/index.mjs"
export RB_WORD_FONT_DIR="$repo/artifacts/fonts/windows:$repo/artifacts/fonts/windows-extra"

run_set () {
  local stem=$1 cells=$2 exdir=$3 extsv=$4
  local browser="$out/$stem-browser" native="$out/$stem-native" browser_as="$out/$stem-browser-pdfs"
  mkdir -p "$browser" "$native" "$browser_as"
  node "$repo/tools/xlsx-browser/browser-count-proof.mjs" --dist "$dist" \
    --fixture "$repo/tests/test-documents-excel/$stem.xlsx" --out "$browser" \
    --fonts "$repo/artifacts/fonts/windows-extra" --label "$stem" > "$browser/harness.json"
  while IFS=$'\t' read -r name idx key file; do
    [[ "$key" == "pdf" && "$file" == *.pdf ]] || continue
    cp "$browser/$file" "$browser_as/ours-$((idx + 1)).pdf"
  done < "$browser/$stem.tsv"
  for i in 0 1 2 3 4 5; do
    "$repo/tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner" \
      "$repo/tests/test-documents-excel/$stem.xlsx" "$native/ours-$((i + 1)).pdf" "$i" >/dev/null
  done
  python3 "$repo/tools/excel-parity/measure-print-numfill.py" "$repo/$cells" "$exdir" "$extsv" \
    "$browser_as" "$out/$stem-browser-comparison.tsv" >/dev/null
  python3 "$repo/tools/excel-parity/measure-print-numfill.py" "$repo/$cells" "$exdir" "$extsv" \
    "$native" "$out/$stem-native-comparison.tsv" >/dev/null
  python3 "$repo/tools/excel-parity/analyze-print-numfill.py" \
    "$repo/tools/excel-parity/print-numfill-holdout-predictions.tsv" \
    "$out/$stem-browser-comparison.tsv" "$stem" > "$out/$stem-browser-results.tsv"
  python3 "$repo/tools/excel-parity/analyze-print-numfill.py" \
    "$repo/tools/excel-parity/print-numfill-holdout-predictions.tsv" \
    "$out/$stem-native-comparison.tsv" "$stem" > "$out/$stem-native-results.tsv"
}

run_set synthetic-print-numfill tools/excel-parity/print-numfill-cells.tsv \
  "$excel/train" "$excel/train/synthetic-print-numfill.tsv"
run_set synthetic-print-numfill-holdout-georgia tools/excel-parity/print-numfill-holdout-georgia-cells.tsv \
  "$excel/holdout-georgia" "$excel/holdout-georgia/synthetic-print-numfill-holdout-georgia.tsv"
run_set synthetic-print-numfill-holdout-tahoma tools/excel-parity/print-numfill-holdout-tahoma-cells.tsv \
  "$excel/holdout-tahoma" "$excel/holdout-tahoma/synthetic-print-numfill-holdout-tahoma.tsv"

python3 - "$out" <<'PY'
import collections, os, sys
out = sys.argv[1]

def load(fn):
    rows = [l.rstrip('\n').split('\t') for l in open(fn)]
    h = rows[0]
    return [dict(zip(h, r)) for r in rows[1:]]

for stem in ('synthetic-print-numfill', 'synthetic-print-numfill-holdout-georgia',
             'synthetic-print-numfill-holdout-tahoma'):
    b = load(os.path.join(out, stem + '-browser-comparison.tsv'))
    n = load(os.path.join(out, stem + '-native-comparison.tsv'))
    key = lambda r: (r['sheet'], r['row'], r['width'], r['side'])
    bt = {key(r): r for r in b}
    nt = {key(r): r for r in n}
    maxgap = 0.0
    textdiff = 0
    countdiff = 0
    for k in set(bt) & set(nt):
        x, y = bt[k], nt[k]
        if x['text'] != y['text']:
            textdiff += 1
        if x['fill_before'] != y['fill_before'] or x['fill_after'] != y['fill_after']:
            countdiff += 1
        for f in ('amount_x0', 'amount_x1', 'dollar_x0', 'dollar_x1'):
            if x[f] and y[f]:
                maxgap = max(maxgap, abs(float(x[f]) - float(y[f])))
    print(f'{stem}: browser-vs-native operator max {maxgap:.4f} pt, text differences {textdiff}, count differences {countdiff}')
PY
