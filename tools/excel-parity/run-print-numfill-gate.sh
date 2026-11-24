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
gate_tmp=$(mktemp -d "${TMPDIR:-/var/tmp}/numfill-gate.XXXXXX")
export TMPDIR="$gate_tmp"
trap 'rm -rf "$gate_tmp"' EXIT

run_set () {
  local stem=$1 cells=$2 exdir=$3 extsv=$4
  local browser="$out/$stem-browser" native="$out/$stem-native" browser_as="$out/$stem-browser-pdfs"
  mkdir -p "$browser" "$native" "$browser_as"
  node "$repo/tools/xlsx-browser/browser-textpad-proof.mjs" --dist "$dist" \
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

python3 - "$out" "$repo" <<'PY'
import os, sys
out = sys.argv[1]
repo_root = sys.argv[2]
fail = False

def load(fn):
    rows = [l.rstrip('\n').split('\t') for l in open(fn) if l.strip()]
    h = rows[0]
    return [dict(zip(h, r)) for r in rows[1:]]

manifests = {
    'synthetic-print-numfill': 'tools/excel-parity/print-numfill-cells.tsv',
    'synthetic-print-numfill-holdout-georgia': 'tools/excel-parity/print-numfill-holdout-georgia-cells.tsv',
    'synthetic-print-numfill-holdout-tahoma': 'tools/excel-parity/print-numfill-holdout-tahoma-cells.tsv',
}
for stem in manifests:
    expected = set()
    for line in open(os.path.join(repo_root, manifests[stem]), encoding='utf-8'):
        if line.startswith('sheet\t') or not line.strip():
            continue
        p = line.rstrip('\n').split('\t')
        for w in p[6].split(','):
            expected.add((p[0], p[2], w))
    # Compare only the OUR sides. The Excel reference rows are common and must not pad the key set.
    b = [r for r in load(os.path.join(out, stem + '-browser-comparison.tsv')) if r['side'] == 'ours']
    n = [r for r in load(os.path.join(out, stem + '-native-comparison.tsv')) if r['side'] == 'ours']
    key = lambda r: (r['sheet'], r['row'], r['width'])
    bt = {key(r): r for r in b}
    nt = {key(r): r for r in n}
    if not bt or not nt:
        print(f'{stem}: FAIL empty ours key set browser={len(bt)} native={len(nt)}')
        fail = True
        continue
    for who, ks in (('browser', set(bt)), ('native', set(nt))):
        if ks != expected:
            print(f'{stem}: FAIL {who} ours keys != oracle truth: expected {len(expected)}, '
                  f'got {len(ks)}, missing {len(expected - ks)}, extra {len(ks - expected)}')
            fail = True
    if set(bt) != set(nt):
        print(f'{stem}: FAIL ours key sets differ browser={len(bt)} native={len(nt)} '
              f'symmetric difference={len(set(bt) ^ set(nt))}')
        fail = True
        continue
    if set(bt) != expected:
        continue
    textdiff = countdiff = coorddiff = 0
    maxgap = 0.0
    for k in bt:
        x, y = bt[k], nt[k]
        if x['text'] != y['text']:
            textdiff += 1
        if x['fill_before'] != y['fill_before'] or x['fill_after'] != y['fill_after']:
            countdiff += 1
        for f in ('amount_x0', 'amount_x1', 'dollar_x0', 'dollar_x1',
                  'amount_gap_left', 'amount_gap_right', 'dollar_gap_left'):
            if x[f] and y[f]:
                d = abs(float(x[f]) - float(y[f]))
                maxgap = max(maxgap, d)
                if d > 1e-6:
                    coorddiff += 1
    ok = textdiff == 0 and countdiff == 0 and coorddiff == 0
    fail = fail or not ok
    print(f'{stem}: ours rows {len(bt)}, browser-vs-native max {maxgap:.5f} pt, '
          f'text diffs {textdiff}, count diffs {countdiff}, coord diffs {coorddiff}, '
          f'{"PASS" if ok else "FAIL"}')
print('GATE', 'FAIL' if fail else 'PASS')
sys.exit(1 if fail else 0)
PY
