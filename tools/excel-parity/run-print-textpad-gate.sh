#!/usr/bin/env bash
# Text-pad Chromium gate. Run only after the source freeze and a single dist build.
# usage: run-textpad-gate.sh DIST EXCEL_ROOT OUT
#   DIST         built dist (scripts/build-dist.sh)
#   EXCEL_ROOT   /var/tmp/textpad-oracle (train-excel, holdout-*-excel, escapes-excel)
#   OUT          gate output directory
set -euo pipefail
repo="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
dist=$(realpath "$1")
excel=${2:-/var/tmp/textpad-oracle}
out=$(realpath -m "$3")
mkdir -p "$out"
export PLAYWRIGHT="$repo/tools/browser-runner/node_modules/playwright/index.mjs"
export RB_WORD_FONT_DIR="$repo/artifacts/fonts/windows:$repo/artifacts/fonts/windows-extra"

run_set () {
  local stem=$1 cells=$2 exdir=$3 extsv=$4
  local browser="$out/$stem-browser" native="$out/$stem-native" browser_as="$out/$stem-browser-pdfs"
  mkdir -p "$browser" "$native" "$browser_as"
  node "$repo/tools/xlsx-browser/browser-textpad-proof.mjs" --dist "$dist" \
    --fixture "$repo/tests/test-documents-excel/$stem.xlsx" --out "$browser" \
    --fonts "$repo/artifacts/fonts/windows-extra" --label "$stem" > "$browser/harness.json"
  for pdf in "$browser"/"$stem"-sheet*.pdf; do
    idx=$(basename "$pdf" | sed -E 's/.*-sheet([0-9]+)\.pdf/\1/')
    cp "$pdf" "$browser_as/ours-$((idx + 1)).pdf"
  done
  for i in 0 1 2; do
    "$repo/tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner" \
      "$repo/tests/test-documents-excel/$stem.xlsx" "$native/ours-$((i + 1)).pdf" "$i" >/dev/null
  done
  python3 "$repo/tools/excel-parity/measure-print-textpad.py" "$repo/$cells" "$exdir" "$extsv" "$browser_as" "$out/$stem-browser-comparison.tsv" >/dev/null
  python3 "$repo/tools/excel-parity/measure-print-textpad.py" "$repo/$cells" "$exdir" "$extsv" "$native" "$out/$stem-native-comparison.tsv" >/dev/null
}

run_set synthetic-print-textpad tools/excel-parity/print-textpad-cells.tsv "$excel/train-excel" "$excel/train-excel/synthetic-print-textpad.tsv"
run_set synthetic-print-textpad-holdout-georgia tools/excel-parity/print-textpad-holdout-georgia-cells.tsv "$excel/holdout-georgia-excel" "$excel/holdout-georgia-excel/synthetic-print-textpad-holdout-georgia.tsv"
run_set synthetic-print-textpad-holdout-tahoma tools/excel-parity/print-textpad-holdout-tahoma-cells.tsv "$excel/holdout-tahoma-excel" "$excel/holdout-tahoma-excel/synthetic-print-textpad-holdout-tahoma.tsv"
run_set synthetic-print-textpad-escapes tools/excel-parity/print-textpad-escapes-cells.tsv "$excel/escapes-excel" "$excel/escapes-excel/synthetic-print-textpad-escapes.tsv"

python3 - "$out" <<'PY'
import collections, glob, os, sys
out = sys.argv[1]
def load(fn):
    rows=[l.rstrip('\n').split('\t') for l in open(fn)][1:]
    return {(r[0],int(r[2]),r[3],r[4]):r for r in rows}
for fn in sorted(glob.glob(os.path.join(out,'*-comparison.tsv'))):
    rows=[l.rstrip('\n').split('\t') for l in open(fn)][1:]
    t=collections.defaultdict(dict)
    for r in rows:
        sheet,idx,row,kind,name,fmt,val,side,el,er,text,xmin,xmax,gl,gr=r
        t[(sheet,int(row),kind,name)][side]=(float(xmin),float(xmax),float(gl),float(gr),text)
    bad=[]; mx=0.0; n=0
    for k,v in t.items():
        if k[2]!='text': continue
        e,o=v.get('excel'),v.get('ours')
        if not e or not o: continue
        n += 1
        d=max(abs(o[0]-e[0]),abs(o[1]-e[1])); mx=max(mx,d)
        if d>1.0 or o[4]!=e[4]: bad.append((k,d,e[4],o[4]))
    print(f'{os.path.basename(fn)}: text rows {n}, max |delta| {mx:.3f} pt, text/position failures {len(bad)}')
    for b in bad[:10]: print('   ',b)
# browser vs native
for stem in ('synthetic-print-textpad','synthetic-print-textpad-holdout-georgia','synthetic-print-textpad-holdout-tahoma','synthetic-print-textpad-escapes'):
    b=load(os.path.join(out,stem+'-browser-comparison.tsv')); n=load(os.path.join(out,stem+'-native-comparison.tsv'))
    mx=0.0; texts=[]
    for k in set(b)|set(n):
        if k not in b or k not in n: continue
        br,nr=b[k],n[k]
        mx=max(mx,abs(float(br[11])-float(nr[11])),abs(float(br[12])-float(nr[12])))
        if br[10]!=nr[10]: texts.append((k,br[10],nr[10]))
    print(f'{stem}: browser vs native max word-box delta {mx:.4f} pt, text differences {len(texts)}')
PY
