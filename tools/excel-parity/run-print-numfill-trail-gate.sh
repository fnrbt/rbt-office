#!/usr/bin/env bash
# Trailing numeric-fill print gate: suffix-empty `0.00* `, `0.00*.`, `0.00*-;0.00` against the matched
# `0.00` control, on the fresh-font/fresh-scale blind set.
#
# usage: run-print-numfill-trail-gate.sh CANDIDATE_DIST EXCEL_ROOT OUT \
#          [BASELINE_DIST] [CANDIDATE_NATIVE_DIR] [BASELINE_NATIVE_DIR]
#   CANDIDATE_DIST       frozen candidate dist (scripts/build-dist.sh)
#   EXCEL_ROOT           /var/tmp/numfill-oracle (trail-train, trail-holdout)
#   OUT                  gate output directory
#   BASELINE_DIST        frozen previous release, default /var/tmp/fill-engine/dist (b38ad676)
#   *_NATIVE_DIR         directory holding SpreadsheetPdfRunner; default the repo Debug runner
#
# Two verdicts are kept separate:
#   ABSOLUTE_0.35  candidate ours vs Excel absolute amount x0/x1 within 0.35 pt. This is reported as
#                  FAIL when the candidate inherits the pre-existing face/scale glyph-model offset;
#                  it is not the promotion gate.
#   PAIRED_RELEASE the scoped gate for this change: complete manifest keys, browser==native, unchanged
#                  N1/F8-negative no-fill controls (same text/hash/count, placement no further
#                  from Excel), no sign/text/hash regressions, all preregistered
#                  relative checks passing, and every newly worsened count/geometry cell quantified
#                  with its magnitude.
set -euo pipefail
repo="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
dist=$(realpath "$1")
excel=${2:-/var/tmp/numfill-oracle}
out=$(realpath -m "$3")
baseline=$(realpath -m "${4:-/var/tmp/fill-engine/dist}")
cand_native=$(realpath -m "${5:-$repo/tools/spreadsheet-pdf-runner/bin/Debug/net10.0}")
base_native=$(realpath -m "${6:-$repo/tools/spreadsheet-pdf-runner/bin/Debug/net10.0}")
mkdir -p "$out"
export PLAYWRIGHT="$repo/tools/browser-runner/node_modules/playwright/index.mjs"
export RB_WORD_FONT_DIR="$repo/artifacts/fonts/windows:$repo/artifacts/fonts/windows-extra"
gate_tmp=$(mktemp -d "${TMPDIR:-/var/tmp}/numfill-trail-gate.XXXXXX")
export TMPDIR="$gate_tmp"
trap 'rm -rf "$gate_tmp"' EXIT

# One font directory with both staged families (Arial from windows, Aptos Narrow from windows-extra).
fonts="$out/fonts"
mkdir -p "$fonts"
for f in "$repo"/artifacts/fonts/windows/*.ttf "$repo"/artifacts/fonts/windows-extra/*.ttf; do
  ln -sf "$(realpath "$f")" "$fonts/$(basename "$f")"
done

browser_proof () {  # DIST STEM OUTDIR
  local d=$1 stem=$2 odir=$3
  mkdir -p "$odir"
  node "$repo/tools/xlsx-browser/browser-textpad-proof.mjs" --dist "$d" \
    --fixture "$repo/tests/test-documents-excel/$stem.xlsx" --out "$odir" \
    --fonts "$fonts" --label "$stem" > "$odir/harness.json"
  while IFS=$'\t' read -r name idx key file; do
    [[ "$key" == "pdf" && "$file" == *.pdf ]] || continue
    cp "$odir/$file" "$odir/ours-$((idx + 1)).pdf"
  done < "$odir/$stem.tsv"
}

native_proof () {  # RUNNER_DIR STEM OUTDIR
  local runner=$1 stem=$2 odir=$3 i
  mkdir -p "$odir"
  for i in $(seq 0 11); do
    "$runner/SpreadsheetPdfRunner" "$repo/tests/test-documents-excel/$stem.xlsx" \
      "$odir/ours-$((i + 1)).pdf" "$i" >/dev/null
  done
}

measure () {  # CELLS EXDIR EXTSV OURS OUT
  python3 "$repo/tools/excel-parity/measure-print-numfill.py" "$repo/$1" "$2" "$3" "$4" "$5" >/dev/null
}

for pair in "synthetic-print-numfill-trail:trail-train" \
            "synthetic-print-numfill-trail-holdout:trail-holdout"; do
  stem=${pair%%:*}; set=${pair##*:}
  cells=tools/excel-parity/print-numfill-trail-cells.tsv
  face='Aptos Narrow'; size=14; scale=100
  if [[ "$set" == trail-holdout ]]; then
    cells=tools/excel-parity/print-numfill-trail-holdout-cells.tsv
    face=Arial; size=11; scale=60
  fi
  exdir="$excel/$set"; extsv="$exdir/$stem.tsv"
  browser_proof "$dist" "$stem" "$out/$stem-candidate-browser"
  native_proof "$cand_native" "$stem" "$out/$stem-candidate-native"
  browser_proof "$baseline" "$stem" "$out/$stem-baseline-browser"
  native_proof "$base_native" "$stem" "$out/$stem-baseline-native"
  measure "$cells" "$exdir" "$extsv" "$out/$stem-candidate-browser" "$out/$stem-candidate-browser-comparison.tsv"
  measure "$cells" "$exdir" "$extsv" "$out/$stem-candidate-native" "$out/$stem-candidate-native-comparison.tsv"
  measure "$cells" "$exdir" "$extsv" "$out/$stem-baseline-browser" "$out/$stem-baseline-browser-comparison.tsv"
  measure "$cells" "$exdir" "$extsv" "$out/$stem-baseline-native" "$out/$stem-baseline-native-comparison.tsv"
  python3 "$repo/tools/excel-parity/analyze-print-numfill-trail.py" check \
    "$repo/tools/excel-parity/print-numfill-trail-holdout-predictions.tsv" \
    "$out/$stem-candidate-browser-comparison.tsv" "$stem" --face "$face" --size "$size" --scale "$scale" \
    > "$out/$stem-candidate-browser-check.tsv"
  python3 "$repo/tools/excel-parity/analyze-print-numfill-trail.py" check \
    "$repo/tools/excel-parity/print-numfill-trail-holdout-predictions.tsv" \
    "$out/$stem-candidate-native-comparison.tsv" "$stem" --face "$face" --size "$size" --scale "$scale" \
    > "$out/$stem-candidate-native-check.tsv"
done

python3 - "$out" "$repo" "$excel" <<'PY'
import collections, math, os, sys
out, repo, excel_root = sys.argv[1:4]
fail_reasons = []
ABS_TOL = 0.35

def load(fn):
    rows = [l.rstrip('\n').split('\t') for l in open(fn, encoding='utf-8') if l.strip()]
    h = rows[0]
    return [dict(zip(h, r)) for r in rows[1:]]

def fnum(r, k):
    v = r.get(k, '')
    return None if v == '' else float(v)

def count(r):
    return int(r['fill_before'] or 0) + int(r['fill_after'] or 0)

def key(r):
    return (r['sheet'], r['row'], r['width'])

def split(path):
    rows = load(path)
    return ({key(r): r for r in rows if r['side'] == 'ours'},
            {key(r): r for r in rows if r['side'] == 'excel'})

def diff_text(x, y):
    return x['text'] != y['text'] or x['amount'] != y['amount']

def coorddiff(x, y, tol=1e-6):
    n = 0
    for f in ('amount_x0', 'amount_x1', 'amount_gap_left', 'amount_gap_right'):
        a, b = fnum(x, f), fnum(y, f)
        if a is not None and b is not None and abs(a - b) > tol:
            n += 1
    return n

sets = {
    'synthetic-print-numfill-trail': 'trail-train',
    'synthetic-print-numfill-trail-holdout': 'trail-holdout',
}
manifest = {
    'synthetic-print-numfill-trail': 'tools/excel-parity/print-numfill-trail-cells.tsv',
    'synthetic-print-numfill-trail-holdout': 'tools/excel-parity/print-numfill-trail-holdout-cells.tsv',
}
summary = ['set\tverdict\tdetail']
for stem, setname in sets.items():
    expected = set()
    for line in open(os.path.join(repo, manifest[stem]), encoding='utf-8'):
        if line.startswith('sheet\t') or not line.strip():
            continue
        p = line.rstrip('\n').split('\t')
        for w in p[6].split(','):
            expected.add((p[0], p[2], w))
    cmp_path = lambda who: os.path.join(out, f'{stem}-{who}-comparison.tsv')
    cn_b, ex = split(cmp_path('candidate-browser'))
    cn_n, _ = split(cmp_path('candidate-native'))
    bl_b, _ = split(cmp_path('baseline-browser'))
    bl_n, _ = split(cmp_path('baseline-native'))
    print(f'== {stem} ({setname})  expected {len(expected)}')

    # --- structural: complete keys and browser==native, both releases
    for who, ks in (('candidate-browser', set(cn_b)), ('candidate-native', set(cn_n)),
                    ('baseline-browser', set(bl_b)), ('baseline-native', set(bl_n))):
        if ks != expected:
            msg = f'{stem} {who} keys != manifest: expected {len(expected)} got {len(ks)} ' \
                  f'missing {len(expected - ks)} extra {len(ks - expected)}'
            print('  FAIL', msg); fail_reasons.append(msg)
    for rel, bs, ns in (('candidate', cn_b, cn_n), ('baseline', bl_b, bl_n)):
        if set(bs) != set(ns):
            msg = f'{stem} {rel} browser/native key sets differ {len(set(bs) ^ set(ns))}'
            print('  FAIL', msg); fail_reasons.append(msg)
            continue
        common = set(bs) & set(ns)
        tb = cb = hb = gb = 0
        for k in common:
            x, y = bs[k], ns[k]
            tb += diff_text(x, y); cb += count(x) != count(y); hb += x['hash'] != y['hash']
            gb += coorddiff(x, y)
        print(f'  {rel} browser-vs-native: text {tb} count {cb} hash {hb} coord {gb}')
        if tb or cb or hb or gb:
            msg = f'{stem} {rel} browser != native (text {tb} count {cb} hash {hb} coord {gb})'
            fail_reasons.append(msg)

    # --- ABSOLUTE 0.35 verdict (candidate vs Excel, all cells)
    abs_stats = collections.Counter()
    abs_fail_cells = []
    common = set(cn_b) & set(ex)
    for k in common:
        x, o = ex[k], cn_b[k]
        grp = 'control' if x['format_name'] == 'N1_plain_0' else 'trailing'
        if x['hash'] != o['hash']:
            abs_stats[(grp, 'hash')] += 1
        elif x['amount'] != o['amount']:
            abs_stats[(grp, 'text')] += 1
        if count(x) != count(o):
            abs_stats[(grp, 'count')] += 1
        for f in ('amount_x0', 'amount_x1'):
            a, b = fnum(x, f), fnum(o, f)
            if a is not None and b is not None and abs(a - b) > ABS_TOL:
                abs_stats[(grp, f)] += 1
                if len(abs_fail_cells) < 4000:
                    abs_fail_cells.append(f'{k[0]}:{k[1]}:{k[2]}\t{x["format_name"]}\t{x["align"]}\t{f}\t{a:.3f}\t{b:.3f}')
    with open(os.path.join(out, f'{stem}-absolute-failures.tsv'), 'w', encoding='utf-8') as f:
        f.write('cell\tformat\talign\tfield\texcel\tcandidate\n')
        f.write('\n'.join(abs_fail_cells) + ('\n' if abs_fail_cells else ''))
    abs_nonzero = sum(abs_stats.values())
    print('  ABSOLUTE_0.35', 'FAIL' if abs_nonzero else 'PASS',
          ' '.join(f'{g}/{k}={v}' for (g, k), v in sorted(abs_stats.items())) or 'clean')
    summary.append(f'{setname}\tABSOLUTE_0.35\t{"FAIL" if abs_nonzero else "PASS"} ' +
                   ' '.join(f'{g}/{k}={v}' for (g, k), v in sorted(abs_stats.items())))

    # --- preregistered relative checks on candidate browser and native
    rel_fail = 0
    for side in ('browser', 'native'):
        path = os.path.join(out, f'{stem}-candidate-{side}-check.tsv')
        for row in load(path):
            if 'test_id' not in row or row['test_id'] == 'workbook':
                continue
            if int(row['ours_failures']) or int(row['excel_failures']):
                rel_fail += 1
                print(f'  FAIL relative {side} {row["test_id"]}: '
                      f'excel {row["excel_failures"]} ours {row["ours_failures"]} {row["examples"]}')
    if rel_fail:
        fail_reasons.append(f'{stem} preregistered relative checks failed {rel_fail}')
    print('  relative checks:', 'PASS' if not rel_fail else 'FAIL')
    summary.append(f'{setname}\tRELATIVE\t{"PASS" if not rel_fail else "FAIL"}')

    # --- scoped paired gate vs baseline
    paired = collections.Counter()
    regress_rows = []
    # controls: N1 and F8 negative keep text/hash/count, and their placement may move only towards
    # Excel: an x0/x1 error that grows by more than one printer dot (0.12 pt) or newly exceeds 0.35 pt
    # fails. (A placement change such as print-inset-oracle.md moves every control on purpose.)
    for k in (set(cn_b) & set(bl_b) & set(ex)):
        x = ex[k]
        inactive = x['format_name'] == 'F8_suffix_empty' and float(x['value']) < 0
        if x['format_name'] != 'N1_plain_0' and not inactive:
            continue
        a, b = bl_b[k], cn_b[k]
        if diff_text(a, b) or a['hash'] != b['hash'] or count(a) != count(b):
            paired['control_changed'] += 1
            if len(regress_rows) < 4000:
                regress_rows.append(f'{k[0]}:{k[1]}:{k[2]}\tcontrol\tbaseline->candidate\t'
                                    f'text {a["amount"]!r}->{b["amount"]!r}\t'
                                    f'count {count(a)}->{count(b)}')
            continue
        if not coorddiff(a, b, 0.001):
            continue
        paired['control_moved'] += 1
        worse = False
        for f in ('amount_x0', 'amount_x1'):
            e, ob, oc = fnum(x, f), fnum(a, f), fnum(b, f)
            if e is None or ob is None or oc is None:
                continue
            eb, ec = abs(e - ob), abs(e - oc)
            if ec > eb + 0.12 or (ec > ABS_TOL and eb <= ABS_TOL):
                worse = True
            elif ec < eb - 1e-9:
                paired[f'control_{f[7:]}_better'] += 1
        if worse:
            paired['control_worse'] += 1
            if len(regress_rows) < 4000:
                regress_rows.append(f'{k[0]}:{k[1]}:{k[2]}\tcontrol\tplacement worse vs Excel\t'
                                    f'x0 {a["amount_x0"]}->{b["amount_x0"]} excel {x["amount_x0"]}')
    if paired['control_changed']:
        fail_reasons.append(f'{stem} no-fill controls changed text/hash/count {paired["control_changed"]}')
    if paired['control_worse']:
        fail_reasons.append(f'{stem} no-fill control placement worse vs Excel {paired["control_worse"]}')

    # no sign/text/hash regressions: baseline correct vs Excel, candidate wrong
    for k in (set(cn_b) & set(bl_b) & set(ex)):
        x, b, c = ex[k], bl_b[k], cn_b[k]
        b_ok = (b['hash'] == x['hash'] and b['amount'] == x['amount'])
        c_ok = (c['hash'] == x['hash'] and c['amount'] == x['amount'])
        if b_ok and not c_ok:
            paired['text_regressed'] += 1
            if len(regress_rows) < 4000:
                regress_rows.append(f'{k[0]}:{k[1]}:{k[2]}\t{ x["format_name"]}\tsign/text/hash regression\t'
                                    f'excel {x["amount"]!r} baseline {b["amount"]!r} candidate {c["amount"]!r}')
    if paired['text_regressed']:
        fail_reasons.append(f'{stem} sign/text/hash regressed {paired["text_regressed"]}')

    # newly worsened counts and geometry, magnitude vs baseline
    worst = {'count': 0, 'x0': 0.0, 'x1': 0.0}
    direction = collections.Counter()
    reg_rows = []
    for k in (set(cn_b) & set(bl_b) & set(ex)):
        x, b, c = ex[k], bl_b[k], cn_b[k]
        ce, be = abs(count(c) - count(x)), abs(count(b) - count(x))
        if ce > be:
            paired['count_worse'] += 1
            worst['count'] = max(worst['count'], ce - be)
        elif ce < be:
            paired['count_better'] += 1
        if count(c) != count(x):
            d = count(c) - count(x)
            direction['over_1' if d == 1 else 'under_1' if d == -1 else 'over_2+' if d > 1 else 'under_2+'] += 1
        errs = {}
        for f, short in (('amount_x0', 'x0'), ('amount_x1', 'x1')):
            a, o, ob = fnum(x, f), fnum(c, f), fnum(b, f)
            if a is None or o is None or ob is None:
                errs[short] = (None, None, None)
                continue
            ec, eb = abs(a - o), abs(a - ob)
            new_over = ec > ABS_TOL and eb <= ABS_TOL
            errs[short] = (eb, ec, new_over)
            if ec > eb + 1e-9:
                paired[f'{short}_worse'] += 1
                worst[short] = max(worst[short], ec - eb)
                if new_over:
                    paired[f'{short}_newly_over_tol'] += 1
            elif ec < eb - 1e-9:
                paired[f'{short}_better'] += 1
            if new_over:
                paired[f'{short}_newly_over_tol_cells'] += 1
        gx0 = errs['x0'][1] - errs['x0'][0] if errs['x0'][0] is not None else None
        gx1 = errs['x1'][1] - errs['x1'][0] if errs['x1'][0] is not None else None
        if ce > be or (gx0 is not None and gx0 > 1e-9) or (gx1 is not None and gx1 > 1e-9):
            def cell(v):
                return '' if v is None else f'{v:.3f}'
            reg_rows.append('\t'.join([
                k[0], k[1], k[2], x['format_name'], x['align'], x['value'],
                str(be), str(ce), cell(gx0), cell(gx1),
                cell(errs['x0'][0]), cell(errs['x0'][1]), cell(errs['x1'][0]), cell(errs['x1'][1]),
                'yes' if errs['x0'][2] or errs['x1'][2] else 'no']))
    with open(os.path.join(out, f'{stem}-paired-regressions.tsv'), 'w', encoding='utf-8') as f:
        f.write('sheet\trow\twidth\tformat\talign\tvalue\tcount_err_baseline\tcount_err_candidate\t'
                'x0_worse_pt\tx1_worse_pt\tx0_err_baseline\tx0_err_candidate\tx1_err_baseline\t'
                'x1_err_candidate\tnewly_over_0.35\n')
        f.write('\n'.join(reg_rows) + ('\n' if reg_rows else ''))
    print('  PAIRED_RELEASE:',
          ' '.join(f'{k}={v}' for k, v in sorted(paired.items())) or 'clean',
          '| count errors ' + (' '.join(f'{k}={v}' for k, v in sorted(direction.items())) or 'none'),
          '| worst count +%d, x0 +%.3f, x1 +%.3f pt' % (worst['count'], worst['x0'], worst['x1']))
    summary.append(f'{setname}\tPAIRED_RELEASE\t' +
                   (' '.join(f'{k}={v}' for k, v in sorted(paired.items())) or 'clean') +
                   ' count errors ' + (' '.join(f'{k}={v}' for k, v in sorted(direction.items())) or 'none') +
                   f' worst count +{worst["count"]} x0 +{worst["x0"]:.3f} x1 +{worst["x1"]:.3f}')
    hard = (paired['control_changed'] or paired['control_worse'] or paired['text_regressed'] or
            paired[f'x0_newly_over_tol'] or paired[f'x1_newly_over_tol'])
    if hard:
        fail_reasons.append(f'{stem} scoped paired gate failed')

print('\n== VERDICTS')
for line in summary:
    print(line.replace('\t', ' | ', 1))
if fail_reasons:
    print('\n'.join('FAIL: ' + r for r in fail_reasons))
    print('PAIRED_RELEASE_GATE FAIL')
    sys.exit(1)
print('PAIRED_RELEASE_GATE PASS')
PY
