#!/usr/bin/env python3
"""Trailing `*x` fill (suffix-empty formats) oracle: training model and locked checks.

    python3 tools/excel-parity/analyze-print-numfill-trail.py train COMPARISON.tsv PREFIX [--face F --size S --scale N]
    python3 tools/excel-parity/analyze-print-numfill-trail.py check PREDICTIONS.tsv COMPARISON.tsv WORKBOOK \
        [--face F --size S --scale N]

COMPARISON.tsv is measure-print-numfill.py's long table. `train` writes PREFIX-cells.tsv (per-cell
truth) and PREFIX-transitions.tsv (count transitions with edge/residual deltas) and prints the model
summary. `check` evaluates the locked relations on Excel first (the model must describe Excel) and
then on ours, and counts per-cell ours-vs-Excel disagreements.

Trailing formats are `0.00* ` (space), `0.00*.` (dot) and `0.00*-;0.00` (dash on the positive/zero
section only), with `0.00` as the matched no-fill control. Counts come from emitted text order, so
the first fill glyph is counted even when its origin overlaps the number's advance box.
"""
import collections
import sys

ALIGN_SHEET = {'left': None, 'center': None, 'right': None}
FILL_LIMIT = {' ': 1, '.': 2, '-': 1}
FILL_NAME = {' ': 'space', '.': 'dot', '-': 'dash'}
TOL = 0.35


def load(path):
    lines = [l.rstrip('\n') for l in open(path, encoding='utf-8') if l.strip()]
    head = lines[0].split('\t')
    return [dict(zip(head, l.split('\t'))) for l in lines[1:]]


def fnum(r, key):
    v = r.get(key, '')
    return None if v in ('', None) else float(v)


def inum(r, key):
    v = r.get(key, '')
    return 0 if v in ('', None) else int(v)


def value_name(r):
    v = float(r['value'])
    return 'pos' if v > 0 else ('neg' if v < 0 else 'zero')


def count(r):
    return inum(r, 'fill_before') + inum(r, 'fill_after')


def by_key(rows, side):
    return {(r['sheet'], r['row'], r['width']): r for r in rows if r['side'] == side}


def index_controls(rows):
    out = {}
    for r in rows:
        if r['format_name'] == 'N1_plain_0':
            out[(r['side'], r['sheet'], r['value'], r['width'])] = r
    return out


def predicted_adv(face, size, scale, ch):
    """The existing 600-dpi printer-advance model used by the count oracle."""
    import importlib.util
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    spec = importlib.util.spec_from_file_location('countpredict', os.path.join(here, 'predict-print-count.py'))
    cp = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cp)
    key = (face, False)
    if key not in cp.FONT_FILES:
        return None
    g, _, _ = cp.font_metrics(os.path.join(here, '..', '..', cp.FONT_DIR, cp.FONT_FILES[key]))
    dots = round(size * scale / 100.0 * 600.0 / 72.0)
    return round(g[ch] * dots) * 72.0 / 600.0


def summarise(cells, face, size, scale):
    """Print the training model: block shift slopes, residuals, transitions, count law."""
    adv = {ch: predicted_adv(face, size, scale, ch) for ch in (' ', '.', '-')}
    print('predicted advances', {k: (None if v is None else round(v, 3)) for k, v in adv.items()})
    for align in ('left', 'center', 'right'):
        ac = [r for r in cells if r['align'] == align]
        if not ac:
            continue
        nz = [r for r in ac if r['count'] > 0]
        zero = [r for r in ac if r['count'] == 0]
        dx0 = [abs(r['x0'] - r['ref_x0']) for r in nz if r['x0'] is not None and r['ref_x0'] is not None]
        dx1 = [abs(r['x1'] - r['ref_x1']) for r in nz if r['x1'] is not None and r['ref_x1'] is not None]
        print(f'== {align}: {len(ac)} cells, {len(nz)} with fill')
        print(f'   count 0 vs N1 |dx0|max {max((abs(r["x0"]-r["ref_x0"]) for r in zero if r["x0"] is not None and r["ref_x0"] is not None), default=float("nan")):.3f}'
              f'  |dx1|max {max((abs(r["x1"]-r["ref_x1"]) for r in zero if r["x1"] is not None and r["ref_x1"] is not None), default=float("nan")):.3f}')
        if dx0:
            print(f'   count>0 |dx0| max {max(dx0):.3f}  |dx1| max {max(dx1):.3f}')
        for ch in (' ', '.', '-'):
            ac2 = [r for r in ac if r['fill'] == ch and r['count'] > 0 and r['adv']]
            if not ac2:
                continue
            k = {'left': 0.0, 'center': 0.5, 'right': 1.0}[align]
            res = [(r['ref_x0'] - r['x0']) - k * r['count'] * r['adv'] for r in ac2]
            print(f'   {FILL_NAME[ch]:5s} n={len(ac2):3d} shift-k*count*adv min {min(res):+.3f} max {max(res):+.3f}'
                  f'  adv_meas {min(r["adv"] for r in ac2):.3f}..{max(r["adv"] for r in ac2):.3f}'
                  f'  adv_pred {adv[ch] if adv[ch] is not None else float("nan"):.3f}')
            resid = [r['cell_right'] - r['composite'] for r in ac2]
            print(f'         residual(cell_right-composite) {min(resid):.3f}..{max(resid):.3f}')
    # transition table and monotonicity
    groups = collections.defaultdict(list)
    for r in cells:
        groups[(r['align'], r['format'], r['value'])].append(r)
    bad = []
    for key, rr in groups.items():
        rr.sort(key=lambda r: float(r['width']))
        for a, b in zip(rr, rr[1:]):
            if b['count'] < a['count']:
                bad.append(f'{key} {a["width"]}->{b["width"]} count {a["count"]}->{b["count"]}')
            lim = FILL_LIMIT.get(b['fill'], 1)
            if b['count'] - a['count'] > lim:
                bad.append(f'{key} {a["width"]}->{b["width"]} count step {a["count"]}->{b["count"]} > {lim}')
    print('monotonicity/step violations:', len(bad))
    for x in bad[:10]:
        print('  ', x)
    # hash equivalence and sign
    hb = sd = 0
    for r in cells:
        if r['ref_hash'] != r['hash']:
            hb += 1
            if hb <= 5:
                print('  hash mismatch', r['sheet'], r['row'], r['width'], r['hash'], r['ref_hash'])
        if r['format'] == 'F8_suffix_empty' and r['value'] == 'neg':
            continue
        neg = r['value'] == 'neg'
        if r['amount'] and neg != r['amount'].startswith('-'):
            sd += 1
    print('hash-vs-N1 mismatches:', hb, ' sign mismatches:', sd)


def build_cells(rows):
    excel = by_key(rows, 'excel')
    ctrl = index_controls(rows)
    out = []
    for r in rows:
        bg = r['format_name']
        if bg not in ('S2_space_after', 'D2_dot_after', 'F8_suffix_empty'):
            continue
        ctrl_row = ctrl.get((r['side'], r['sheet'], r['value'], r['width']))
        abs_row = ctrl.get((r['side'], r['sheet'], '1234.5', r['width']))
        d = {
            'side': r['side'], 'sheet': r['sheet'], 'row': r['row'], 'width': r['width'],
            'align': r['align'], 'format': bg, 'value': value_name(r), 'fill': r['fill_char'],
            'hash': r['hash'] == 'yes', 'text': r['text'], 'amount': r['amount'],
            'x0': fnum(r, 'amount_x0'), 'x1': fnum(r, 'amount_x1'),
            'cell_left': fnum(r, 'cell_left'), 'cell_right': fnum(r, 'cell_right'),
            'count': count(r), 'composite': fnum(r, 'composite_x1'), 'adv': fnum(r, 'fill_adv'),
            'ctrl_hash': None if ctrl_row is None else ctrl_row['hash'] == 'yes',
            'nx0': None if ctrl_row is None else fnum(ctrl_row, 'amount_x0'),
            'nx1': None if ctrl_row is None else fnum(ctrl_row, 'amount_x1'),
            'actrl_hash': None if abs_row is None else abs_row['hash'] == 'yes',
            'ax0': None if abs_row is None else fnum(abs_row, 'amount_x0'),
            'ax1': None if abs_row is None else fnum(abs_row, 'amount_x1'),
        }
        # The negative section of a multi-section format prints the absolute value; `0.00*-;0.00`
        # therefore has no minus, and its fit/hash/geometry control is the positive `0.00` cell.
        if bg == 'F8_suffix_empty' and d['value'] == 'neg':
            d['ref_hash'], d['ref_x0'], d['ref_x1'] = d['actrl_hash'], d['ax0'], d['ax1']
        else:
            d['ref_hash'], d['ref_x0'], d['ref_x1'] = d['ctrl_hash'], d['nx0'], d['nx1']
        out.append(d)
    return out


def write_tables(cells, prefix):
    cols = ['side', 'sheet', 'row', 'width', 'align', 'format', 'value', 'fill', 'hash', 'text',
            'amount', 'x0', 'x1', 'cell_left', 'cell_right', 'count', 'composite', 'residual',
            'adv', 'ctrl_hash', 'nx0', 'nx1', 'ref_hash', 'ref_x0', 'ref_x1']
    with open(prefix + '-cells.tsv', 'w', encoding='utf-8') as f:
        f.write('\t'.join(cols) + '\n')
        for r in cells:
            d = dict(r)
            d['residual'] = '' if r['composite'] is None or r['cell_right'] is None else f'{r["cell_right"] - r["composite"]:.3f}'
            f.write('\t'.join('' if d[c] is None else (f'{d[c]:.3f}' if isinstance(d[c], float) else str(d[c])) for c in cols) + '\n')
    groups = collections.defaultdict(list)
    for r in cells:
        groups[(r['align'], r['format'], r['value'])].append(r)
    with open(prefix + '-transitions.tsv', 'w', encoding='utf-8') as f:
        f.write('align\tformat\tvalue\tfill\tfrom_w\tto_w\tn0\tn1\tdW\tdx0\tdx1\tdcomposite\tresid0\tresid1\tadv\n')
        for key, rr in sorted(groups.items()):
            rr.sort(key=lambda r: float(r['width']))
            for a, b in zip(rr, rr[1:]):
                if a['count'] == b['count']:
                    continue
                # Cell-relative positions: bands are separate sheets, so absolute page x jumps at a
                # band boundary and is meaningless in a width sequence.
                def rel(r, field):
                    v, l = r[field], r['cell_left']
                    return None if v is None or l is None else v - l
                def dd(field):
                    x, y = rel(a, field), rel(b, field)
                    return '' if x is None or y is None else f'{y - x:+.3f}'
                wa = None if a['cell_right'] is None else a['cell_right'] - a['cell_left']
                wb = None if b['cell_right'] is None else b['cell_right'] - b['cell_left']
                f.write('\t'.join(map(str, [
                    key[0], key[1], key[2], FILL_NAME.get(b['fill'], b['fill']),
                    a['width'], b['width'], a['count'], b['count'],
                    '' if wa is None or wb is None else f'{wb - wa:+.3f}',
                    dd('x0'), dd('x1'), dd('composite'),
                    '' if a['composite'] is None or a['cell_right'] is None else f'{a["cell_right"]-a["composite"]:+.3f}',
                    '' if b['composite'] is None or b['cell_right'] is None else f'{b["cell_right"]-b["composite"]:+.3f}',
                    '' if b['adv'] is None else f'{b["adv"]:.3f}',
                ])) + '\n')


def check(pred_path, comp_path, workbook, face, size, scale):
    rows = load(comp_path)
    cells = build_cells(rows)
    preds = [l.rstrip('\n').split('\t') for l in open(pred_path, encoding='utf-8')
             if not l.startswith('#') and l.strip()][1:]
    out = ['workbook\ttest_id\tcells\texcel_failures\tours_failures\texpected\texamples']
    by_side = collections.defaultdict(list)
    for c in cells:
        by_side[c['side']].append(c)
    # No-fill control measurements for the content-width law: left inset from the left sheet and
    # right inset from the right sheet of the same width band (identical printed column width).
    n1 = {}
    for r in rows:
        if r['format_name'] == 'N1_plain_0' and r['amount']:
            n1[(r['side'], r['sheet'], r['value'], r['width'])] = r
    for p in preds:
        test_id, group, desc, formats, values, expected, accept = p
        fmts = formats.split(',')
        vals = values.split(',')
        res = {}
        for side in ('excel', 'ours'):
            cc = [c for c in by_side[side] if c['format'] in fmts and c['value'] in vals]
            fails = []
            if test_id in ('TRAIL_LEFT_EDGE', 'TRAIL_CENTER_SHIFT', 'TRAIL_RIGHT_SHIFT', 'TRAIL_TEXT',
                           'TRAIL_COUNT0', 'TRAIL_INSIDE', 'TRAIL_RESIDUAL', 'TRAIL_ADV'):
                for c in cc:
                    if c['hash'] or c['ref_hash']:
                        continue
                    if test_id == 'TRAIL_LEFT_EDGE' and c['align'] == 'left':
                        if c['x0'] is None or c['ref_x0'] is None or abs(c['x0'] - c['ref_x0']) > TOL \
                                or c['x1'] is None or c['ref_x1'] is None or abs(c['x1'] - c['ref_x1']) > TOL:
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]}')
                    elif test_id == 'TRAIL_CENTER_SHIFT' and c['align'] == 'center':
                        if c['x0'] is None or c['ref_x0'] is None or not c['adv']:
                            continue
                        pred = c['ref_x0'] - c['count'] * c['adv'] / 2.0
                        if abs(c['x0'] - pred) > TOL:
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} {c["x0"]:.2f} vs {pred:.2f}')
                    elif test_id == 'TRAIL_RIGHT_SHIFT' and c['align'] == 'right':
                        if c['x0'] is None or c['ref_x0'] is None or not c['adv']:
                            continue
                        pred = c['ref_x0'] - c['count'] * c['adv']
                        if abs(c['x0'] - pred) > TOL:
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} {c["x0"]:.2f} vs {pred:.2f}')
                    elif test_id == 'TRAIL_TEXT':
                        if c['format'] == 'F8_suffix_empty' and c['value'] == 'neg':
                            want = '1234.50'
                        else:
                            want = {'pos': '1234.50', 'neg': '-1234.50', 'zero': '0.00'}[c['value']]
                        if c['amount'] != want:
                            fails.append(f'{c["sheet"]}:{c["row"]} text {c["amount"]!r}')
                    elif test_id == 'TRAIL_COUNT0':
                        if c['count'] == 0 and (c['x0'] is None or c['ref_x0'] is None
                                                or abs(c['x0'] - c['ref_x0']) > 0.011
                                                or abs(c['x1'] - c['ref_x1']) > 0.011):
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]}')
                    elif test_id == 'TRAIL_INSIDE':
                        if c['x0'] is not None and c['cell_left'] is not None \
                                and c['x0'] < c['cell_left'] + 0.5:
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} left')
                        if c['composite'] is not None and c['cell_right'] is not None \
                                and c['composite'] > c['cell_right'] - 0.5:
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} right')
                    elif test_id == 'TRAIL_RESIDUAL' and c['count'] > 0:
                        if c['composite'] is None or c['cell_right'] is None or not c['adv']:
                            continue
                        resid = c['cell_right'] - c['composite']
                        if not (0.5 <= resid <= c['adv'] + 2.5):
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} resid {resid:.2f}')
                    elif test_id == 'TRAIL_ADV':
                        if c['adv'] is None:
                            continue
                        pd = predicted_adv(face, size, scale, c['fill'])
                        if pd is None or abs(c['adv'] - pd) > 0.15:
                            fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} {c["adv"]:.3f} vs {pd}')
            elif test_id == 'TRAIL_HASH_EQ':
                for c in cc:
                    if c['ref_hash'] != c['hash']:
                        fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} {c["hash"]} vs control {c["ref_hash"]}')
            elif test_id == 'TRAIL_COUNT_LAW':
                for c in cc:
                    if c['align'] != 'left' or c['hash'] or c['ref_hash'] or not c['adv']:
                        continue
                    rsheet = c['sheet'].replace('_left', '_right')
                    lref = n1.get((side, c['sheet'], c['value'], c['width']))
                    rref = n1.get((side, rsheet, c['value'], c['width']))
                    if not lref or not rref or not rref['amount']:
                        continue
                    L = float(lref['amount_x0']) - float(lref['cell_left'])
                    R = float(rref['cell_right']) - float(rref['amount_x1'])
                    W = c['cell_right'] - c['cell_left']
                    pd = predicted_adv(face, size, scale, c['fill'])
                    if pd is None:
                        continue
                    pred = int((W - L - R - (c['x1'] - c['x0'])) // pd)
                    pred = max(0, pred)
                    if abs(pred - c['count']) > 1:
                        fails.append(f'{c["sheet"]}:{c["row"]}:{c["width"]} N={c["count"]} pred={pred}')
            elif test_id == 'TRAIL_MONO':
                groups = collections.defaultdict(list)
                for c in cc:
                    groups[(c['align'], c['format'], c['value'])].append(c)
                for key, rr in groups.items():
                    rr.sort(key=lambda c: float(c['width']))
                    for a, b in zip(rr, rr[1:]):
                        lim = FILL_LIMIT.get(b['fill'], 1)
                        if b['count'] < a['count'] or b['count'] - a['count'] > lim:
                            fails.append(f'{key} {a["width"]}->{b["width"]} {a["count"]}->{b["count"]}')
            elif test_id == 'TRAIL_SIGN':
                for c in cc:
                    if c['format'] == 'F8_suffix_empty' and c['value'] == 'neg':
                        if c['amount'] and c['amount'].startswith('-'):
                            fails.append(f'{c["sheet"]}:{c["row"]} section2 has a minus')
                        continue
                    neg = c['value'] == 'neg'
                    if c['amount'] and neg != c['amount'].startswith('-'):
                        fails.append(f'{c["sheet"]}:{c["row"]} {c["amount"]!r}')
            else:
                fails.append('unknown check ' + test_id)
            res[side] = fails
        out.append('\t'.join([
            workbook, test_id, str(len(by_side['excel'])), str(len(res['excel'])), str(len(res['ours'])),
            accept, '; '.join(res['excel'][:3] or res['ours'][:3])]))
    print('\n'.join(out))
    # per-cell ours-vs-excel agreement
    ex = by_key(rows, 'excel')
    ou = by_key(rows, 'ours')
    diffs = collections.Counter()
    examples = collections.defaultdict(list)
    keyset = set(ex)
    if not ou:
        return
    for k in sorted(keyset):
        e, o = ex[k], ou.get(k)
        if e['format_name'] not in ('S2_space_after', 'D2_dot_after', 'F8_suffix_empty'):
            continue
        if o is None:
            diffs['missing'] += 1
            continue
        if e['hash'] != o['hash']:
            diffs['hash'] += 1
            examples['hash'].append(f'{k} {e["hash"]}->{o["hash"]}')
            continue
        if count(e) != count(o):
            diffs['count'] += 1
            examples['count'].append(f'{k} {count(e)}->{count(o)}')
        if e['amount'] != o['amount']:
            diffs['text'] += 1
            examples['text'].append(f'{k} {e["amount"]!r}->{o["amount"]!r}')
        for fld in ('amount_x0', 'amount_x1'):
            a, b = fnum(e, fld), fnum(o, fld)
            if a is not None and b is not None and abs(a - b) > TOL:
                diffs[fld] += 1
                examples[fld].append(f'{k} {a:.2f}->{b:.2f}')
    if ou:
        print('OURS_VS_EXCEL', ' '.join(f'{k}={v}' for k, v in sorted(diffs.items())) or 'none')
    for k, v in examples.items():
        print(f'  {k}: ' + '; '.join(v[:5]))


def main():
    mode = sys.argv[1]
    args = sys.argv[2:]
    face, size, scale = None, None, None
    if '--face' in args:
        i = args.index('--face')
        face = args[i + 1]
        del args[i:i + 2]
    if '--size' in args:
        i = args.index('--size')
        size = float(args[i + 1])
        del args[i:i + 2]
    if '--scale' in args:
        i = args.index('--scale')
        scale = int(args[i + 1])
        del args[i:i + 2]
    if mode == 'train':
        comp, prefix = args[:2]
        cells = build_cells(load(comp))
        write_tables(cells, prefix)
        summarise([c for c in cells if c['side'] == 'excel'], face or 'Aptos Narrow', size or 14.0, scale or 100)
        print('wrote', prefix + '-cells.tsv', prefix + '-transitions.tsv')
    elif mode == 'check':
        pred, comp, workbook = args[:3]
        check(pred, comp, workbook, face or 'Arial', size or 11.0, scale or 60)
    else:
        raise SystemExit(f'unknown mode {mode!r}')


if __name__ == '__main__':
    main()
