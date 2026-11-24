#!/usr/bin/env python3
"""Check the locked numeric-fill holdout predictions against a measured comparison.

    python3 tools/excel-parity/analyze-print-numfill.py \
        tools/excel-parity/print-numfill-holdout-predictions.tsv \
        COMPARISON.tsv WORKBOOK [BASELINE_COMPARISON.tsv] > RESULTS.tsv

COMPARISON.tsv is measure-print-numfill.py's long table. Every relation is evaluated on Excel first
(the `excel_ok` column: the model must describe Excel itself) and then on ours. BASELINE is optional;
when given, the `OURS_RECORD` rows also carry the before -> after delta.
"""
import collections, sys

ALIGN_SHEET = {'left': 'S_left', 'center': 'S_center', 'right': 'S_right'}


def load(path):
    lines = [l.rstrip('\n') for l in open(path) if l.strip()]
    header = lines[0].split('\t')
    return [dict(zip(header, l.split('\t'))) for l in lines[1:]]


def table(rows):
    t = collections.defaultdict(dict)
    for r in rows:
        t[(r['sheet'], int(r['row']), r['width'])][r['side']] = r
    return t


def fnum(r, key):
    v = r.get(key, '')
    return None if v == '' else float(v)


def main():
    pred_path, comp_path, workbook = sys.argv[1:4]
    base_path = sys.argv[4] if len(sys.argv) > 4 else None
    def norm(rows):
        for r in rows:
            try:
                v = float(r.get('value', 'nan'))
                r['value'] = 'pos' if v > 0 else ('neg' if v < 0 else 'zero')
            except (ValueError, TypeError):
                pass
        return rows
    t = table(norm(load(comp_path)))
    b = table(norm(load(base_path))) if base_path else None

    def get(sheet, row, width, side='excel'):
        return t.get((sheet, int(row), width), {}).get(side)

    def n0(sheet, row, width, align, value=None, side='excel'):
        target = ALIGN_SHEET[align]
        if value is not None:
            for (s, rr, w), sides in t.items():
                x = sides.get(side)
                if (x and s == target and w == width and x['format_name'] == 'N0_plain'
                        and x['value'] == value):
                    return x
        return get(target, row, width, side)

    def n1(sheet, row, width, align, value=None, side='excel'):
        target = ALIGN_SHEET[align]
        if value is not None:
            for (s, rr, w), sides in t.items():
                x = sides.get(side)
                if (x and s == target and w == width and x['format_name'] == 'N1_plain_0'
                        and x['value'] == value):
                    return x
        return None

    checks = [l.rstrip('\n').split('\t') for l in open(pred_path)
              if not l.startswith('#') and l.strip()][1:]
    out = ['workbook\ttest_id\tcells\texcel_failures\tours_failures\texpected\texamples']
    for row in checks:
        test_id, group, check, formats, values, expected, accept = row
        formats = formats.split(',') if formats != '-' else []
        values = values.split(',')
        cells = [r for k, sides in t.items() if 'excel' in sides
                 for r in [sides['excel']] if r['format_name'] in formats]
        if test_id == 'MONO_COUNT':
            cells = []
        exam = ours = 0
        examples = []
        def bad(which, ex):
            nonlocal exam, ours
            if which == 'excel':
                exam += 1
            else:
                ours += 1
            if len(examples) < 3:
                examples.append(ex)

        if test_id in ('ANCHOR_R_S1', 'ANCHOR_R_D1', 'INACTIVE_I1_NEG'):
            for r in cells:
                if r['value'] not in values:
                    continue
                if r['hash'] == 'yes' or int(r['fill_before']) + int(r['fill_after']) == 0:
                    continue
                ref = n1(r['sheet'], r['row'], r['width'], 'right', r['value']) if test_id == 'INACTIVE_I1_NEG' else n0(r['sheet'], r['row'], r['width'], 'right', r['value'])
                if not ref or ref['hash'] == 'yes':
                    continue
                e = fnum(r, 'amount_gap_right'); n = fnum(ref, 'amount_gap_right')
                o = get(r['sheet'], r['row'], r['width'], 'ours')
                if e is None or n is None or abs(e - n) > 0.25:
                    bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} {e} vs {n}")
                if o and o['hash'] == 'no' and fnum(o, 'amount_gap_right') is not None:
                    if abs(fnum(o, 'amount_gap_right') - (n if n is not None else e)) > 0.85:
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']}")
                elif o and o['hash'] == 'yes':
                    bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} hash")

        elif test_id == 'ANCHOR_L_S2':
            for r in cells:
                if r['hash'] == 'yes':
                    continue
                ref = n1(r['sheet'], r['row'], r['width'], 'left', r['value'])
                if not ref or ref['hash'] == 'yes':
                    continue
                e = fnum(r, 'amount_gap_left'); n = fnum(ref, 'amount_gap_left')
                o = get(r['sheet'], r['row'], r['width'], 'ours')
                if e is None or n is None or not (-0.3 <= e - n <= 2.5):
                    bad('excel', f"{r['sheet']}:{r['row']}:{r['width']}")
                if o and o['hash'] == 'yes':
                    bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} hash")
                elif o and e is not None and abs(fnum(o, 'amount_gap_left') - e) > 0.85:
                    bad('ours', f"{r['sheet']}:{r['row']}:{r['width']}")

        elif test_id == 'MONO_COUNT':
            groups = collections.defaultdict(list)
            for k, sides in t.items():
                r = sides.get('excel')
                if r and r['format_name'] in formats:
                    groups[(r['sheet'], r['row'], r['format_name'], r['value'])].append(r)
            for key, rs in groups.items():
                rs.sort(key=lambda r: float(r['width']))
                for side, field in (('excel', 'fill_before'), ('excel_after', 'fill_before')):
                    pass
                for i in range(1, len(rs)):
                    a, bb = int(rs[i - 1]['fill_before']) + int(rs[i - 1]['fill_after']), int(rs[i]['fill_before']) + int(rs[i]['fill_after'])
                    if bb < a or bb - a > 2:
                        bad('excel', f"{key} {rs[i-1]['width']}->{rs[i]['width']} {a}->{bb}")
                for side in ('ours',):
                    arr = [get(rs[i]['sheet'], rs[i]['row'], rs[i]['width'], side) for i in range(len(rs))]
                    for i in range(1, len(arr)):
                        if not arr[i - 1] or not arr[i] or arr[i - 1]['hash'] == 'yes' or arr[i]['hash'] == 'yes':
                            bad(side, f"{key} hash at {rs[i]['width']}")
                            continue
                        a = int(arr[i - 1]['fill_before']) + int(arr[i - 1]['fill_after'])
                        bb = int(arr[i]['fill_before']) + int(arr[i]['fill_after'])
                        if bb < a or bb - a > 2:
                            bad(side, f"{key} {rs[i-1]['width']}->{rs[i]['width']} {a}->{bb}")

        elif test_id == 'COUNT0_FIT':
            for r in cells:
                ref = n0(r['sheet'], r['row'], r['width'], r['align'], r['value'])
                if not ref or ref['hash'] == 'yes':
                    continue
                if r['hash'] == 'yes':
                    bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} hashed")
                o = get(r['sheet'], r['row'], r['width'], 'ours')
                if o and ref['hash'] == 'no' and o['hash'] == 'yes':
                    bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} hash")

        elif test_id == 'ZERO_FALLBACK':
            for r in cells:
                if r['value'] not in values or r['hash'] == 'yes' or int(r['fill_before']) + int(r['fill_after']) != 0:
                    continue
                ref = n0(r['sheet'], r['row'], r['width'], r['align'], r['value'])
                if not ref or ref['hash'] == 'yes':
                    continue
                f = 'amount_gap_right'
                e, n = fnum(r, f), fnum(ref, f)
                if e is None or n is None or abs(e - n) > 0.25:
                    bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} {f} {e} vs {n}")
                o = get(r['sheet'], r['row'], r['width'], 'ours')
                if o and (o['hash'] == 'yes' or fnum(o, f) is None or abs(fnum(o, f) - n) > 0.85):
                    bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} {f}")

        elif test_id == 'INACTIVE_I1_POS':
            for r in cells:
                if r['value'] not in values or r['hash'] == 'yes':
                    continue
                ref = n1(r['sheet'], r['row'], r['width'], r['align'], r['value'])
                if not ref or ref['hash'] == 'yes':
                    continue
                for f in ('amount_gap_left', 'amount_gap_right'):
                    e, n = fnum(r, f), fnum(ref, f)
                    if e is None or n is None or abs(e - n) > 0.10:
                        bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} {f} {e} vs {n}")
                    o = get(r['sheet'], r['row'], r['width'], 'ours')
                    if o and (o['hash'] == 'yes' or fnum(o, f) is None or abs(fnum(o, f) - n) > 0.85):
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} {f}")

        elif test_id == 'LITERAL_Q1':
            for r in cells:
                if r['value'] not in values:
                    continue
                ref = n0(r['sheet'], r['row'], r['width'], r['align'], r['value'])
                if not ref:
                    continue
                for side in ('excel', 'ours'):
                    x = r if side == 'excel' else get(r['sheet'], r['row'], r['width'], 'ours')
                    if not x:
                        continue
                    if x['hash'] == 'no' and x['text'].count('*') != 1:
                        bad(side, f"{r['sheet']}:{r['row']}:{r['width']} stars {x['text']!r}")
                    if side == 'ours' and x['hash'] == 'no' and r['hash'] == 'no':
                        if fnum(x, 'amount_x0') is None or abs(fnum(x, 'amount_x0') - fnum(r, 'amount_x0')) > 0.85:
                            bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} amount")

        elif test_id in ('PREFIX_C1', 'RIGHT_C1', 'ACCT_A1'):
            for r in cells:
                if r['hash'] == 'yes':
                    continue
                align = 'left' if test_id == 'PREFIX_C1' else 'right'
                ref = n0(r['sheet'], r['row'], r['width'], align, r['value'])
                if not ref or ref['hash'] == 'yes':
                    continue
                o = get(r['sheet'], r['row'], r['width'], 'ours')
                if test_id == 'PREFIX_C1':
                    e, n = fnum(r, 'dollar_gap_left'), fnum(ref, 'amount_gap_left')
                    if e is None or n is None or abs(e - n) > 0.75:
                        bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} {e} vs {n}")
                    if o and (o['hash'] == 'yes' or fnum(o, 'dollar_gap_left') is None or abs(fnum(o, 'dollar_gap_left') - e) > 1.0):
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} $")
                elif test_id == 'RIGHT_C1':
                    e, n = fnum(r, 'amount_gap_right'), fnum(ref, 'amount_gap_right')
                    if e is None or n is None or abs(e - n) > 0.50:
                        bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} {e} vs {n}")
                    if o and (o['hash'] == 'yes' or fnum(o, 'amount_gap_right') is None or abs(fnum(o, 'amount_gap_right') - n) > 0.65):
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']}")
                else:
                    e, n = fnum(r, 'amount_gap_right'), fnum(ref, 'amount_gap_right')
                    if e is None or n is None or not (1.5 <= e - n <= 7.0):
                        bad('excel', f"{r['sheet']}:{r['row']}:{r['width']} {e} vs {n}")
                    if o and (o['hash'] == 'yes' or fnum(o, 'amount_gap_right') is None or abs(fnum(o, 'amount_gap_right') - e) > 0.65):
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']}")

        elif test_id == 'NO_OVERFLOW':
            for r in cells:
                if r['hash'] == 'yes':
                    continue
                left, right = fnum(r, 'cell_left'), fnum(r, 'cell_right')
                for side in ('excel', 'ours'):
                    x = r if side == 'excel' else get(r['sheet'], r['row'], r['width'], 'ours')
                    if not x or x['hash'] == 'yes':
                        continue
                    x0, x1 = fnum(x, 'amount_x0'), fnum(x, 'amount_x1')
                    if x1 is None or x0 is None:
                        continue
                    if r['format_name'] == 'S2_space_after':
                        ok = x0 >= left + 0.5
                    else:
                        ok = x1 <= right - 0.5
                    if not ok:
                        bad(side, f"{r['sheet']}:{r['row']}:{r['width']} {x0}/{x1}")

        elif test_id == 'TEXT':
            for r in cells:
                if r['hash'] == 'yes':
                    continue
                want = '0.00' if r['value'] == 'zero' else '1234.50'
                for side in ('excel', 'ours'):
                    x = r if side == 'excel' else get(r['sheet'], r['row'], r['width'], 'ours')
                    if not x or x['hash'] == 'yes':
                        if side == 'ours':
                            bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} hash")
                        continue
                    norm_amount = x['amount'].replace(',', '')
                    norm_text = x['text'].replace(',', '')
                    zero_dash = r['value'] == 'zero' and '"-"??' in r['format']
                    if zero_dash:
                        if '-' not in x['text']:
                            bad(side, f"{r['sheet']}:{r['row']}:{r['width']} no zero dash")
                    elif want not in norm_amount and want not in norm_text:
                        bad(side, f"{r['sheet']}:{r['row']}:{r['width']} {x['amount']!r}")
                    if r['value'] == 'neg' and '-' not in x['text'] and '(' not in x['text']:
                        bad(side, f"{r['sheet']}:{r['row']}:{r['width']} no sign")
                    if '$' in r['format'] and x['dollar'] != '$':
                        bad(side, f"{r['sheet']}:{r['row']}:{r['width']} no $")

        elif test_id in ('OURS_SUPPORTED', 'OURS_RECORD'):
            for r in cells:
                if r['value'] not in ('pos', 'neg', 'zero'):
                    continue
                o = get(r['sheet'], r['row'], r['width'], 'ours')
                if not o:
                    if test_id == 'OURS_SUPPORTED':
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} missing")
                    continue
                detail = []
                if o['hash'] != r['hash']:
                    detail.append(f"hash {r['hash']}->{o['hash']}")
                for f in ('amount_gap_left', 'amount_gap_right', 'dollar_gap_left'):
                    a, bb = fnum(r, f), fnum(o, f)
                    if a is not None and bb is not None and abs(a - bb) > 0.85:
                        detail.append(f"{f} {a:.3f}->{bb:.3f}")
                ca = int(r['fill_before']) + int(r['fill_after'])
                cb = int(o['fill_before']) + int(o['fill_after'])
                if abs(ca - cb) > 1:
                    detail.append(f"count {ca}->{cb}")
                if detail:
                    if test_id == 'OURS_SUPPORTED':
                        bad('ours', f"{r['sheet']}:{r['row']}:{r['width']} " + '; '.join(detail))
                    else:
                        examples.append(f"{r['sheet']}:{r['row']}:{r['width']} " + '; '.join(detail))
        out.append('\t'.join([workbook, test_id, str(len(cells)), str(exam), str(ours), expected,
                              ' || '.join(examples)]))
    print('\n'.join(out))


if __name__ == '__main__':
    main()
