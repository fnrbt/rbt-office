#!/usr/bin/env python3
"""Scores the hint-discriminating challenge (baseline vs FreeType-v35 hinted) on Excel's PDF.

Both locks were written before the export; this script only joins them to
measure-print-advance-oracle.py's results.  It reports full/hash decisions and,
on hashed cells, the printed `#` count against each lock's `n` prediction.

    python3 tools/excel-parity/compare-print-hint-challenge.py \
        ENGINE_LOCK.tsv OUR_LOCK.tsv RESULTS.tsv OUT.tsv
"""
import collections
import csv
import sys


def load(path):
    with open(path, encoding='utf-8') as f:
        return list(csv.DictReader((l for l in f if not l.startswith('#')), delimiter='\t'))


def main():
    if len(sys.argv) != 5:
        raise SystemExit(__doc__)
    engine_path, our_path, results_path, out_path = sys.argv[1:]
    E = {(d['sheet'], d['column'], d['row']): d for d in load(engine_path)}
    O = {(d['sheet'], d['column'], d['row']): d for d in load(our_path)}
    R = {(d['sheet'], d['column'], d['row']): d for d in load(results_path)}
    # lock agreement check
    mism = [(k, O[k]['baseline_full'], E[k]['base_full'], O[k]['hinted_full'], E[k]['cand_full'])
            for k in E if k in O and (O[k]['baseline_full'] != E[k]['base_full']
                                      or O[k]['hinted_full'] != E[k]['cand_full']
                                      or abs(float(O[k]['baseline_text_pt']) - float(E[k]['base_run_pt'])) > 1e-4
                                      or abs(float(O[k]['hinted_text_pt']) - float(E[k]['cand_run_pt'])) > 1e-4)]
    out = ['section\tsheet\ttext\tcells\texcel_full\tbase_full\tcand_full\t'
           'base_wrong\tcand_wrong\thash_cells\tbase_count_wrong\tcand_count_wrong\t'
           'base_misses\tcand_misses\tcand_count_misses']
    clusters = collections.defaultdict(lambda: dict(cells=0, excel=0, bf=0, cf=0, bw=0, cw=0,
                                                    hashes=0, bn=0, cn=0, bm=[], cm=[], cmm=[]))
    wide = collections.defaultdict(lambda: [0, 0.0, 0.0, 0.0])
    for k, e in E.items():
        if k[2] != '1' or k not in R or k not in O:
            continue
        r = R[k]
        excel_full = int(r['nHashes']) == 0
        bf = e['base_full'] == 'yes'
        cf = e['cand_full'] == 'yes'
        if e['kind'] == 'wide':
            if r['nHashes'] == '0':
                w = wide[(e['sheet'], e['text'])]
                w[0] += 1
                w[1] += float(r['pdf_adv_pt'])
                w[2] += float(e['base_run_pt'])
                w[3] += float(e['cand_run_pt'])
            continue
        c = clusters[(e['sheet'], e['text'])]
        c['cells'] += 1
        c['excel'] += excel_full
        c['bf'] += bf
        c['cf'] += cf
        if bf != excel_full:
            c['bw'] += 1
            c['bm'].append(r['column'])
        if cf != excel_full:
            c['cw'] += 1
            c['cm'].append(r['column'])
        if r['nHashes'] != '0':
            c['hashes'] += 1
            if int(e['base_n']) != int(r['nHashes']):
                c['bn'] += 1
            if int(e['cand_n']) != int(r['nHashes']):
                c['cn'] += 1
                c['cmm'].append(f"{r['column']}:{e['cand_n']}->{r['nHashes']}")
    te = tc = tbn = tcn = 0
    for (sheet, text), c in sorted(clusters.items()):
        out.append(f"cutoff\t{sheet}\t{text}\t{c['cells']}\t{c['excel']}\t{c['bf']}\t{c['cf']}\t"
                   f"{c['bw']}\t{c['cw']}\t{c['hashes']}\t{c['bn']}\t{c['cn']}\t"
                   f"{','.join(c['bm'])}\t{','.join(c['cm'])}\t{','.join(c['cmm'])}")
        te += c['bw']; tc += c['cw']; tbn += c['bn']; tcn += c['cn']
    out += ['', 'section\tsheet\ttext\tcells\tmeasured\tbaseline_adv\thinted_adv\t'
            'baseline-minus-measured\thinted-minus-measured']
    for (sheet, text), w in sorted(wide.items()):
        n = w[0] or 1
        out.append(f'advance\t{sheet}\t{text}\t{w[0]}\t{w[1] / n:.5f}\t{w[2] / n:.5f}\t'
                   f'{w[3] / n:.5f}\t{(w[2] - w[1]) / n:+.5f}\t{(w[3] - w[1]) / n:+.5f}')
    out += ['', 'lock_disagreements\t', str(len(mism)),
            'totals\tcells\tbase_wrong\tcand_wrong\tbase_count_wrong\tcand_count_wrong',
            f'dense\t{sum(c["cells"] for c in clusters.values())}\t{te}\t{tc}\t{tbn}\t{tcn}']
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print('\n'.join(out[-6:]))


if __name__ == '__main__':
    main()
