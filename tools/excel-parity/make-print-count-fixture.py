#!/usr/bin/env python3
"""Builds the printed hash-count transition oracle workbooks (training / holdout).

Every test column has its own solid fill, so `measure-print-count-transitions.py`
reads the column's physical width from the PDF fill rectangle and the printed
run from Poppler's word box. Row 1 is a numeric value too wide for the column
(hashes) or, at the widest diagnostic, the full formatted value. Row 2 is a
non-overflow control (`-98.41`, format `0.00`).

Widths are deliberately dense (1/16 character, i.e. about 0.35 pt at 100 %)
across several adjacent count transitions, so a transition boundary is located
to better than half a grid step rather than at the coarse 1/4-character sweep
of make-print-hash-geometry.py.

The training split covers upright/italic, unformatted (`0`) and accounting
pads, Aptos Narrow 14 / Calibri 11 / Arial 12, and 100 % / 75 % print scale.
The holdout split is an unseen face/size/scale and the percentage and grouped
decimal formats; its predictions are locked before this workbook is exported.

    python3 tools/excel-parity/make-print-count-fixture.py OUT.xlsx META.tsv train|holdout
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

# format code -> custom numFmtId (1 is the built-in `0`).
FMT_ID = {
    '0': 1,
    '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)': 164,
    '"$"#,##0_)': 165,
    '0.00%': 166,
    '#,##0.00': 167,
}
FMT_ORDER = list(FMT_ID)


def fill_color(i):
    """Identical palette to measure-print-hash-geometry.py, so the same SVG reader maps a fill to its column."""
    return (0xE0 + (i % 16), 0x20 + ((i * 37) % 0xC0), 0x30 + ((i * 71) % 0xA0))


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def widths(start, stop, step=0.0625):
    n = int(round((stop - start) / step)) + 1
    return [round(start + i * step, 4) for i in range(n)]


# name, face, size, italic, scale, format, value, widths, split
CASES = [
    # ---- training: multiple adjacent transitions per case ----
    ('T_aptos14_up_0', 'Aptos Narrow', 14, False, 100, '0', 1123456789012,
     widths(8.0, 12.125), 'train'),
    ('T_aptos14_it_0', 'Aptos Narrow', 14, True, 100, '0', 1123456789012,
     widths(8.5, 12.625), 'train'),
    ('T_calibri11_up_0', 'Calibri', 11, False, 100, '0', 1123456789012,
     widths(8.5, 12.625), 'train'),
    ('T_calibri11_it_0', 'Calibri', 11, True, 100, '0', 1123456789012,
     widths(8.75, 13.125), 'train'),
    ('T_arial12_up_0', 'Arial', 12, False, 100, '0', 1123456789012,
     widths(10.0, 15.375), 'train'),
    ('T_aptos14_up_acct', 'Aptos Narrow', 14, False, 100,
     '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)', 1234567,
     widths(8.0, 12.0), 'train'),
    ('T_calibri11_up_acct', 'Calibri', 11, False, 100,
     '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)', 1234567,
     widths(8.0, 11.25), 'train'),
    ('T_aptos14_up_0_75', 'Aptos Narrow', 14, False, 75, '0', 1123456789012,
     widths(8.25, 12.375), 'train'),
    ('T_aptos14_it_0_75', 'Aptos Narrow', 14, True, 75, '0', 1123456789012,
     widths(8.75, 12.875), 'train'),
    # ---- training: dense windows at the transition from hashes to the full number ----
    ('T_aptos14_up_full', 'Aptos Narrow', 14, False, 100, '0', 1123456789012,
     widths(16.5, 17.75), 'train'),
    ('T_aptos14_it_full', 'Aptos Narrow', 14, True, 100, '0', 1123456789012,
     widths(17.0, 18.25), 'train'),
    ('T_calibri11_up_full', 'Calibri', 11, False, 100, '0', 1123456789012,
     widths(12.75, 13.75), 'train'),
    ('T_calibri11_it_full', 'Calibri', 11, True, 100, '0', 1123456789012,
     widths(13.25, 14.25), 'train'),
    ('T_arial12_up_full', 'Arial', 12, False, 100, '0', 1123456789012,
     widths(15.75, 16.75), 'train'),
    ('T_aptos14_up_acct_full', 'Aptos Narrow', 14, False, 100,
     '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)', 1234567,
     widths(12.0, 15.0), 'train'),
    ('T_aptos14_up_0_75_full', 'Aptos Narrow', 14, False, 75, '0', 1123456789012,
     widths(16.5, 18.5), 'train'),
    # ---- holdout: unseen face/size/scale and different formats ----
    ('H_verdana13_up_pct', 'Verdana', 13, False, 100, '0.00%', 12345.6789,
     widths(8.0, 15.0) + widths(16.0, 21.0), 'holdout'),
    ('H_georgia15_it_grp', 'Georgia', 15, True, 90, '#,##0.00', 1234567,
     widths(8.0, 15.0) + widths(16.0, 21.0), 'holdout'),
]


def attr(text):
    return escape(text, {'"': '&quot;'})


def build(out, meta_path, split):
    cases = [c for c in CASES if c[8] == split]
    if not cases:
        raise SystemExit(f'no cases for split {split}')

    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for _, face, size, italic, *_ in cases:
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>')
    max_cols = max(len(c[7]) for c in cases)
    fills = ['<fill><patternFill patternType="none"/></fill>',
             '<fill><patternFill patternType="gray125"/></fill>']
    for i in range(max_cols):
        r, g, b = fill_color(i)
        fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{r:02X}{g:02X}{b:02X}"/>'
                     f'<bgColor indexed="64"/></patternFill></fill>')

    used_fmts = sorted({FMT_ID[c[5]] for c in cases if FMT_ID[c[5]] != 1})
    numfmts = ''.join(f'<numFmt numFmtId="{fid}" formatCode="{attr(code)}"/>'
                      for fid in used_fmts for code in FMT_ORDER if FMT_ID[code] == fid)

    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
           '<alignment horizontal="left"/></xf>']
    style_of = {}
    for ci, (_, _, _, _, _, fmt, _, ws, _) in enumerate(cases):
        font_id = ci + 1
        for col in range(len(ws)):
            for row, fid in ((1, FMT_ID[fmt]), (2, 2)):
                style_of[(ci, col, row)] = len(xfs)
                xfs.append(f'<xf numFmtId="{fid}" fontId="{font_id}" fillId="{2 + col}" borderId="0" xfId="0" '
                           f'applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                           f'<alignment horizontal="right"/></xf>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(used_fmts)}">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(fills)}">' + ''.join(fills) + '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    meta = ['sheet\tcolumn\twidth_chars\talign\tface\tsize_pt\titalic\tscale\tkind\trow1\trow2\tformat\tvalue']
    sheet_files = []
    for ci, (name, face, size, italic, scale, fmt, value, ws, _) in enumerate(cases):
        cols = ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>'
                       for j, w in enumerate(ws))
        cells = [f'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>{escape(name)}</t></is></c>']
        for j in range(len(ws)):
            ref = f'{col_letter(2 + j)}1'
            cells.append(f'<c r="{ref}" s="{style_of[(ci, j, 1)]}"><v>{value}</v></c>')
        cells.append('</row><row r="2">')
        for j in range(len(ws)):
            ref = f'{col_letter(2 + j)}2'
            cells.append(f'<c r="{ref}" s="{style_of[(ci, j, 2)]}"><v>-98.41</v></c>')
            meta.append('\t'.join([name, str(2 + j), f'{ws[j]:g}', 'right', face, str(size),
                                   '1' if italic else '0', str(scale), 'overflow',
                                   str(value), '-98.41', fmt, str(value)]))
        cells.append('</row>')
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if ci == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols><col min="1" max="1" width="22" customWidth="1"/>{cols}</cols>'
            f'<sheetData>{"".join(cells)}</sheetData>'
            '<pageMargins left="0.4" right="0.4" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            f'<pageSetup paperSize="1" orientation="landscape" scale="{scale}"/></worksheet>')
        sheet_files.append((name, xml))

    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                      for i, (name, _) in enumerate(sheet_files))
            + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                      for i in range(len(sheet_files)))
            + f'<Relationship Id="rId{len(sheet_files) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheet_files):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    with open(meta_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(meta) + '\n')
    print(f'wrote {out} ({len(sheet_files)} sheets, {sum(len(c[7]) for c in cases)} test columns) and {meta_path}')


if __name__ == '__main__':
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    build(sys.argv[1], sys.argv[2], sys.argv[3])
