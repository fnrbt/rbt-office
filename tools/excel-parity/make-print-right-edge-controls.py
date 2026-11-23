#!/usr/bin/env python3
"""Writes the literal-vs-generated controls for the printed right edge of right-aligned numbers.

The question is whether the right-edge reserve seen on italic right-aligned cell text depends on the *source* of the
displayed string (a literal text cell vs a number formatted by Excel) or only on the rendered string/font. Each sheet
holds one configuration (face, size, scale, italic, alignment, column width) and adjacent rows alternate a numeric
cell and a literal text cell with exactly the same displayed characters. The value column is filled #FFF2CC so its
printed right edge is the cell's physical right border; the metric is Poppler's word box (`pdftotext -bbox`) to that
edge, never raster ink.

Pairs, one configuration per sheet, upright and italic:
  - `0.000%` numeric overflow -> "#######" vs literal "#######" (Aptos Narrow 14, column width 10);
    plus a General number wide enough to overflow, with a literal seven-hash twin, so number-vs-text is separated
    from percent formatting and from displayed-string equality.
  - `0.00%` numeric -> "-98.41%" etc vs literal same strings.
  - `0.00` ordinary numeric -> "-0.98" etc vs literal same strings.
  - accounting: standard negative "(0.98)" (no trailing `_)` pad) vs literal "(0.98)"; a custom section that adds a
    trailing `_)` pad after the same visible "(0.98)" and its literal twin; and standard positive "0.50" with its
    literal twin (the pad shifts the numeric word, not the literal).
Held out (not used to derive the relations): Georgia 12 `0.00%`, `0.00" kg"` strings, and Aptos Narrow 14 `0.00%`
at 60 % print scale, each upright+italic with literal twins.

    python3 tools/excel-parity/make-print-right-edge-controls.py OUT.xlsx [--manifest OUT.tsv]
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

PCT = '0.00%'
PCT3 = '0.000%'
NUM = '0.00'
KG = '0.00" kg"'
ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
# Same visible negative "(0.98)", but the negative section ends in `_)` (an invisible pad after the `)`).
ACCOUNTING_PAD = '_(* #,##0.00_);_(* \\(#,##0.00\\)_);_(* "-"??_);_(@_)'
HASH_LITERAL = '"#######"'
NUMFMTS = [(200, PCT), (201, NUM), (202, ACCOUNTING), (203, KG), (204, PCT3), (205, ACCOUNTING_PAD), (206, HASH_LITERAL)]
FMT_ID = {PCT: 200, PCT3: 204, NUM: 201, KG: 203, ACCOUNTING: 202, ACCOUNTING_PAD: 205, HASH_LITERAL: 206, None: 0}


def attr(text):
    """Escape an XML attribute value (quotes included)."""
    return escape(text, {'"': '&quot;'})


def txt(text):
    return escape(text)


def num_row(value, text, fmt, pair):
    return {'kind': 'num', 'value': value, 'text': text, 'fmt': fmt, 'pair': pair}


def lit_row(text, fmt, pair):
    return {'kind': 'lit', 'value': None, 'text': text, 'fmt': fmt, 'pair': pair}


PERCENT_ROWS = [
    num_row(-0.9841, '-98.41%', PCT, 'pct_a'),
    lit_row('-98.41%', PCT, 'pct_a'),
    num_row(-1, '-100.00%', PCT, 'pct_b'),
    lit_row('-100.00%', PCT, 'pct_b'),
    num_row(0.5, '50.00%', PCT, 'pct_c'),
    lit_row('50.00%', PCT, 'pct_c'),
    num_row(1234567, '123456700.00%', PCT, 'pct_d'),
    lit_row('123456700.00%', PCT, 'pct_d'),
]
NUMBER_ROWS = [
    num_row(-0.9841, '-0.98', NUM, 'num_a'),
    lit_row('-0.98', NUM, 'num_a'),
    num_row(-1, '-1.00', NUM, 'num_b'),
    lit_row('-1.00', NUM, 'num_b'),
    num_row(0.5, '0.50', NUM, 'num_c'),
    lit_row('0.50', NUM, 'num_c'),
    num_row(1234567, '1234567.00', NUM, 'num_d'),
    lit_row('1234567.00', NUM, 'num_d'),
]
HASH_ROWS = [
    num_row(-0.9841, '#######', PCT3, 'hash_pct_a'),
    lit_row('#######', PCT3, 'hash_pct_a'),
    num_row(1, '#######', HASH_LITERAL, 'hash_fmt_lit'),
    lit_row('#######', HASH_LITERAL, 'hash_fmt_lit'),
    num_row(1234567, '#######', PCT3, 'hash_pct_b'),
    lit_row('#######', PCT3, 'hash_pct_b'),
    # A General number does not overflow to hashes: Excel chooses scientific notation. Kept as a documented
    # non-pair control; the measurement records the real "1.2E+08" (upright) / "1E+08" (italic), not a hash.
    num_row(123456789, '1.2E+08', None, ''),
]
ACCOUNTING_ROWS = [
    # Negative section without a trailing _) pad, literal twin in General (no text-section pads).
    num_row(-0.9841, '(0.98)', ACCOUNTING, 'acc_neg'),
    lit_row('(0.98)', None, 'acc_neg'),
    # Negative section with a trailing _) pad; its literal twin also carries the accounting text section _( @ _).
    num_row(-0.9841, '(0.98)', ACCOUNTING_PAD, 'acc_pad_neg'),
    lit_row('(0.98)', ACCOUNTING_PAD, 'acc_pad_neg'),
    # Positive section has the trailing _) pad in both the number and the text section.
    num_row(0.5, '0.50', ACCOUNTING, 'acc_pos'),
    lit_row('0.50', ACCOUNTING, 'acc_pos'),
    num_row(1234567, '1,234,567.00', ACCOUNTING, ''),
]
KG_ROWS = [
    num_row(-0.9841, '-0.98 kg', KG, 'kg_a'),
    lit_row('-0.98 kg', KG, 'kg_a'),
    num_row(0.5, '0.50 kg', KG, 'kg_b'),
    lit_row('0.50 kg', KG, 'kg_b'),
    num_row(1234567, '1234567.00 kg', KG, 'kg_c'),
    lit_row('1234567.00 kg', KG, 'kg_c'),
]

# name, face, size, italic, width, scale, rows
SHEETS = [
    ('C00_hash_upright', 'Aptos Narrow', 14, False, 10, 100, HASH_ROWS),
    ('C01_hash_italic', 'Aptos Narrow', 14, True, 10, 100, HASH_ROWS),
    ('C02_pct_upright', 'Aptos Narrow', 14, False, 18, 100, PERCENT_ROWS),
    ('C03_pct_italic', 'Aptos Narrow', 14, True, 18, 100, PERCENT_ROWS),
    ('C04_num_upright', 'Aptos Narrow', 14, False, 18, 100, NUMBER_ROWS),
    ('C05_num_italic', 'Aptos Narrow', 14, True, 18, 100, NUMBER_ROWS),
    ('C06_acc_upright', 'Aptos Narrow', 14, False, 18, 100, ACCOUNTING_ROWS),
    ('C07_acc_italic', 'Aptos Narrow', 14, True, 18, 100, ACCOUNTING_ROWS),
    ('C08_georgia12_upright', 'Georgia', 12, False, 18, 100, PERCENT_ROWS),
    ('C09_georgia12_italic', 'Georgia', 12, True, 18, 100, PERCENT_ROWS),
    ('C10_kg_upright', 'Aptos Narrow', 14, False, 24, 100, KG_ROWS),
    ('C11_kg_italic', 'Aptos Narrow', 14, True, 24, 100, KG_ROWS),
    ('C12_scale60_upright', 'Aptos Narrow', 14, False, 18, 60, PERCENT_ROWS),
    ('C13_scale60_italic', 'Aptos Narrow', 14, True, 18, 60, PERCENT_ROWS),
]


def build(out, manifest_path=None):
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for _, face, size, italic, _, _, _ in SHEETS:
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>')
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>')
    borders = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
               '<border><left/><right style="thin"><color rgb="FF808080"/></right><top/><bottom/><diagonal/></border>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>']
    # One xf per (sheet, displayed number format), so literal and numeric twins share the same style object.
    style_of = {}
    for i, (_, _, _, _, _, _, rows) in enumerate(SHEETS):
        for fmt in dict.fromkeys(row['fmt'] for row in rows):
            style_of[(i, fmt)] = len(xfs)
            xfs.append(f'<xf numFmtId="{FMT_ID[fmt]}" fontId="{i + 1}" fillId="2" borderId="1" xfId="0" '
                       f'applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                       f'<alignment horizontal="right"/></xf>')
    numfmts = ''.join(f'<numFmt numFmtId="{i}" formatCode="{attr(code)}"/>' for i, code in NUMFMTS)
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(NUMFMTS)}">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="3">{fills}</fills>'
        f'<borders count="2">{borders}</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    sheet_files = []
    manifest = []
    for i, (name, face, size, italic, width, scale, rows) in enumerate(SHEETS):
        cells = []
        for r, row in enumerate(rows, 1):
            style = style_of[(i, row['fmt'])]
            if row['kind'] == 'num':
                cells.append(f'<c r="B{r}" s="{style}"><v>{row["value"]}</v></c>')
            else:
                cells.append(f'<c r="B{r}" s="{style}" t="inlineStr"><is><t>{txt(row["text"])}</t></is></c>')
            manifest.append((name, i + 1, r, row['kind'], row['text'], row['pair'] or '', face, size,
                             'italic' if italic else 'upright', scale, row['fmt'] or 'General'))
        row_xml = ''.join(f'<row r="{r}">{cell}</row>' for r, cell in enumerate(cells, 1))
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="{width}" customWidth="1"/></cols>'
            f'<sheetData>{row_xml}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            f'<pageSetup paperSize="1" orientation="portrait" scale="{scale}"/></worksheet>')
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
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(sheet_files))
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
    if manifest_path:
        with open(manifest_path, 'w', encoding='utf-8') as f:
            f.write('sheet\tindex\trow\tkind\texpected_text\tpair\tface\tsize\tstyle\tscale\tformat\n')
            for m in manifest:
                f.write('\t'.join(str(v) for v in m) + '\n')


if __name__ == '__main__':
    args = sys.argv[1:]
    out = args[0]
    manifest = None
    if '--manifest' in args:
        manifest = args[args.index('--manifest') + 1]
    build(out, manifest)
