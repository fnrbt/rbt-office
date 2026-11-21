#!/usr/bin/env python3
"""Writes synthetic workbooks (not saved by Excel) for the sizing and drawing rules the grid follows to draw a sheet
as Excel 2019 does (found on real-world real-estate models, each checked in Excel 2019; tools/xlsx-browser check
`sizing-excel-parity`, tests/office-formats/xlsx-sizing.fsx):

- synthetic-font-widths.xlsx: the Normal font is Arial MT 12 (as Excel substitutes it: Arial), maximum digit
  width 9 px, so a column of width w is trunc(((256 w + trunc(128 / 9)) / 256) x 9) px wide:
  - "Widths": A..H of widths 1, 2.09765625, 8.8984375, 10, 20, 38.796875, 0.5, 2 (9, 19, 80, 90, 180, 349, 4,
    18 px); I.. no width: the default, 8 characters, 8 x 9 + 7 padding up to a multiple of 8 px: 80 px. Rows
    without a height: 15 pt (Arial 12's own), whatever `defaultRowHeight` says without `customHeight`.
  - "DefaultCol": defaultColWidth="20": the columns of no width are 180 px.
  - "BaseCol": baseColWidth="12" and no defaultColWidth: 12 x 9 + 7 up to a multiple of 8: 120 px.
  - "RowDefault": defaultRowHeight="30" without customHeight: rows 15 pt all the same (Excel's own default).
  - "RowCustom": defaultRowHeight="30" customHeight="1": rows 30 pt.
- synthetic-normal-font-first.xlsx / synthetic-normal-font-style.xlsx: which font is the Normal one - fonts[0]
  Arial 12 with the Normal style's xf on Calibri 11, and the other way round.
- synthetic-row-fit.xlsx (Calibri 11): "Rows": which rows Excel fits to their cells when it opens a file - rows
  of no `ht` (larger fonts, wrapped text), a pixel more for each of `thickTop` / `thickBot`; "RowStyle": a row's style (`<row s customFormat="1">`) to the last
  column; "ColStyle", "ColStyleAll", "ColStyleSome": a column style's font in the rows' height (all columns of
  Arial 10: rows of 12.75 pt); "Widths": Calibri 11's widths (maximum digit width 7, truncated: 9.140625 is 64 px).
- synthetic-theme-shapes.xlsx: shapes of theme fills (Office 2007 theme): `fillRef idx` 0..3 (none, the theme's
  solid, its two gradients, phClr replaced by the reference's colour) and `gradFill` of `schemeClr` with
  lumMod / lumOff.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-sizing-fixtures.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUTDIR = os.path.join(REPO, 'tests', 'test-documents-excel')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'


def styles(fonts, style_font=0, xfs=None, fills=''):
    """A style sheet: `fonts` the <font> bodies, the Normal style's xf on font `style_font`, `xfs` the cellXfs
    (numFmtId fontId fillId attributes and an optional alignment), `fills` extra solid fills (from fillId 2)."""
    xfs = xfs or ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    fill_list = ('<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
                 + ''.join(f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill>'
                           for rgb in fills))
    return DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{2 + len(fills)}">{fill_list}</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        f'<cellStyleXfs count="1"><xf numFmtId="0" fontId="{style_font}" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/></styleSheet>')


def inline(ref, style, text):
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def sheet(rows, fmt='<sheetFormatPr defaultRowHeight="15"/>', cols='', extra='', selected=False):
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
        f'<sheetViews><sheetView {"tabSelected=\"1\" " if selected else ""}workbookViewId="0"/></sheetViews>'
        + fmt + (f'<cols>{cols}</cols>' if cols else '') + '<sheetData>' + ''.join(rows) + '</sheetData>' + extra
        + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        + ('<drawing r:id="rId1"/>' if 'DRAWING' in fmt else '') + '</worksheet>').replace('<!--DRAWING-->', '')


def col(first, last, width=None, style=None):
    return (f'<col min="{first}" max="{last}"' + (f' width="{width}" customWidth="1"' if width is not None else '')
            + (f' style="{style}"' if style is not None else '') + '/>')


THEME = DECL + (
    f'<a:theme xmlns:a="{A}" name="Office Theme"><a:themeElements>'
    '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
    '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
    '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
    '<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
    '<a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
    '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>'
    '<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/>'
    '<a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>'
    '</a:fmtScheme></a:themeElements></a:theme>')


def package(path, sheets, styles_xml, drawings=None, theme=False):
    """`sheets`: (name, xml); `drawings`: {sheet index: drawing xml}."""
    drawings = drawings or {}
    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>' for i in range(len(sheets)))
    over += ''.join(f'<Override PartName="/xl/drawings/drawing{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                    for i in drawings)
    if theme:
        over += '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>' + over + '</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                                   + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(sheets))
                                   + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL}/styles" Target="styles.xml"/>'
            + (f'<Relationship Id="rId{len(sheets) + 2}" Type="{REL}/theme" Target="theme/theme1.xml"/>' if theme else '')
            + '</Relationships>'),
        'xl/styles.xml': styles_xml,
    }
    if theme:
        files['xl/theme/theme1.xml'] = THEME
    for i, (_, xml) in enumerate(sheets):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    for i, xml in drawings.items():
        files[f'xl/drawings/drawing{i + 1}.xml'] = xml
        files[f'xl/worksheets/_rels/sheet{i + 1}.xml.rels'] = DECL + (
            f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing{i + 1}.xml"/></Relationships>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(path)


ARIAL_MT_12 = '<sz val="12"/><name val="Arial MT"/>'
ARIAL_12 = '<sz val="12"/><name val="Arial"/><family val="2"/>'
CALIBRI_11 = '<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>'
CALIBRI_20 = '<sz val="20"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>'
CALIBRI_8 = '<sz val="8"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>'
LONG = 'Releasing cost per unit, including the broker fee and the marketing'


def width_rows():
    return [f'<row r="{r}">' + inline(f'A{r}', 0, f'row {r}') + '</row>' for r in (1, 2, 3)]


WIDTH_COLS = ''.join(col(i + 1, i + 1, w) for i, w in enumerate([1, 2.09765625, 8.8984375, 10, 20, 38.796875, 0.5, 2]))

package(os.path.join(OUTDIR, 'synthetic-font-widths.xlsx'), [
    ('Widths', sheet(width_rows(), '<sheetFormatPr defaultRowHeight="15.7"/>', WIDTH_COLS, selected=True)),
    ('DefaultCol', sheet(width_rows(), '<sheetFormatPr defaultColWidth="20" defaultRowHeight="15.7"/>')),
    ('BaseCol', sheet(width_rows(), '<sheetFormatPr baseColWidth="12" defaultRowHeight="15.7"/>')),
    ('RowDefault', sheet(width_rows(), '<sheetFormatPr defaultRowHeight="30"/>')),
    ('RowCustom', sheet(width_rows(), '<sheetFormatPr defaultRowHeight="30" customHeight="1"/>')),
], styles([ARIAL_MT_12]))

package(os.path.join(OUTDIR, 'synthetic-normal-font-first.xlsx'), [
    ('Widths', sheet(width_rows(), '<sheetFormatPr defaultRowHeight="15"/>', WIDTH_COLS, selected=True)),
], styles([ARIAL_12, CALIBRI_11], style_font=1))

package(os.path.join(OUTDIR, 'synthetic-normal-font-style.xlsx'), [
    ('Widths', sheet(width_rows(), '<sheetFormatPr defaultRowHeight="15"/>', WIDTH_COLS, selected=True)),
], styles([CALIBRI_11, ARIAL_12], style_font=1, xfs=['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>']))

# Calibri 11. Styles: 0 plain, 1 Calibri 20, 2 wrapped, 3 Calibri 8, 4 yellow, 5 grey, 6 wrapped top-aligned.
ROW_XFS = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90"/></xf>',
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
]
FIT_ROWS = [
    '<row r="1">' + inline('A1', 1, 'Big') + '</row>',
    '<row r="2">' + inline('A2', 0, 'wrap') + inline('B2', 2, LONG) + '</row>',
    '<row r="3" ht="15">' + inline('A3', 0, 'ht 15') + inline('B3', 2, LONG) + '</row>',
    '<row r="4" ht="15" customHeight="1">' + inline('A4', 0, 'custom') + inline('B4', 2, LONG) + '</row>',
    '<row r="5" ht="30">' + inline('A5', 0, 'ht 30') + '</row>',
    '<row r="6">' + inline('A6', 2, LONG) + '<c r="B6" s="2"/></row>',
    '<row r="7"><c r="A7" s="1"/></row>',
    '<row r="8"><c r="A8" s="1"><v>123</v></c></row>',
    '<row r="9">' + inline('A9', 0, 'short') + inline('B9', 2, 'ok') + '</row>',
    '<row r="10">' + inline('A10', 0, 'line one&#10;line two') + '</row>',
    '<row r="11" ht="12">' + inline('A11', 1, 'ht 12') + '</row>',
    '<row r="12">' + inline('A12', 3, 'small') + '</row>',
    '<row r="13">' + inline('A13', 0, 'lines') + inline('B13', 6, 'one&#10;two&#10;three') + '</row>',
    '<row r="14" spans="1:2" x14ac:dyDescent="0.25" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
    + inline('A14', 1, 'Big') + '</row>',
    '<row r="15" customHeight="1">' + inline('A15', 1, 'Big') + '</row>',
    '<row r="16">' + inline('A16', 2, LONG) + '<c r="B16" s="2"/></row>',
    '<row r="17" s="1" customFormat="1"/>',
    '<row r="18" s="1" customFormat="1">' + inline('A18', 0, 'row of 20 pt') + '</row>',
    '<row r="19">' + inline('A19', 0, 'col style') + '</row>',
    '<row r="20" hidden="1">' + inline('A20', 1, 'Big') + '</row>',
    '<row r="21">' + inline('A21', 0, 'hidden col') + inline('E21', 1, 'Big') + '</row>',
    '<row r="22">' + inline('A22', 7, 'Turned text') + '</row>',
    '<row r="23">' + inline('A23', 0, 'arial') + inline('B23', 8, LONG) + '</row>',
    '<row r="24">' + inline('A24', 9, 'Arial 12') + '</row>',
    '<row r="25">' + inline('A25', 10, 'Arial 20') + '</row>',
    '<row r="26">' + inline('A26', 11, 'Times 14') + '</row>',
    '<row r="27">' + inline('A27', 12, 'Big bold') + '</row>',
    '<row r="28" thickBot="1">' + inline('A28', 0, 'thickBot') + '</row>',
    '<row r="29" thickTop="1" thickBot="1">' + inline('A29', 1, 'Big, both') + '</row>',
]
STYLE_ROWS = [
    '<row r="3" spans="1:3" s="4" customFormat="1">' + inline('A3', 4, 'row style') + '<c r="B3" s="4"/><c r="C3" s="4"/></row>',
    '<row r="5" s="4" customFormat="1"/>',
    '<row r="7" s="4">' + inline('A7', 4, 'no customFormat') + '</row>',
    '<row r="9" s="4" customFormat="1">' + inline('B9', 0, 'plain cell') + '</row>',
]
package(os.path.join(OUTDIR, 'synthetic-row-fit.xlsx'), [
    ('Rows', sheet(FIT_ROWS, '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12) + col(5, 5, 0),
                   extra='<mergeCells count="1"><mergeCell ref="A16:B16"/></mergeCells>', selected=True)),
    ('S1', sheet(FIT_ROWS[:2], '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12))),
    ('S2', sheet(FIT_ROWS[:2], '<sheetFormatPr defaultRowHeight="15" customHeight="1"/>', col(2, 2, 12))),
    ('S3', sheet(FIT_ROWS[:3], '<sheetFormatPr defaultRowHeight="15" customHeight="1"/>', col(2, 2, 12))),
    ('S4', sheet(FIT_ROWS[:2] + [FIT_ROWS[3]], '<sheetFormatPr defaultRowHeight="15" customHeight="1"/>', col(2, 2, 12))),
    ('S5', sheet(FIT_ROWS[:2] + [FIT_ROWS[3]], '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12))),
    ('S6', sheet(FIT_ROWS[:2] + ['<row r="4" ht="30">' + inline('A4', 0, 'ht 30') + '</row>'], '<sheetFormatPr defaultRowHeight="15" customHeight="1"/>', col(2, 2, 12))),
    ('S7', sheet(FIT_ROWS[:2], '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>', col(2, 2, 12))),
    ('S8', sheet(FIT_ROWS[:2], '<sheetFormatPr defaultRowHeight="20"/>', col(2, 2, 12))),
    ('ColStyleAll', sheet(['<row r="1">' + inline('A1', 0, 'Calibri') + '</row>', '<row r="2">' + inline('A2', 13, 'Arial') + '</row>',
                           '<row r="4"/>'], '<sheetFormatPr defaultRowHeight="15"/>', col(1, 16384, 9.140625, 13))),
    ('ColStyleSome', sheet(['<row r="2">' + inline('A2', 13, 'Arial') + '</row>'], '<sheetFormatPr defaultRowHeight="15"/>', col(1, 3, 9.140625, 13))),
    ('ColStyle', sheet(FIT_ROWS[:2] + ['<row r="3">' + inline('A3', 0, 'plain') + '</row>'], '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12) + col(4, 4, 9.140625, 1))),
    ('RowStyle', sheet(STYLE_ROWS, '<sheetFormatPr defaultRowHeight="15"/>', col(8, 9, 9.140625, style=5))),
    ('Widths', sheet(width_rows(), '<sheetFormatPr defaultRowHeight="15"/>',
                     ''.join(col(i + 1, i + 1, w) for i, w in enumerate([9.140625, 0.5, 2.09765625, 13, 1, 8.8984375])))),
], styles([CALIBRI_11, CALIBRI_20, CALIBRI_8, '<sz val="10"/><name val="Arial"/><family val="2"/>', ARIAL_12,
            '<sz val="20"/><name val="Arial"/><family val="2"/>', '<sz val="14"/><name val="Times New Roman"/><family val="1"/>',
            '<b/><sz val="20"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>'], xfs=ROW_XFS, fills=['FFFFFF00', 'FFD9D9D9']))


def anchor(row, body, col0=1, col1=4):
    return (f'<xdr:twoCellAnchor><xdr:from><xdr:col>{col0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{col1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row + 2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            + body + '<xdr:clientData/></xdr:twoCellAnchor>')


def shape(id_, prst, fill, fill_ref, line_ref='<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>', text=''):
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id_ + 1}" name="Shape {id_}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="381000"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill}</xdr:spPr>'
            f'<xdr:style>{line_ref}{fill_ref}<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
            '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1100"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr><a:t>{text or id_}</a:t></a:r></a:p></xdr:txBody></xdr:sp>')


def ref(idx, color, mods=''):
    return f'<a:fillRef idx="{idx}"><a:schemeClr val="{color}">{mods}</a:schemeClr></a:fillRef>'


LUM_GRAD = ('<a:gradFill flip="none" rotWithShape="1"><a:gsLst>'
            '<a:gs pos="0"><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:gs>'
            '<a:gs pos="100000"><a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:gs>'
            '</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>')
TINT_GRAD = ('<a:gradFill rotWithShape="1"><a:gsLst>'
             '<a:gs pos="0"><a:schemeClr val="accent2"><a:tint val="66000"/><a:satMod val="160000"/></a:schemeClr></a:gs>'
             '<a:gs pos="50000"><a:schemeClr val="accent2"><a:tint val="44500"/><a:satMod val="160000"/></a:schemeClr></a:gs>'
             '<a:gs pos="100000"><a:schemeClr val="accent2"><a:tint val="23500"/><a:satMod val="160000"/></a:schemeClr></a:gs>'
             '</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>')
SHAPES = [
    shape(1, 'rect', '', ref(1, 'accent1')),
    shape(2, 'rect', '', ref(2, 'dk1'), '<a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef>'),
    shape(3, 'rect', '', ref(2, 'accent1')),
    shape(4, 'rect', '', ref(3, 'accent1')),
    shape(5, 'rect', '', ref(0, 'accent1'), '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'),
    shape(6, 'rect', LUM_GRAD, ref(1, 'accent1')),
    shape(7, 'rect', TINT_GRAD, ref(1, 'accent2')),
    shape(8, 'roundRect', '', ref(2, 'dk1'), '<a:lnRef idx="1"><a:schemeClr val="dk1"/></a:lnRef>', 'Button'),
    shape(9, 'rect', '', ref(1001, 'accent1')),
    shape(10, 'rect', '', ref(1002, 'accent3')),
]
DRAWING = DECL + f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}">' + ''.join(anchor(1 + 3 * i, s) for i, s in enumerate(SHAPES)) + '</xdr:wsDr>'
package(os.path.join(OUTDIR, 'synthetic-theme-shapes.xlsx'), [
    ('Shapes', sheet([f'<row r="{2 + 3 * i}">' + inline(f'A{2 + 3 * i}', 0, str(i + 1)) + '</row>' for i in range(len(SHAPES))],
                     '<sheetFormatPr defaultRowHeight="15"/><!--DRAWING-->', selected=True)),
], styles([CALIBRI_11]), drawings={0: DRAWING}, theme=True)
