#!/usr/bin/env python3
"""Writes synthetic workbooks (not saved by Excel) of charts as the real-estate models have them, for the rules of
how Excel 2019 draws a chart (each checked in Excel 2019 on the VM: `chart-probe.vbs` reads what Excel made of every
chart - its axes' scales, its plot area, its labels - and `shoot-sheets.ps1` its picture; tools/xlsx-browser check
`charts-excel-parity`, tests/formula/stage52-charts.fsx "Excel 2019:").

- synthetic-chart-parity.xlsx, sheet "Charts" (Calibri 11), one chart a few rows apart, each named as below:
  - "ColScale": columns of 3, 7, 5, 6 at 480 x 288 px: the automatic scale (0 .. 8 by 1).
  - "ColShort": the same 480 x 110 px: fewer intervals on a short axis (0 .. 8 by 2?).
  - "ColClose": 50, 60, 55, 52: values close together far from zero (44 .. 62 by 2).
  - "ColZero": zeros only (0 .. 1).
  - "BarWide": bars to the right of 250,000, 120,000, 50,000 in 300 x 200 px: a narrow value axis along the bottom.
  - "DateLine": a line over a date axis (months of 2024 .. 2033, 120 points, and 40 rows of "" after them), major
    unit two years, labels "yyyy" turned up (`rot="-5400000"`), Arial 9 - as vertex42's mortgage calculator.
  - "ScatterText": a scatter whose x values are a range of two columns (No., Date): Excel plots 1, 2, 3 ...
  - "PieLabels": a pie of six slices labelled with the category and the per cent, no separator, best fit, in
    Segoe UI Black 9 - as A.CRE's unit splits.
  - "PieOut": two slices, labels outside the end (`outEnd`), each label moved a little (`c:dLbl` manual layout).
  - "PieLegend": a pie with labels of their own colour per point (`c:dLbl` txPr), a legend in a box of its own
    (`c:legend` manual layout) with entries of their own colour (`c:legendEntry`) - as the investment-property
    calculator's.
  - "Skip": a line chart of 59 categories (1 .. 59) and no values, 345 x 225 px, Calibri 9: which labels are shown.
  - "Layout": columns of two series, a title, a legend at the right, the plot area filled (to see where it is).

- synthetic-chart-lines.xlsx, sheet "Charts" (each checked the same way, and with style-probe through COM):
  - "Widths": eight flat lines of 0.25 .. 4.5 pt, a legend: how many pixels a line and its legend key are.
  - "Plain2" / "Plain34" / "PlainNone" / "ScatterPlain": series stating no line under chart styles 2, 34, none.
  - "Keys9" / "Keys14": a column, a line and a line with markers: the legend's keys at 9 and 14 pt.
  - "PiePoints": slices of their own outlines (1 pt red, none, a look without a line, 4.5 pt white) under a
    series' 3 pt black; "PiePlain": a pie stating no lines.
  - "LabelsWrap" / "Labels45" / "Labels90" / "LabelsSkip": category labels Excel breaks at spaces, turns -45 and
    -90 degrees, shows every second of (60 numbers, oyabayo's cash flows).

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-chart-fixtures.py
"""
import datetime
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
OUTDIR = os.path.join(REPO, 'tests', 'test-documents-excel')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
EMU = 9525


def esc(text):
    return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def col_name(c):
    s = ''
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def ref(c, r):
    return f'{col_name(c)}{r + 1}'


# ---- chart parts -------------------------------------------------------------------------------------------

def rpr(size=None, bold=None, color=None, face=None, tag='a:defRPr'):
    attrs = (f' sz="{size}"' if size else '') + (f' b="{1 if bold else 0}"' if bold is not None else '')
    body = (f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else '')
    if face:
        body += f'<a:latin typeface="{face}"/><a:ea typeface="{face}"/><a:cs typeface="{face}"/>'
    return f'<{tag}{attrs}>{body}</{tag}>'


def txpr(size=None, bold=None, color=None, face=None, rot=None):
    body = f'<a:bodyPr rot="{rot}" vert="horz"/>' if rot is not None else '<a:bodyPr/>'
    return f'<c:txPr>{body}<a:lstStyle/><a:p><a:pPr>{rpr(size, bold, color, face)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'


def str_ref(formula, values):
    pts = ''.join(f'<c:pt idx="{i}"><c:v>{esc(v)}</c:v></c:pt>' for i, v in enumerate(values) if v is not None)
    return f'<c:strRef><c:f>{formula}</c:f><c:strCache><c:ptCount val="{len(values)}"/>{pts}</c:strCache></c:strRef>'


def num_ref(formula, values, code='General'):
    pts = ''.join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values) if isinstance(v, (int, float)))
    return (f'<c:numRef><c:f>{formula}</c:f><c:numCache><c:formatCode>{esc(code)}</c:formatCode>'
            f'<c:ptCount val="{len(values)}"/>{pts}</c:numCache></c:numRef>')


def solid(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def ser(i, name, cat, val, sppr='', extra='', key=('cat', 'val')):
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:v>{esc(name)}</c:v></c:tx>'
            + (f'<c:spPr>{sppr}</c:spPr>' if sppr else '') + extra
            + (f'<c:{key[0]}>{cat}</c:{key[0]}>' if cat else '') + f'<c:{key[1]}>{val}</c:{key[1]}></c:ser>')


GRID = '<c:majorGridlines><c:spPr><a:ln w="9525">' + solid('D9D9D9') + '</a:ln></c:spPr></c:majorGridlines>'
AXLINE = '<c:spPr><a:ln w="9525">' + solid('BFBFBF') + '</a:ln></c:spPr>'


def cat_ax(id_, cross, pos='b', tx='', extra='', kind='catAx', fmt=''):
    return (f'<c:{kind}><c:axId val="{id_}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
            f'<c:axPos val="{pos}"/>{fmt}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            f'{AXLINE}{tx}<c:crossAx val="{cross}"/><c:crosses val="autoZero"/>'
            + (f'<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>' if kind == 'catAx' else '')
            + extra + f'</c:{kind}>')


def val_ax(id_, cross, pos='l', tx='', fmt='<c:numFmt formatCode="General" sourceLinked="1"/>', between='between', scaling=''):
    return (f'<c:valAx><c:axId val="{id_}"/><c:scaling><c:orientation val="minMax"/>{scaling}</c:scaling><c:delete val="0"/>'
            f'<c:axPos val="{pos}"/>{GRID}{fmt}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            f'<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{tx}<c:crossAx val="{cross}"/><c:crosses val="autoZero"/>'
            f'<c:crossBetween val="{between}"/></c:valAx>')


def chart_space(plot, title=None, legend=None, text=None, plot_sppr='<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>',
                plot_layout='<c:layout/>', border=True, style=2, rounded=False, user_shapes=False):
    t = ''
    if title and title.startswith('<c:title>'):
        t = title + '<c:autoTitleDeleted val="0"/>'
    elif title:
        t = (f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>{rpr(1400, False, "595959")}</a:pPr><a:r><a:rPr lang="en-US"/>'
             f'<a:t>{esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>')
    else:
        t = '<c:autoTitleDeleted val="1"/>'
    frame = border if isinstance(border, str) else ('<c:spPr>' + solid('FFFFFF') + '<a:ln w="9525">' + solid('D9D9D9') + '</a:ln></c:spPr>') if border else \
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
    return (DECL + f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{REL}"><c:date1904 val="0"/><c:lang val="en-US"/>'
            + (f'<c:roundedCorners val="{1 if rounded else 0}"/>' if rounded is not None else '') + (f'<c:style val="{style}"/>' if style else '') + '<c:chart>' + t + f'<c:plotArea>{plot_layout}{plot}{plot_sppr}</c:plotArea>'
            + (legend or '') + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' + frame
            + (text if text is not None else txpr(900, None, '595959'))
            + ('<c:userShapes r:id="rId1"/>' if user_shapes else '') + '</c:chartSpace>')


def bar_plot(series, direction='col', axes=(1, 2), gap=219, overlap=-27):
    return (f'<c:barChart><c:barDir val="{direction}"/><c:grouping val="clustered"/><c:varyColors val="0"/>{series}'
            f'<c:gapWidth val="{gap}"/><c:overlap val="{overlap}"/><c:axId val="{axes[0]}"/><c:axId val="{axes[1]}"/></c:barChart>')


def column_chart(cats, values, title=None, legend=None, plot_sppr=None, more=None):
    n = len(cats)
    series = ser(0, 'S1', str_ref(f'Charts!$A$2:$A${n + 1}', cats), num_ref(f'Charts!$B$2:$B${n + 1}', values),
                 solid('4472C4'))
    if more:
        series += ser(1, 'S2', str_ref(f'Charts!$A$2:$A${n + 1}', cats), num_ref(f'Charts!$C$2:$C${n + 1}', more), solid('ED7D31'))
    kw = {'plot_sppr': plot_sppr} if plot_sppr else {}
    return chart_space(bar_plot(series) + cat_ax(1, 2) + val_ax(2, 1), title, legend, **kw)


def legend(pos='r', layout='', entries='', text='', sppr=''):
    return f'<c:legend><c:legendPos val="{pos}"/>{entries}{layout}<c:overlay val="0"/>{sppr}{text}</c:legend>'


def manual(x, y, w=None, h=None, inner=False):
    return ('<c:layout><c:manualLayout>' + ('<c:layoutTarget val="inner"/>' if inner else '')
            + f'<c:xMode val="edge"/><c:yMode val="edge"/><c:x val="{x}"/><c:y val="{y}"/>'
            + (f'<c:w val="{w}"/><c:h val="{h}"/>' if w is not None else '') + '</c:manualLayout></c:layout>')


def labels(cat=False, val=False, pct=False, pos=None, text='', points='', sep=None):
    return ('<c:dLbls>' + points + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' + text
            + (f'<c:dLblPos val="{pos}"/>' if pos else '')
            + f'<c:showLegendKey val="0"/><c:showVal val="{int(val)}"/><c:showCatName val="{int(cat)}"/><c:showSerName val="0"/>'
            + f'<c:showPercent val="{int(pct)}"/><c:showBubbleSize val="0"/>' + (f'<c:separator>{esc(sep)}</c:separator>' if sep else '')
            + '<c:showLeaderLines val="0"/></c:dLbls>')


def pie_chart(first_row, cats, values, colors, lbls, leg=None, layout='<c:layout/>', angle=0, text=None):
    n = len(cats)
    points = ''.join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr>{solid(c)}<a:ln w="19050">{solid("FFFFFF")}</a:ln></c:spPr></c:dPt>'
                     for i, c in enumerate(colors))
    series = ser(0, 'Pie', str_ref(f'Charts!$E${first_row}:$E${first_row + n - 1}', cats),
                 num_ref(f'Charts!$F${first_row}:$F${first_row + n - 1}', values), extra=points + lbls)
    plot = f'<c:pieChart><c:varyColors val="1"/>{series}<c:firstSliceAng val="{angle}"/></c:pieChart>'
    return chart_space(plot, None, leg, text=text, plot_layout=layout, border=False)


# ---- the workbook ------------------------------------------------------------------------------------------

def cell(c, r, value, style=0):
    if value is None:
        return ''
    if isinstance(value, str):
        return f'<c r="{ref(c, r)}" t="inlineStr"{f" s={chr(34)}{style}{chr(34)}" if style else ""}><is><t xml:space="preserve">{esc(value)}</t></is></c>'
    return f'<c r="{ref(c, r)}"{f" s={chr(34)}{style}{chr(34)}" if style else ""}><v>{value}</v></c>'


def excel_date(d):
    return (d - datetime.date(1899, 12, 30)).days


def build():
    rows = {}

    def put(c, r, v, s=0):
        rows.setdefault(r, {})[c] = (v, s)

    put(0, 0, 'Cat'); put(1, 0, 'S1'); put(2, 0, 'S2')
    for i, (cat, a, b) in enumerate([('A', 3, 4), ('B', 7, 2), ('C', 5, 6), ('D', 6, 5)]):
        put(0, i + 1, cat); put(1, i + 1, a); put(2, i + 1, b)
    # Pie data in E / F from row 2.
    pie_rows = {}
    blocks = [('PieLabels', ['Studio', 'One BR', 'One BR + Den', 'Two BR/One BA', 'Two BR/Two BA', 'Three BR/Two BA'], [15, 30, 20, 15, 15, 10]),
              ('PieOut', ['Occupied Units:', 'Vacant Units:'], [79, 26]),
              ('PieLegend', ['Taxman', 'Tenant', 'Investors'], [0.133, 0.7987, 0.0683])]
    r = 1
    for name, cats, vals in blocks:
        pie_rows[name] = r + 1
        for cat, v in zip(cats, vals):
            put(4, r, cat); put(5, r, v)
            r += 1
        r += 1
    # Columns H.. : other series data.
    put(7, 0, 'Close'); put(8, 0, 'Zero'); put(9, 0, 'Wide')
    for i, (c1, z, w) in enumerate([(50, 0, 250000), (60, 0, 120000), (55, 0, 50000), (52, 0, None)]):
        put(7, i + 1, c1); put(8, i + 1, z)
        if w is not None:
            put(9, i + 1, w)
    # Dates in L (No.), M (Date), N (Balance), from row 2: 120 months, then 40 rows of "".
    put(11, 0, 'No.'); put(12, 0, 'Date'); put(13, 0, 'Balance'); put(14, 0, 'Rate')
    start = datetime.date(2024, 1, 1)
    dates = []
    for i in range(160):
        if i < 120:
            d = datetime.date(start.year + (start.month - 1 + i) // 12, (start.month - 1 + i) % 12 + 1, 1)
            dates.append(excel_date(d))
            put(11, i + 1, i + 1); put(12, i + 1, excel_date(d), 1); put(13, i + 1, round(250000 * (1 - (i / 120) ** 1.3), 2)); put(14, i + 1, 0.055)
        else:
            dates.append('')
            put(11, i + 1, ''); put(12, i + 1, ''); put(13, i + 1, ''); put(14, i + 1, '')
    # 59 categories in Q.
    put(16, 0, 'N')
    for i in range(59):
        put(16, i + 1, i + 1)

    charts = []   # (name, row, col, width px, height px, xml)
    cats4 = ['A', 'B', 'C', 'D']
    charts.append(('ColScale', 1, 19, 480, 288, column_chart(cats4, [3, 7, 5, 6])))
    charts.append(('ColShort', 17, 19, 480, 110, column_chart(cats4, [3, 7, 5, 6])))
    close = chart_space(bar_plot(ser(0, 'Close', str_ref('Charts!$A$2:$A$5', cats4), num_ref('Charts!$H$2:$H$5', [50, 60, 55, 52]), solid('4472C4')))
                        + cat_ax(1, 2) + val_ax(2, 1))
    charts.append(('ColClose', 24, 19, 480, 288, close))
    zero = chart_space(bar_plot(ser(0, 'Zero', str_ref('Charts!$A$2:$A$5', cats4), num_ref('Charts!$I$2:$I$5', [0, 0, 0, 0]), solid('4472C4')))
                       + cat_ax(1, 2) + val_ax(2, 1))
    charts.append(('ColZero', 40, 19, 480, 288, zero))
    wide = chart_space(bar_plot(ser(0, 'Wide', str_ref('Charts!$A$2:$A$4', cats4[:3]), num_ref('Charts!$J$2:$J$4', [250000, 120000, 50000]), solid('4472C4')), 'bar')
                       + cat_ax(1, 2, 'l') + val_ax(2, 1, 'b', fmt='<c:numFmt formatCode="#,##0" sourceLinked="0"/>'))
    charts.append(('BarWide', 56, 19, 300, 200, wide))
    arial9 = txpr(900, False, '000000', 'Arial')
    date_series = ser(0, 'Balance', str_ref('Charts!$M$2:$M$161', [None] * 160), num_ref('Charts!$N$2:$N$161', [None] * 160, '#,##0.00'),
                      '<a:ln w="38100">' + solid('000080') + '</a:ln>', '<c:marker><c:symbol val="none"/></c:marker>')
    date_ax = ('<c:dateAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
               '<c:numFmt formatCode="yyyy" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
               '<c:spPr><a:ln w="3175">' + solid('000000') + '</a:ln></c:spPr>' + txpr(900, False, '000000', 'Arial', -5400000)
               + '<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="months"/>'
               '<c:majorUnit val="2"/><c:majorTimeUnit val="years"/><c:minorUnit val="1"/><c:minorTimeUnit val="years"/></c:dateAx>')
    date_val = ('<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'
                '<c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
                '<c:spPr><a:ln w="3175">' + solid('000000') + '</a:ln></c:spPr>' + arial9 + '<c:crossAx val="1"/><c:crosses val="autoZero"/>'
                '<c:crossBetween val="between"/></c:valAx>')
    line = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{date_series}<c:marker val="1"/>'
            '<c:axId val="1"/><c:axId val="2"/></c:lineChart>')
    charts.append(('DateLine', 72, 19, 400, 300, chart_space(line + date_ax + date_val, text=txpr(900, False, '000000', 'Arial'))))
    xval = ('<c:multiLvlStrRef><c:f>Charts!$L$2:$M$161</c:f><c:multiLvlStrCache><c:ptCount val="160"/><c:lvl>'
            + ''.join(f'<c:pt idx="{i}"><c:v>{i + 1}</c:v></c:pt>' for i in range(120)) + '</c:lvl></c:multiLvlStrCache></c:multiLvlStrRef>')
    scatter_series = ser(0, 'Rate', xval, num_ref('Charts!$O$2:$O$161', [None] * 160, '0.00%'),
                         '<a:ln w="25400">' + solid('000080') + '</a:ln>', '<c:marker><c:symbol val="none"/></c:marker>', ('xVal', 'yVal'))
    scatter = (f'<c:scatterChart><c:scatterStyle val="smoothMarker"/><c:varyColors val="0"/>{scatter_series}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>'
               + val_ax(1, 2, 'b', arial9, '<c:numFmt formatCode="General" sourceLinked="1"/>', 'midCat')
               + val_ax(2, 1, 'l', arial9, '<c:numFmt formatCode="0.00%" sourceLinked="0"/>', 'midCat', '<c:min val="0"/>'))
    charts.append(('ScatterText', 72, 27, 300, 220, chart_space(scatter, text=txpr(900, False, '000000', 'Arial'))))
    segoe = txpr(900, False, '000000', 'Segoe UI Black')
    pl = pie_rows['PieLabels']
    charts.append(('PieLabels', 94, 19, 300, 300,
                   pie_chart(pl, blocks[0][1], blocks[0][2], ['1A5C8A', 'ED7D31', '1B6B2A', '1AA0DA', '9E2A9A', '4EA72E'],
                             labels(cat=True, pct=True, pos='bestFit', text=segoe), angle=29, text=segoe)))
    po = pie_rows['PieOut']
    out_points = ('<c:dLbl><c:idx val="0"/><c:layout><c:manualLayout><c:x val="0.05"/><c:y val="-0.04"/></c:manualLayout></c:layout>'
                  '<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/>'
                  '<c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbl>')
    charts.append(('PieOut', 94, 27, 300, 300,
                   pie_chart(po, blocks[1][1], blocks[1][2], ['3B7D23', 'C55A11'], labels(cat=True, pct=True, pos='outEnd', text=segoe, points=out_points),
                             angle=29, text=segoe)))
    pg = pie_rows['PieLegend']
    colors = ['333399', '008000', '800080']
    point_labels = ''.join(f'<c:dLbl><c:idx val="{i}"/><c:numFmt formatCode="0.00%" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
                           + txpr(1000, True, c, 'Arial') + '<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
                           '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>' for i, c in enumerate(colors))
    entries = ''.join(f'<c:legendEntry><c:idx val="{i}"/>{txpr(920, True, c, "Arial")}</c:legendEntry>' for i, c in enumerate(colors))
    pie_lbls = labels(val=True, pos='outEnd', text=txpr(1000, False, '000000', 'Arial')).replace('<c:dLbls>', '<c:dLbls>' + point_labels)
    charts.append(('PieLegend', 94, 35, 450, 300,
                   pie_chart(pg, blocks[2][1], blocks[2][2], colors, pie_lbls,
                             legend('r', manual(0.6776, 0.3128, 0.3104, 0.674), entries, txpr(920, False, '000000', 'Arial')),
                             manual(0.0985, 0.1542, 0.4896, 0.7225, True), text=txpr(1000, False, '000000', 'Arial'))))
    skip_series = ser(0, 'Flow', num_ref('Charts!$Q$2:$Q$60', list(range(1, 60))), num_ref('Charts!$R$2:$R$60', [None] * 59), solid('4472C4'))
    skip = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{skip_series}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>'
            + cat_ax(1, 2) + val_ax(2, 1))
    charts.append(('Skip', 118, 19, 345, 225, chart_space(skip, 'Cumulative Net Cash Flow')))
    lay = column_chart(cats4, [3, 7, 5, 6], 'Chart Title', legend('r'), '<c:spPr>' + solid('FFF2CC') + '</c:spPr>', [4, 2, 6, 5])
    charts.append(('Layout', 118, 27, 480, 288, lay))

    put(18, 0, 'Charts:')
    write_book('synthetic-chart-parity.xlsx', rows, charts)


def write_book(book_name, rows, charts):
    """A workbook of one sheet "Charts" of these cells ({row: {col: (value, style)}}) and charts (name, row, col,
    width px, height px, chart part)."""
    sheet_rows = ''.join(f'<row r="{r + 1}">' + ''.join(cell(c, r, v, s) for c, (v, s) in sorted(cs.items())) + '</row>'
                         for r, cs in sorted(rows.items()))
    sheet_xml = (DECL + f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>'
                 '<sheetFormatPr defaultRowHeight="15"/><cols><col min="20" max="40" width="9.140625"/></cols>'
                 f'<sheetData>{sheet_rows}</sheetData><drawing r:id="rId1"/></worksheet>')
    anchors = ''
    for i, (name, r, c, w, h, *_) in enumerate(charts):
        anchors += (f'<xdr:oneCellAnchor><xdr:from><xdr:col>{c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                    f'<xdr:ext cx="{w * EMU}" cy="{h * EMU}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{i + 2}" name="{name}"/>'
                    '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                    f'<a:graphic><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="rId{i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame>'
                    '<xdr:clientData/></xdr:oneCellAnchor>')
    drawing = DECL + f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{REL}">{anchors}</xdr:wsDr>'
    styles = (DECL + f'<styleSheet xmlns="{MAIN}"><numFmts count="1"><numFmt numFmtId="164" formatCode="m/d/yyyy"/></numFmts>'
              '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
              '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}.worksheet+xml"/>'
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            + ''.join(f'<Override PartName="/xl/charts/chart{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                      for i in range(len(charts))) + '</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Charts" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': DECL + (f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
                                              f'<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
        'xl/worksheets/sheet1.xml': sheet_xml,
        'xl/worksheets/_rels/sheet1.xml.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
        'xl/drawings/drawing1.xml': drawing,
        'xl/drawings/_rels/drawing1.xml.rels': DECL + (f'<Relationships xmlns="{PKG}">' + ''.join(
            f'<Relationship Id="rId{i + 1}" Type="{REL}/chart" Target="../charts/chart{i + 1}.xml"/>' for i in range(len(charts))) + '</Relationships>'),
    }
    for i, (_, _, _, _, _, xml, *more) in enumerate(charts):
        files[f'xl/charts/chart{i + 1}.xml'] = xml
        if more and more[0]:
            # (The chart's shapes: a drawing part of its own, `c:userShapes`.)
            files[f'xl/charts/_rels/chart{i + 1}.xml.rels'] = DECL + (
                f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/chartUserShapes" Target="../drawings/drawing{i + 101}.xml"/></Relationships>')
            files[f'xl/drawings/drawing{i + 101}.xml'] = more[0]
            files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>',
                f'<Override PartName="/xl/drawings/drawing{i + 101}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"/></Types>')
    path = os.path.join(OUTDIR, book_name)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(path)


def build_lines():
    """synthetic-chart-lines.xlsx: how wide Excel 2019 draws lines and legend keys, a pie's own slice outlines."""
    rows = {}

    def put(c, r, v, s=0):
        rows.setdefault(r, {})[c] = (v, s)

    cats = ['A', 'B', 'C', 'D', 'E', 'F']
    for i, cat in enumerate(cats):
        put(0, i + 1, cat)
    widths = [3175, 6350, 9525, 12700, 19050, 28575, 38100, 57150]
    colors = ['000000', 'C00000', '0070C0', '00B050', '7030A0', 'ED7D31', '000080', 'BF8F00']
    for k, w in enumerate(widths):
        put(1 + k, 0, f'w{w}')
        for i in range(len(cats)):
            put(1 + k, i + 1, 8 - k)
    none = '<c:marker><c:symbol val="none"/></c:marker>'
    cat_ref = str_ref(f'Charts!$A$2:$A${len(cats) + 1}', cats)

    def col(k):
        return col_name(1 + k)

    def series(k, sppr=''):
        return ser(k, f'w{widths[k]}', cat_ref, num_ref(f'Charts!${col(k)}$2:${col(k)}${len(cats) + 1}', [8 - k] * len(cats)), sppr, none)

    def line_plot(body):
        return (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{body}<c:marker val="1"/>'
                '<c:axId val="1"/><c:axId val="2"/></c:lineChart>' + cat_ax(1, 2) + val_ax(2, 1))

    charts = []
    # Eight series of the widths stated, 0.25 .. 4.5 pt, legend at the right.
    stated = ''.join(series(k, f'<a:ln w="{w}" cap="rnd">{solid(colors[k])}<a:round/></a:ln>') for k, w in enumerate(widths))
    charts.append(('Widths', 1, 12, 480, 360, chart_space(line_plot(stated), legend=legend('r'))))
    # Three series that state no line, under chart styles 2, 34, none; and a scatter.
    plain = ''.join(series(k) for k in range(3))
    charts.append(('Plain2', 1, 21, 400, 260, chart_space(line_plot(plain), legend=legend('r'))))
    charts.append(('Plain34', 20, 21, 400, 260, chart_space(line_plot(plain), legend=legend('r'), style=34)))
    charts.append(('PlainNone', 39, 21, 400, 260, chart_space(line_plot(plain), legend=legend('r'), style=None)))
    xy = ''.join(ser(k, f'w{widths[k]}', num_ref(f'Charts!$J$2:$J$7', [1, 2, 3, 4, 5, 6]),
                     num_ref(f'Charts!${col(k)}$2:${col(k)}${len(cats) + 1}', [8 - k] * len(cats)), '', none, ('xVal', 'yVal')) for k in range(3))
    for i in range(6):
        put(9, i + 1, i + 1)
    scatter = (f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{xy}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>'
               + val_ax(1, 2, 'b', between='midCat') + val_ax(2, 1, 'l', between='midCat'))
    charts.append(('ScatterPlain', 58, 21, 400, 260, chart_space(scatter, legend=legend('r'))))
    # Legend keys of a line, a column, a line with markers, at 9 and 14 pt.
    bars = bar_plot(ser(0, 'Column', cat_ref, num_ref('Charts!$B$2:$B$7', [8] * 6), solid('4472C4')))
    lines = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
             + ser(1, 'Line', cat_ref, num_ref('Charts!$C$2:$C$7', [7] * 6), f'<a:ln w="28575">{solid("ED7D31")}</a:ln>', none)
             + ser(2, 'Marked', cat_ref, num_ref('Charts!$D$2:$D$7', [6] * 6), f'<a:ln w="28575">{solid("00B050")}</a:ln>',
                   f'<c:marker><c:symbol val="square"/><c:size val="7"/><c:spPr>{solid("00B050")}</c:spPr></c:marker>')
             + '<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>')
    for name, size, row in [('Keys9', 900, 20), ('Keys14', 1400, 39)]:
        charts.append((name, row, 12, 480, 260, chart_space(bars + lines + cat_ax(1, 2) + val_ax(2, 1),
                                                            legend=legend('r', text=txpr(size, None, '595959')))))
    # A pie whose points state their own outlines: 1 pt red, none, nothing (the series' 3 pt black?), 4.5 pt white.
    for i, v in enumerate([30, 25, 25, 20]):
        put(10, i + 1, f'P{i + 1}'); put(11, i + 1, v)
    pts = [f'<a:ln w="12700">{solid("FF0000")}</a:ln>', '<a:ln><a:noFill/></a:ln>', '', f'<a:ln w="57150">{solid("FFFFFF")}</a:ln>']
    fills = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000']
    points = ''.join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr>{solid(fills[i])}{ln}</c:spPr></c:dPt>' for i, ln in enumerate(pts))
    pie = ser(0, 'Pie', str_ref('Charts!$K$2:$K$5', ['P1', 'P2', 'P3', 'P4']), num_ref('Charts!$L$2:$L$5', [30, 25, 25, 20]),
              f'<a:ln w="38100">{solid("000000")}</a:ln>', points)
    charts.append(('PiePoints', 58, 12, 360, 300, chart_space(f'<c:pieChart><c:varyColors val="1"/>{pie}<c:firstSliceAng val="0"/></c:pieChart>',
                                                              legend=legend('r'), border=False)))
    # The same pie stating no lines at all (style 2).
    plain_pie = ser(0, 'Pie', str_ref('Charts!$K$2:$K$5', ['P1', 'P2', 'P3', 'P4']), num_ref('Charts!$L$2:$L$5', [30, 25, 25, 20]))
    charts.append(('PiePlain', 77, 12, 360, 300, chart_space(f'<c:pieChart><c:varyColors val="1"/>{plain_pie}<c:firstSliceAng val="0"/></c:pieChart>',
                                                             legend=legend('r'), border=False)))
    # Category labels that do not fit in a line (Calibri 9): broken at spaces, turned -45 and -90 degrees, every
    # second of 60 numbers (as oyabayo's cash flows).
    for i in range(60):
        put(13, i + 1, f'Cat {i + 1}'); put(14, i + 1, f'Category {i + 1}'); put(15, i + 1, i + 1); put(16, i + 1, 1 + (i + 1) % 3)
    for name, row, col, letter, n, w, h in [('LabelsWrap', 1, 30, 'N', 6, 157, 293), ('Labels45', 1, 34, 'O', 6, 133, 293),
                                            ('Labels90', 20, 30, 'O', 12, 236, 293), ('LabelsSkip', 20, 34, 'P', 60, 462, 303)]:
        texts = [rows[r + 1][13 + 'NOP'.index(letter)][0] for r in range(n)]
        cat = (str_ref if letter != 'P' else num_ref)(f'Charts!${letter}$2:${letter}${n + 1}', texts)
        bars = bar_plot(ser(0, 'S', cat, num_ref(f'Charts!$Q$2:$Q${n + 1}', [1 + (i + 1) % 3 for i in range(n)]), solid('4472C4')))
        charts.append((name, row, col, w, h, chart_space(bars + cat_ax(1, 2) + val_ax(2, 1))))
    write_book('synthetic-chart-lines.xlsx', rows, charts)


CDR = 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing'


def text_box(anchor, sp_id, name, paragraphs, link=None, fill='<a:noFill/>', line='<a:ln><a:noFill/></a:ln>', anchor_at='t'):
    """A text box of a chart's drawing (`cdr:sp`): paragraphs of runs (text, size in hundredths, colour, bold)."""
    body = ''.join('<a:p><a:pPr algn="ctr"/>' + ''.join(
        (f'<a:fld id="{{0F0F0F0F-0000-0000-0000-00000000000{sp_id}}}" type="TxLink">' if link else '<a:r>')
        + f'<a:rPr lang="en-US" sz="{size}" b="{1 if bold else 0}">{solid(color)}<a:latin typeface="Calibri"/></a:rPr>'
        + (('<a:pPr/>' if link else '') + f'<a:t>{esc(text)}</a:t>') + ('</a:fld>' if link else '</a:r>')
        for text, size, color, bold in runs) + '</a:p>' for runs in paragraphs)
    return (anchor[0] + f'<cdr:sp macro="" textlink="{esc(link or "")}"><cdr:nvSpPr><cdr:cNvPr id="{sp_id}" name="{name}"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>'
            f'<cdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}{line}</cdr:spPr>'
            f'<cdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="{anchor_at}"/><a:lstStyle/>{body}</cdr:txBody></cdr:sp>' + anchor[1])


def rel_anchor(x0, y0, x1, y1):
    return (f'<cdr:relSizeAnchor><cdr:from><cdr:x>{x0}</cdr:x><cdr:y>{y0}</cdr:y></cdr:from><cdr:to><cdr:x>{x1}</cdr:x><cdr:y>{y1}</cdr:y></cdr:to>',
            '</cdr:relSizeAnchor>')


def abs_anchor(x0, y0, cx, cy):
    return (f'<cdr:absSizeAnchor><cdr:from><cdr:x>{x0}</cdr:x><cdr:y>{y0}</cdr:y></cdr:from><cdr:ext cx="{cx}" cy="{cy}"/>', '</cdr:absSizeAnchor>')


def build_looks():
    """synthetic-chart-looks.xlsx: what Excel 2019 makes of a chart's own looks - the labels' typeface (a line of
    it), shapes over a chart (`c:userShapes`), titles of several runs, a legend's and the plot area's fill and
    border, rounded corners, the automatic markers, area and line charts on the ticks."""
    rows = {}

    def put(c, r, v, s=0):
        rows.setdefault(r, {})[c] = (v, s)

    cats = ['North', 'South', 'East', 'West', 'Central']
    vals = [[2400, 1300, 2650, 900, 1800], [1500, 2100, 1200, 1700, 2500], [800, 1100, 1900, 2300, 600]]
    put(0, 0, 'Region'); put(1, 0, 'Sales'); put(2, 0, 'Cost'); put(3, 0, 'Other')
    for i, cat in enumerate(cats):
        put(0, i + 1, cat)
        for k in range(3):
            put(1 + k, i + 1, vals[k][i])
    cat_ref = str_ref('Charts!$A$2:$A$6', cats)

    def col_series(k, color=None):
        letter = 'BCD'[k]
        return ser(k, ['Sales', 'Cost', 'Other'][k], cat_ref, num_ref(f'Charts!${letter}$2:${letter}$6', vals[k]), solid(color) if color else '')

    charts = []
    # The labels' typeface: the same columns in Calibri, Garamond, Tw Cen MT, Segoe UI and Arial 9 pt, 400 x 260 px.
    for k, face in enumerate(['Calibri', 'Garamond', 'Tw Cen MT', 'Segoe UI', 'Arial']):
        charts.append((f'Face{face.replace(" ", "")}', 1 + 14 * k, 6, 400, 260,
                       chart_space(bar_plot(col_series(0, '4472C4')) + cat_ax(1, 2) + val_ax(2, 1), text=txpr(900, False, '595959', face))))
    # A doughnut with text boxes over it: one from a corner to a corner, linked to a cell; one of a size of its own.
    for i, (cat, v) in enumerate([('Fashions Direct', 71), ('Next Look', 29)]):
        put(4, i + 1, cat); put(5, i + 1, v)
    put(4, 4, 'Fashions Direct 71%')
    dpts = ''.join(f'<c:dPt><c:idx val="{i}"/><c:bubble3D val="0"/><c:spPr>{solid(c)}<a:ln w="19050">{solid("FFFFFF")}</a:ln></c:spPr></c:dPt>'
                   for i, c in enumerate(['4472C4', 'ED7D31']))
    dough = ser(0, 'Share', str_ref('Charts!$E$2:$E$3', ['Fashions Direct', 'Next Look']), num_ref('Charts!$F$2:$F$3', [71, 29]), extra=dpts)
    shapes = (DECL + f'<c:userShapes xmlns:c="{C}" xmlns:a="{A}" xmlns:cdr="{CDR}">'
              + text_box(rel_anchor(0.4841, 0.66949, 1, 1), 2, 'TextBox 1', [[('Fashions Direct 71%', 900, '4472C4', False)]], link="Charts!$E$5")
              + text_box(rel_anchor(0.00752, 0.02273, 0.44463, 0.42273), 3, 'TextBox 2',
                         [[('Next Look', 1200, 'ED7D31', True)], [('29%', 900, '595959', False)]])
              + text_box(abs_anchor(0.35, 0.42, 1143000, 381000), 4, 'TextBox 3', [[('Centre', 1400, 'C00000', True)]],
                         fill=solid('FFF2CC'), line=f'<a:ln w="12700">{solid("C00000")}</a:ln>', anchor_at='ctr')
              + '</c:userShapes>')
    charts.append(('UserShapes', 1, 14, 360, 260,
                   chart_space(f'<c:doughnutChart><c:varyColors val="1"/>{dough}<c:firstSliceAng val="0"/><c:holeSize val="60"/></c:doughnutChart>',
                               border=False, user_shapes=True), shapes))
    # A title of several runs (sizes, colours, bold, italic) over two lines; an axis title of two runs.
    runs = [('Sales ', 1800, 'C00000', True, False), ('by ', 1200, '595959', False, True), ('Region', 1800, '0070C0', True, False)]
    title_xml = ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>' + rpr(1400, False, '595959') + '</a:pPr>'
                 + ''.join(f'<a:r><a:rPr lang="en-US" sz="{sz}" b="{int(b)}" i="{int(it)}">{solid(c)}</a:rPr><a:t>{esc(t)}</a:t></a:r>' for t, sz, c, b, it in runs)
                 + '</a:p><a:p><a:pPr>' + rpr(1000, False, '595959') + '</a:pPr><a:r><a:rPr lang="en-US" sz="1000"/><a:t>second line, 10 pt</a:t></a:r></a:p>'
                 '</c:rich></c:tx><c:overlay val="0"/></c:title>')
    ax_title = ('<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:pPr>' + rpr(1000, True, '595959') + '</a:pPr>'
                f'<a:r><a:rPr lang="en-US" sz="1000" b="1">{solid("595959")}</a:rPr><a:t>Sales </a:t></a:r>'
                f'<a:r><a:rPr lang="en-US" sz="1400" b="0" i="1">{solid("00B050")}</a:rPr><a:t>(USD)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')
    val_titled = val_ax(2, 1).replace('<c:numFmt', ax_title + '<c:numFmt', 1)
    charts.append(('RichTitle', 1, 22, 480, 300, chart_space(bar_plot(col_series(0, '4472C4')) + cat_ax(1, 2) + val_titled, title_xml)))
    # A legend with a fill and a border; the plot area with a fill and a border; the chart's corners rounded,
    # a 2.25 pt border.
    leg_sppr = f'<c:spPr>{solid("FFF2CC")}<a:ln w="19050">{solid("C00000")}</a:ln></c:spPr>'
    plot_sppr = f'<c:spPr>{solid("DDEBF7")}<a:ln w="12700">{solid("0070C0")}</a:ln></c:spPr>'
    frame = f'<c:spPr>{solid("FFFFFF")}<a:ln w="28575">{solid("7030A0")}</a:ln></c:spPr>'
    for name, pos, row in [('LegendRight', 'r', 20), ('LegendBottom', 'b', 40)]:
        charts.append((name, row, 22, 480, 300,
                       chart_space(bar_plot(col_series(0, '4472C4') + col_series(1, 'ED7D31') + col_series(2, 'A5A5A5')) + cat_ax(1, 2) + val_ax(2, 1),
                                   'Boxes', legend(pos, sppr=leg_sppr), plot_sppr=plot_sppr, border=frame, rounded=(name == 'LegendRight'))))
    # Corners: no roundedCorners element at all, and Excel's grey frame (no spPr).
    bare = chart_space(bar_plot(col_series(0, '4472C4')) + cat_ax(1, 2) + val_ax(2, 1), rounded=None, border='')
    charts.append(('RoundedBare', 60, 22, 300, 200, bare))
    # Automatic markers: eight line series and eight scatter series that state nothing of their markers.
    for k in range(8):
        put(8 + k, 0, f'M{k + 1}')
        for i in range(5):
            put(8 + k, i + 1, 10 - k + (i % 2))
    for i in range(5):
        put(6, i + 1, i + 1)

    def plain(k, xy=False):
        letter = col_name(8 + k)
        values = num_ref(f'Charts!${letter}$2:${letter}$6', [10 - k + (i % 2) for i in range(5)])
        return ser(k, f'M{k + 1}', num_ref('Charts!$G$2:$G$6', [1, 2, 3, 4, 5]) if xy else cat_ref, values, '', '', ('xVal', 'yVal') if xy else ('cat', 'val'))
    marked = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{"".join(plain(k) for k in range(8))}<c:marker val="1"/>'
              '<c:axId val="1"/><c:axId val="2"/></c:lineChart>' + cat_ax(1, 2) + val_ax(2, 1))
    charts.append(('Markers', 1, 30, 480, 300, chart_space(marked, legend=legend('r'))))
    charts.append(('Markers34', 20, 30, 480, 300, chart_space(marked, legend=legend('r'), style=34)))
    charts.append(('Markers42', 40, 30, 480, 300, chart_space(marked, legend=legend('r'), style=42)))
    xy = (f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{"".join(plain(k, True) for k in range(8))}'
          '<c:axId val="1"/><c:axId val="2"/></c:scatterChart>' + val_ax(1, 2, 'b', between='midCat') + val_ax(2, 1, 'l', between='midCat'))
    charts.append(('MarkersXY', 60, 30, 480, 300, chart_space(xy, legend=legend('r'))))
    # Area and line charts on the ticks (crossBetween midCat) with long labels; the same between.
    months = ['January 2024', 'February 2024', 'March 2024', 'April 2024', 'May 2024', 'June 2024']
    for i, label in enumerate(months):
        put(17, i + 1, label); put(18, i + 1, [3, 5, 4, 6, 5, 7][i])
    area_series = ser(0, 'Area', str_ref('Charts!$R$2:$R$7', months), num_ref('Charts!$S$2:$S$7', [3, 5, 4, 6, 5, 7]), solid('4472C4'))
    for name, kind, between, row in [('AreaTicks', 'area', 'midCat', 1), ('AreaBetween', 'area', 'between', 20),
                                     ('LineTicks', 'line', 'midCat', 40), ('LineBetween', 'line', 'between', 60)]:
        body = area_series if kind == 'area' else area_series.replace(solid('4472C4'), f'<a:ln w="28575">{solid("4472C4")}</a:ln>')
        plot = (f'<c:{kind}Chart><c:grouping val="standard"/><c:varyColors val="0"/>{body}' + ('<c:marker val="1"/>' if kind == 'line' else '')
                + f'<c:axId val="1"/><c:axId val="2"/></c:{kind}Chart>' + cat_ax(1, 2) + val_ax(2, 1, between=between))
        charts.append((name, row, 38, 480, 260, chart_space(plot)))
    put(5, 0, 'Charts:')
    write_book('synthetic-chart-looks.xlsx', rows, charts)


if __name__ == '__main__':
    build()
    build_lines()
    build_looks()

