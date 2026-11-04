#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-slicers.xlsx, a synthetic workbook (not saved by Excel: written
from the ECMA-376 and MS-XLSX schemas) with every kind of slicer and a timeline:

- "Data": a table (Sales: Date, Region, Product, Amount) filtered on Product to Apples and Pears, with a
  table slicer on that column ("Product 1": x15 tableSlicerCache, its buttons the column's values).
- "Report": a pivot table of the table's pivot cache (Sum of Amount by Region) with two pivot slicers in
  one slicers part - "Region" (two columns: North and South selected, East not, West without data) and
  "Product 2" (three columns, in a custom slicer style, "Slicer Custom": x14:slicerStyles for the buttons,
  the tableStyle of that name for the whole slicer and its header) - and a timeline on Date ("Date", months,
  March to December 2012 selected, in a custom timeline style), its cache's selection and bounds.

The slicer and timeline caches are in the workbook's extension list (x14:slicerCaches for the pivot ones,
x15:slicerCaches for the table's, x15:timelineCacheRefs), the slicers and the timeline in the sheets'
(x14:slicerList, the x15 table slicer list, x15:timelineRefs), their frames in the drawings as
mc:AlternateContent (sle:slicer, tsle:timeslicer) with the text box fallback Excel writes. The theme is the
Office theme of pandas-test_decimal.xlsx. Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-slicers-fixture.py
"""
import os
import zipfile
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-slicers.xlsx')
THEME_SOURCE = os.path.join(REPO, 'tests', 'test-documents-excel', 'pandas-test_decimal.xlsx')

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
SLICER_REL = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
SLICER_CACHE_REL = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
TIMELINE_REL = 'http://schemas.microsoft.com/office/2011/relationships/timeline'
TIMELINE_CACHE_REL = 'http://schemas.microsoft.com/office/2011/relationships/timelineCache'
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
PIVOT_CACHE_ID = 1234567

ROWS = [
    (date(2012, 1, 15), 'East', 'Apples', 10),
    (date(2012, 2, 10), 'North', 'Pears', 20),
    (date(2012, 3, 5), 'South', 'Apples', 30),
    (date(2012, 4, 20), 'East', 'Plums', 40),
    (date(2012, 5, 11), 'North', 'Apples', 15),
    (date(2012, 6, 30), 'South', 'Pears', 25),
    (date(2012, 8, 14), 'East', 'Apples', 35),
    (date(2012, 10, 2), 'North', 'Plums', 45),
    (date(2013, 1, 9), 'South', 'Apples', 12),
    (date(2013, 3, 17), 'East', 'Pears', 22),
    (date(2013, 6, 21), 'North', 'Apples', 32),
    (date(2013, 11, 28), 'South', 'Plums', 42),
]
REGIONS = ['East', 'North', 'South', 'West']      # West: an item of the cache without records
PRODUCTS = ['Apples', 'Pears', 'Plums']
TABLE_FILTER = ['Apples', 'Pears']                 # the table's filter on Product (the table slicer's selection)
SELECTED_REGIONS = ['North', 'South']              # the Region slicer's selection (West: selected, no data)
SELECTION = (date(2012, 3, 1), date(2012, 12, 31))  # the timeline's
BOUNDS = (date(2012, 1, 1), date(2014, 1, 1))


def serial(day):
    return (day - date(1899, 12, 30)).days


def iso(day):
    return day.isoformat() + 'T00:00:00'


strings = []


def sst(text):
    if text not in strings:
        strings.append(text)
    return strings.index(text)


def c_str(ref, text):
    return '<c r="%s" t="s"><v>%d</v></c>' % (ref, sst(text))


def c_num(ref, value, style=0):
    return '<c r="%s"%s><v>%s</v></c>' % (ref, (' s="%d"' % style) if style else '', value)


# ---- the Data sheet: the table, filtered on Product -------------------------------------------------
data_rows = ['<row r="1">%s</row>' % ''.join(c_str('%s1' % col, name) for col, name in zip('ABCD', ['Date', 'Region', 'Product', 'Amount']))]
for i, (day, region, product, amount) in enumerate(ROWS):
    r = i + 2
    hidden = ' hidden="1"' if product not in TABLE_FILTER else ''
    data_rows.append('<row r="%d"%s>%s%s%s%s</row>' % (r, hidden, c_num('A%d' % r, serial(day), 1), c_str('B%d' % r, region), c_str('C%d' % r, product), c_num('D%d' % r, amount)))
last = len(ROWS) + 1
table = (DECL + '<table xmlns="%s" id="1" name="Sales" displayName="Sales" ref="A1:D%d" totalsRowShown="0">'
         '<autoFilter ref="A1:D%d"><filterColumn colId="2"><filters>%s</filters></filterColumn></autoFilter>'
         '<tableColumns count="4"><tableColumn id="1" name="Date"/><tableColumn id="2" name="Region"/><tableColumn id="3" name="Product"/><tableColumn id="4" name="Amount"/></tableColumns>'
         '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>') % (
    MAIN, last, last, ''.join('<filter val="%s"/>' % value for value in TABLE_FILTER))
data_sheet = (DECL + '<worksheet xmlns="%s" xmlns:r="%s" xmlns:mc="%s" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
              '<dimension ref="A1:D%d"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
              '<cols><col min="1" max="1" width="11" customWidth="1"/><col min="2" max="4" width="10" customWidth="1"/></cols>'
              '<sheetData>%s</sheetData>'
              '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
              '<drawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
              '<extLst><ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="%s"><x14:slicerList xmlns:x14="%s"><x14:slicer r:id="rId3"/></x14:slicerList></ext></extLst>'
              '</worksheet>') % (MAIN, REL, MC, last, ''.join(data_rows), X15, X14)
data_rels = (DECL + '<Relationships xmlns="%s">'
             '<Relationship Id="rId1" Type="%s/table" Target="../tables/table1.xml"/>'
             '<Relationship Id="rId2" Type="%s/drawing" Target="../drawings/drawing2.xml"/>'
             '<Relationship Id="rId3" Type="%s" Target="../slicers/slicer2.xml"/></Relationships>') % (PKG, REL, REL, SLICER_REL)

# ---- the pivot cache of the table, and the pivot table of the Report sheet ---------------------------
in_selection = [row for row in ROWS if SELECTION[0] <= row[0] <= SELECTION[1]]
shown = [row for row in in_selection if row[1] in SELECTED_REGIONS]
dates = sorted({row[0] for row in ROWS})
cache_fields = (
    '<cacheField name="Date" numFmtId="14"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="%s" maxDate="%s" count="%d">%s</sharedItems></cacheField>'
    % (iso(dates[0]), iso(dates[-1]), len(dates), ''.join('<d v="%s"/>' % iso(day) for day in dates))
    + '<cacheField name="Region" numFmtId="0"><sharedItems count="%d">%s</sharedItems></cacheField>' % (len(REGIONS), ''.join('<s v="%s"/>' % r for r in REGIONS))
    + '<cacheField name="Product" numFmtId="0"><sharedItems count="%d">%s</sharedItems></cacheField>' % (len(PRODUCTS), ''.join('<s v="%s"/>' % p for p in PRODUCTS))
    + '<cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="%d" maxValue="%d"/></cacheField>'
    % (min(row[3] for row in ROWS), max(row[3] for row in ROWS)))
cache_definition = (DECL + '<pivotCacheDefinition xmlns="%s" xmlns:r="%s" r:id="rId1" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="%d">'
                    '<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource><cacheFields count="4">%s</cacheFields>'
                    '<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="%s"><x14:pivotCacheDefinition pivotCacheId="%d"/></ext></extLst>'
                    '</pivotCacheDefinition>') % (MAIN, REL, len(ROWS), cache_fields, X14, PIVOT_CACHE_ID)
cache_records = (DECL + '<pivotCacheRecords xmlns="%s" xmlns:r="%s" count="%d">%s</pivotCacheRecords>') % (
    MAIN, REL, len(ROWS), ''.join('<r><x v="%d"/><x v="%d"/><x v="%d"/><n v="%d"/></r>' % (dates.index(day), REGIONS.index(region), PRODUCTS.index(product), amount)
                                 for day, region, product, amount in ROWS))
cache_rels = (DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>') % (PKG, REL)

regions_shown = [r for r in REGIONS if r in SELECTED_REGIONS]
sums = {r: sum(row[3] for row in shown if row[1] == r) for r in regions_shown}
pivot_last = 4 + len(regions_shown)
region_items = ''.join('<item x="%d"%s/>' % (i, '' if r in SELECTED_REGIONS or r == 'West' else ' h="1"') for i, r in enumerate(REGIONS))
pivot_table = (DECL + '<pivotTableDefinition xmlns="%s" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
               '<location ref="A3:B%d" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
               '<pivotFields count="4"><pivotField numFmtId="14" showAll="0"/><pivotField axis="axisRow" showAll="0"><items count="%d">%s<item t="default"/></items></pivotField>'
               '<pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>'
               '<rowFields count="1"><field x="1"/></rowFields><rowItems count="%d">%s<i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems>'
               '<dataFields count="1"><dataField name="Sum of Amount" fld="3" baseField="0" baseItem="0"/></dataFields>'
               '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
               '<filters count="1"><filter fld="0" type="dateBetween" evalOrder="-1" id="1" name="Date"><autoFilter ref="A1"><filterColumn colId="0"><customFilters and="1">'
               '<customFilter operator="greaterThanOrEqual" val="%d"/><customFilter operator="lessThanOrEqual" val="%d"/></customFilters></filterColumn></autoFilter>'
               '<extLst><ext uri="{0605FD5F-26C8-4aeb-8148-2DB25E43C511}" xmlns:x15="%s"><x15:pivotFilter useWholeDay="1"/></ext></extLst></filter></filters>'
               '</pivotTableDefinition>') % (
    MAIN, pivot_last, len(REGIONS) + 1, region_items, len(regions_shown) + 1,
    ''.join('<i><x v="%d"/></i>' % REGIONS.index(r) for r in regions_shown), serial(SELECTION[0]), serial(SELECTION[1]), X15)
pivot_rels = (DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>') % (PKG, REL)

report_rows = ['<row r="1">%s</row>' % c_str('A1', 'Sales by region (filtered by the slicers and the timeline)'),
               '<row r="3">%s%s</row>' % (c_str('A3', 'Row Labels'), c_str('B3', 'Sum of Amount'))]
for i, region in enumerate(regions_shown):
    r = 4 + i
    report_rows.append('<row r="%d">%s%s</row>' % (r, c_str('A%d' % r, region), c_num('B%d' % r, sums[region])))
report_rows.append('<row r="%d">%s%s</row>' % (pivot_last, c_str('A%d' % pivot_last, 'Grand Total'), c_num('B%d' % pivot_last, sum(sums.values()))))
report_sheet = (DECL + '<worksheet xmlns="%s" xmlns:r="%s"><dimension ref="A1:B%d"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="2" width="15" customWidth="1"/></cols>'
                '<sheetData>%s</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId2"/>'
                '<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="%s"><x14:slicerList><x14:slicer r:id="rId3"/></x14:slicerList></ext>'
                '<ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="%s"><x15:timelineRefs><x15:timelineRef r:id="rId4"/></x15:timelineRefs></ext></extLst>'
                '</worksheet>') % (MAIN, REL, pivot_last, ''.join(report_rows), X14, X15)
report_rels = (DECL + '<Relationships xmlns="%s">'
               '<Relationship Id="rId1" Type="%s/pivotTable" Target="../pivotTables/pivotTable1.xml"/>'
               '<Relationship Id="rId2" Type="%s/drawing" Target="../drawings/drawing1.xml"/>'
               '<Relationship Id="rId3" Type="%s" Target="../slicers/slicer1.xml"/>'
               '<Relationship Id="rId4" Type="%s" Target="../timelines/timeline1.xml"/></Relationships>') % (PKG, REL, REL, SLICER_REL, TIMELINE_REL)

# ---- slicers, their caches, the timeline and its cache -----------------------------------------------
SLICERS_HEAD = '<slicers xmlns="%s" xmlns:mc="%s" mc:Ignorable="x" xmlns:x="%s">' % (X14, MC, MAIN)
slicers_report = (DECL + SLICERS_HEAD
                  + '<slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="2" rowHeight="241300"/>'
                  + '<slicer name="Product 2" cache="Slicer_Product1" caption="Product (pivot)" columnCount="3" style="Slicer Custom" rowHeight="241300"/>'
                  + '</slicers>')
slicers_data = (DECL + SLICERS_HEAD + '<slicer name="Product 1" cache="Slicer_Product" caption="Product" rowHeight="241300"/></slicers>')

region_items_cache = ''.join(
    '<i x="%d"%s%s/>' % (REGIONS.index(r), ' s="1"' if r in SELECTED_REGIONS or r == 'West' else '', ' nd="1"' if r == 'West' else '')
    for r in ['North', 'South', 'East', 'West'])
CACHE_HEAD = '<slicerCacheDefinition xmlns="%s" xmlns:mc="%s" mc:Ignorable="x" xmlns:x="%s"' % (X14, MC, MAIN)
slicer_cache_region = (DECL + CACHE_HEAD + ' name="Slicer_Region" sourceName="Region"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables>'
                       '<data><tabular pivotCacheId="%d"><items count="4">%s</items></tabular></data></slicerCacheDefinition>') % (PIVOT_CACHE_ID, region_items_cache)
slicer_cache_product_pivot = (DECL + CACHE_HEAD + ' name="Slicer_Product1" sourceName="Product"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables>'
                              '<data><tabular pivotCacheId="%d"><items count="3">%s</items></tabular></data></slicerCacheDefinition>') % (
    PIVOT_CACHE_ID, ''.join('<i x="%d" s="1"/>' % i for i in range(len(PRODUCTS))))
slicer_cache_product_table = (DECL + '<slicerCacheDefinition xmlns="%s" xmlns:mc="%s" mc:Ignorable="x x15" xmlns:x="%s" xmlns:x15="%s" name="Slicer_Product" sourceName="Product">'
                              '<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"><x15:tableSlicerCache tableId="1" column="3"/></x:ext></extLst>'
                              '</slicerCacheDefinition>') % (X14, MC, MAIN, X15)
timelines = (DECL + '<timelines xmlns="%s" xmlns:mc="%s" mc:Ignorable="x" xmlns:x="%s">'
             '<timeline name="Date" cache="NativeTimeline_Date" caption="Date" level="2" selectionLevel="2" scrollPosition="%s" style="Timeline Custom"/></timelines>') % (
    X15, MC, MAIN, iso(BOUNDS[0]))
timeline_cache = (DECL + '<timelineCacheDefinition xmlns="%s" xmlns:mc="%s" mc:Ignorable="x" xmlns:x="%s" name="NativeTimeline_Date" sourceName="Date">'
                  '<pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables>'
                  '<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="%d" filterType="dateBetween">'
                  '<selection startDate="%s" endDate="%s"/><bounds startDate="%s" endDate="%s"/></state></timelineCacheDefinition>') % (
    X15, MC, MAIN, PIVOT_CACHE_ID, iso(SELECTION[0]), iso(SELECTION[1]), iso(BOUNDS[0]), iso(BOUNDS[1]))


# ---- drawings: the frames, as Excel writes them ------------------------------------------------------
def anchor(frame_id, name, cols, rows, kind):
    (c0, c1), (r0, r1) = cols, rows
    if kind == 'slicer':
        requires, ns, uri, element = 'a14', 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"', 'http://schemas.microsoft.com/office/drawing/2010/slicer', \
            '<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="%s"/>' % name
        what = 'slicer'
    else:
        requires, ns, uri, element = 'tsle', 'xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer"', 'http://schemas.microsoft.com/office/drawing/2012/timeslicer', \
            '<tsle:timeslicer name="%s"/>' % name
        what = 'timeline'
    return ('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<mc:AlternateContent xmlns:mc="%s"><mc:Choice %s Requires="%s"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="%s">%s</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
            '<mc:Fallback xmlns=""><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
            '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="2524125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill>'
            '<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
            '<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>This shape represents a %s. It is supported in Excel 2010 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>'
            '</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>') % (c0, r0, c1, r1, MC, ns, requires, frame_id, name, uri, element, what)


def drawing(anchors):
    return (DECL + '<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s">%s</xdr:wsDr>') % (XDR, A, ''.join(anchors))


drawing_report = drawing([anchor(2, 'Region', (3, 6), (1, 12), 'slicer'),
                          anchor(3, 'Product 2', (7, 12), (1, 7), 'slicer'),
                          anchor(4, 'Date', (3, 12), (13, 21), 'timeline')])
drawing_data = drawing([anchor(2, 'Product 1', (5, 8), (1, 10), 'slicer')])

# ---- styles: a date format, the custom slicer style and timeline style -------------------------------
dxfs = [
    # 0 the custom slicer's whole: a dark red outline
    '<dxf><font><color rgb="FF000000"/></font><border><left style="thin"><color rgb="FFC00000"/></left><right style="thin"><color rgb="FFC00000"/></right>'
    '<top style="thin"><color rgb="FFC00000"/></top><bottom style="thin"><color rgb="FFC00000"/></bottom></border></dxf>',
    # 1 its header: bold dark red, a line below
    '<dxf><font><b/><color rgb="FFC00000"/></font><border><bottom style="thin"><color rgb="FFC00000"/></bottom></border></dxf>',
]
x14_dxfs = [
    # 0 selectedItemWithData: orange, bold
    '<dxf><font><b/><color rgb="FF000000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor rgb="FFFFC000"/></patternFill></fill>'
    '<border><left style="thin"><color rgb="FFC00000"/></left><right style="thin"><color rgb="FFC00000"/></right><top style="thin"><color rgb="FFC00000"/></top><bottom style="thin"><color rgb="FFC00000"/></bottom></border></dxf>',
    # 1 unselectedItemWithData: pale yellow
    '<dxf><font><color rgb="FF7F6000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>',
    # 2 selectedItemWithNoData
    '<dxf><font><color rgb="FFA6A6A6"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor rgb="FFFFE699"/></patternFill></fill></dxf>',
    # 3 unselectedItemWithNoData
    '<dxf><font><color rgb="FFA6A6A6"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></dxf>',
    # 4 the custom timeline's selected periods: green
    '<dxf><fill><patternFill patternType="solid"><fgColor theme="9"/><bgColor theme="9"/></patternFill></fill></dxf>',
]
styles = (DECL + '<styleSheet xmlns="%s" xmlns:mc="%s" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
          '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
          '<dxfs count="%d">%s</dxfs>'
          '<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">'
          '<tableStyle name="Slicer Custom" pivot="0" table="0" count="2"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="1"/></tableStyle></tableStyles>'
          '<extLst>'
          '<ext uri="{46F421CA-312F-682f-3DD2-61675219B42D}" xmlns:x14="%s"><x14:dxfs count="%d">%s</x14:dxfs></ext>'
          '<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="%s"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1">'
          '<x14:slicerStyle name="Slicer Custom"><x14:slicerStyleElements><x14:slicerStyleElement type="unselectedItemWithData" dxfId="1"/>'
          '<x14:slicerStyleElement type="unselectedItemWithNoData" dxfId="3"/><x14:slicerStyleElement type="selectedItemWithData" dxfId="0"/>'
          '<x14:slicerStyleElement type="selectedItemWithNoData" dxfId="2"/></x14:slicerStyleElements></x14:slicerStyle></x14:slicerStyles></ext>'
          '<ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="%s"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1">'
          '<x15:timelineStyle name="Timeline Custom"><x15:timelineStyleElements><x15:timelineStyleElement type="selectedTimeBlock" dxfId="4"/></x15:timelineStyleElements></x15:timelineStyle>'
          '</x15:timelineStyles></ext></extLst></styleSheet>') % (
    MAIN, MC, len(dxfs), ''.join(dxfs), X14, len(x14_dxfs), ''.join(x14_dxfs), X14, X15)

shared = (DECL + '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>') % (
    MAIN, len(strings), len(strings), ''.join('<si><t>%s</t></si>' % s for s in strings))

# ---- the workbook ------------------------------------------------------------------------------------
workbook = (DECL + '<workbook xmlns="%s" xmlns:r="%s" xmlns:mc="%s" mc:Ignorable="x15" xmlns:x15="%s">'
            '<workbookPr defaultThemeVersion="164011"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300" activeTab="1"/></bookViews>'
            '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets>'
            '<definedNames><definedName name="NativeTimeline_Date">#N/A</definedName><definedName name="Slicer_Product">#N/A</definedName>'
            '<definedName name="Slicer_Product1">#N/A</definedName><definedName name="Slicer_Region">#N/A</definedName></definedNames>'
            '<calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="1" r:id="rId6"/></pivotCaches>'
            '<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="%s"><x14:slicerCaches><x14:slicerCache r:id="rId7"/><x14:slicerCache r:id="rId8"/></x14:slicerCaches></ext>'
            '<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="%s"><x15:slicerCaches xmlns:x14="%s"><x14:slicerCache r:id="rId9"/></x15:slicerCaches></ext>'
            '<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="%s"><x15:timelineCacheRefs><x15:timelineCacheRef r:id="rId10"/></x15:timelineCacheRefs></ext>'
            '</extLst></workbook>') % (MAIN, REL, MC, X15, X14, X15, X14, X15)
workbook_rels = (DECL + '<Relationships xmlns="%s">'
                 '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
                 '<Relationship Id="rId2" Type="%s/worksheet" Target="worksheets/sheet2.xml"/>'
                 '<Relationship Id="rId3" Type="%s/theme" Target="theme/theme1.xml"/>'
                 '<Relationship Id="rId4" Type="%s/styles" Target="styles.xml"/>'
                 '<Relationship Id="rId5" Type="%s/sharedStrings" Target="sharedStrings.xml"/>'
                 '<Relationship Id="rId6" Type="%s/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>'
                 '<Relationship Id="rId7" Type="%s" Target="slicerCaches/slicerCache1.xml"/>'
                 '<Relationship Id="rId8" Type="%s" Target="slicerCaches/slicerCache2.xml"/>'
                 '<Relationship Id="rId9" Type="%s" Target="slicerCaches/slicerCache3.xml"/>'
                 '<Relationship Id="rId10" Type="%s" Target="timelineCaches/timelineCache1.xml"/></Relationships>') % (
    PKG, REL, REL, REL, REL, REL, REL, SLICER_CACHE_REL, SLICER_CACHE_REL, SLICER_CACHE_REL, TIMELINE_CACHE_REL)

with zipfile.ZipFile(THEME_SOURCE) as source:
    theme = source.read('xl/theme/theme1.xml')

OVERRIDES = [
    ('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
    ('/xl/worksheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ('/xl/worksheets/sheet2.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ('/xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'),
    ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
    ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'),
    ('/xl/tables/table1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'),
    ('/xl/pivotCache/pivotCacheDefinition1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'),
    ('/xl/pivotCache/pivotCacheRecords1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'),
    ('/xl/pivotTables/pivotTable1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'),
    ('/xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
    ('/xl/drawings/drawing2.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
    ('/xl/slicers/slicer1.xml', 'application/vnd.ms-excel.slicer+xml'),
    ('/xl/slicers/slicer2.xml', 'application/vnd.ms-excel.slicer+xml'),
    ('/xl/slicerCaches/slicerCache1.xml', 'application/vnd.ms-excel.slicerCache+xml'),
    ('/xl/slicerCaches/slicerCache2.xml', 'application/vnd.ms-excel.slicerCache+xml'),
    ('/xl/slicerCaches/slicerCache3.xml', 'application/vnd.ms-excel.slicerCache+xml'),
    ('/xl/timelines/timeline1.xml', 'application/vnd.ms-excel.timeline+xml'),
    ('/xl/timelineCaches/timelineCache1.xml', 'application/vnd.ms-excel.timelineCache+xml'),
]
content_types = (DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>%s</Types>') % ''.join(
    '<Override PartName="%s" ContentType="%s"/>' % item for item in OVERRIDES)
root_rels = (DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>') % (PKG, REL)

parts = [
    ('[Content_Types].xml', content_types),
    ('_rels/.rels', root_rels),
    ('xl/workbook.xml', workbook),
    ('xl/_rels/workbook.xml.rels', workbook_rels),
    ('xl/worksheets/sheet1.xml', data_sheet),
    ('xl/worksheets/_rels/sheet1.xml.rels', data_rels),
    ('xl/worksheets/sheet2.xml', report_sheet),
    ('xl/worksheets/_rels/sheet2.xml.rels', report_rels),
    ('xl/theme/theme1.xml', theme),
    ('xl/styles.xml', styles),
    ('xl/sharedStrings.xml', shared),
    ('xl/tables/table1.xml', table),
    ('xl/pivotCache/pivotCacheDefinition1.xml', cache_definition),
    ('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', cache_rels),
    ('xl/pivotCache/pivotCacheRecords1.xml', cache_records),
    ('xl/pivotTables/pivotTable1.xml', pivot_table),
    ('xl/pivotTables/_rels/pivotTable1.xml.rels', pivot_rels),
    ('xl/drawings/drawing1.xml', drawing_report),
    ('xl/drawings/drawing2.xml', drawing_data),
    ('xl/slicers/slicer1.xml', slicers_report),
    ('xl/slicers/slicer2.xml', slicers_data),
    ('xl/slicerCaches/slicerCache1.xml', slicer_cache_region),
    ('xl/slicerCaches/slicerCache2.xml', slicer_cache_product_pivot),
    ('xl/slicerCaches/slicerCache3.xml', slicer_cache_product_table),
    ('xl/timelines/timeline1.xml', timelines),
    ('xl/timelineCaches/timelineCache1.xml', timeline_cache),
]

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as package:
    for name, data in parts:
        info = zipfile.ZipInfo(name, date_time=(2026, 11, 4, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        package.writestr(info, data if isinstance(data, bytes) else data.encode('utf-8'))
print(OUT)
