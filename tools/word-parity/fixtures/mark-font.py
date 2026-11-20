#!/usr/bin/env python3
"""Fixture family: how tall an empty paragraph's line is - the font of its mark (w:pPr/w:rPr) or its style's?

    python3 tools/word-parity/fixtures/mark-font.py tests/word/fixtures/mark-font
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/mark-font --refresh-word
    python3 tools/word-parity/fixtures/mark-font.py --table tests/word/fixtures/mark-font/word/*.pdf \
        > tests/word/fixtures/mark-font/word/heights.tsv

docx4j-getting-started's code samples (Consolas 8pt runs, single spacing, nothing after) have empty paragraphs
between them whose marks are Consolas 8pt under document defaults of the theme's minor font (Calibri): Word's
lines 9.36pt apart across them, ours 9.77 (Calibri 8pt) - its p39 one line early. Each document: Letter, 1in
margins, nothing before or after any paragraph, single spacing; "Topq", ten empty paragraphs, "Botq" (Calibri
11pt, 13.43pt lines). The empty paragraphs' line: (Botq's top - Topq's top - 13.43) / 10.
Names: <defaults>-<mark>[-<where>][-c14]. Defaults: dflt (w:rFonts Calibri), thm (the theme's minorHAnsi, the
eastAsia slot minorEastAsia, w:lang eastAsia ko-KR - as docx4j-getting-started). Marks: none, sz16 (8pt), con16
(Consolas 8pt), con (Consolas 11pt), tnr40 (Times New Roman 20pt), ari (Arial 11pt), cstyle (w:rStyle of a
character style Consolas 8pt), ea16 (only w:eastAsia Consolas, 8pt), pcode (no mark properties, the paragraph
style Code: Consolas 8pt), pcode-tnr40 (style Code, the mark Times New Roman 20pt). Where: cell (in a one-cell
table). `-c14`: Word 2010's layout. nrm-<mark>: Normal Arial 8pt over document defaults of Calibri 11pt (as
docx4j-conventions), the probes' runs Calibri 11pt; the empty paragraphs `<w:p/>` (pure), `<w:p><w:pPr/></w:p>`
(ppr), only a w:jc (jc), an empty mark w:rPr (none), or a mark of 8pt (sz16). n12-*: Normal 12pt, no font (as
tests/repros/word/table-spaced-sequence-page-break.docx). `-tbl`: the empty paragraphs after a one-cell table
("Cellq") - their line (Botq's top less `-tbl0`'s, no empty paragraphs) / 10, not the table above.
`--table <pdf> ...` writes <doc> TAB <empty line height> from Word's PDFs (pdftotext's word boxes).
No dependencies: the package is written as XML.
"""
import os, re, subprocess, sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">'
         '<a:themeElements><a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
         '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
         '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
         '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
         '<a:fontScheme name="F"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
         '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
         '<a:fmtScheme name="M"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
         '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
         '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
         '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
         '</a:themeElements></a:theme>')
CON = 'w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"'
TNR = 'w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"'
MARKS = {"none": "", "sz16": '<w:sz w:val="16"/>', "con16": f'<w:rFonts {CON}/><w:sz w:val="16"/>',
         "con": f'<w:rFonts {CON}/>', "tnr40": f'<w:rFonts {TNR}/><w:sz w:val="40"/>',
         "ari": '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>', "cstyle": '<w:rStyle w:val="CodeChar"/>',
         "ea16": '<w:rFonts w:eastAsia="Consolas"/><w:sz w:val="16"/>', "pcode": "",
         "pcode-tnr40": f'<w:rFonts {TNR}/><w:sz w:val="40"/>'}
# nrm: the document defaults Calibri 11pt, Normal Arial 8pt (as docx4j-conventions); the probes' runs Calibri 11pt.
NORMAL_ARIAL = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
DEFAULTS = {"dflt": ('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>', ""),
            "thm": ('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>',
                    '<w:lang w:val="en-AU" w:eastAsia="ko-KR" w:bidi="ar-SA"/>')}


def write(out, name, defaults, mark, cell=False, mode=15, table_before=False, count=10):
    nrm = defaults in ("nrm", "n12")
    fonts, lang = DEFAULTS["dflt" if nrm else defaults]
    normal = {"nrm": NORMAL_ARIAL, "n12": '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'}.get(defaults, "")
    pstyle = '<w:pStyle w:val="Code"/>' if mark.startswith("pcode") else ""
    run = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr>' if nrm else ''
    text = lambda t: f'<w:p><w:r>{run}<w:t>{t}</w:t></w:r></w:p>'
    empty = {"pure": '<w:p/>', "ppr": '<w:p><w:pPr/></w:p>', "jc": '<w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p>'}.get(
        mark, f'<w:p><w:pPr>{pstyle}<w:rPr>{MARKS.get(mark, "")}</w:rPr></w:pPr></w:p>') * count
    if cell:
        empty = ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                 f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{empty}</w:tc></w:tr></w:tbl>')
    if table_before:
        empty = ('<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                 f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{text("Cellq")}</w:tc></w:tr></w:tbl>') + empty
    body = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{text("Topq")}{empty}{text("Botq")}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>{fonts}<w:sz w:val="22"/><w:szCs w:val="22"/>{lang}</w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>{normal}</w:style>'
              '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
              f'<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts {CON}/><w:sz w:val="16"/></w:rPr></w:style>'
              f'<w:style w:type="character" w:styleId="CodeChar"><w:name w:val="Code Char"/><w:rPr><w:rFonts {CON}/><w:sz w:val="16"/></w:rPr></w:style>'
              '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
              '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
              '</w:styles>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
                f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             f'<Relationship Id="rId1" Type="{R}styles" Target="styles.xml"/><Relationship Id="rId2" Type="{R}settings" Target="settings.xml"/>'
             f'<Relationship Id="rId3" Type="{R}theme" Target="theme/theme1.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels); z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", body); z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings)
        z.writestr("word/theme/theme1.xml", THEME)


def variants():
    for defaults in DEFAULTS:
        for mark in MARKS:
            yield f"{defaults}-{mark}", defaults, mark, False, 15
    yield "thm-con16-c14", "thm", "con16", False, 14
    yield "thm-sz16-c14", "thm", "sz16", False, 14
    yield "thm-con16-cell", "thm", "con16", True, 15
    yield "dflt-tnr40-cell", "dflt", "tnr40", True, 15
    for mark in ("pure", "ppr", "jc", "none", "sz16"):
        yield f"nrm-{mark}", "nrm", mark, False, 15
    for defaults in ("nrm", "n12"):
        for mark in ("pure", "ppr"):
            yield f"{defaults}-{mark}-tbl", defaults, mark, False, 15, True
        yield f"{defaults}-tbl0", defaults, "pure", False, 15, True, 0
    yield "n12-pure", "n12", "pure", False, 15


def table(pdfs):
    for pdf in sorted(pdfs):
        tops = {}
        rows = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        for m in re.finditer(r'yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">(Topq|Botq)</word>', rows):
            tops.setdefault(m.group(2), float(m.group(1)))
        name = os.path.basename(pdf)[:-4]
        print(f"{name}\t{(tops['Botq'] - tops['Topq'] - 13.43) / 10:.3f}" if len(tops) == 2 else f"{name}\t?")


if __name__ == "__main__":
    if sys.argv[1:2] == ["--table"]:
        table(sys.argv[2:])
    else:
        out = sys.argv[1] if len(sys.argv) > 1 else "mark-font"
        os.makedirs(out, exist_ok=True)
        for name, *rest in variants():
            write(out, name, *rest)
