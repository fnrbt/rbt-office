#!/usr/bin/env python3
"""Fixture family: how high an empty paragraph is - its mark's size from the paragraph style (Normal) or
from the document defaults.

    python3 tools/word-parity/fixtures/empty-mark.py tests/word/fixtures/empty-mark
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/empty-mark --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/empty-mark/word/*.pdf > tests/word/fixtures/empty-mark/word/geometry.tsv

docx4j-conventions (Normal: Arial 8pt, the defaults 11pt, 10pt after at 1.15) has bare `<w:p/>` paragraphs
between its code samples: Word gives each 20.58pt (an 8pt Arial line at 1.15 and 10pt after), ours 24.54
(an 11pt line: an earlier fix sized a "pure empty" paragraph from the document defaults). Each document:
1in margins, the defaults Calibri 11pt (w:sz 22), single-spaced, nothing before or after unless said;
"A 1" .. "A 3", then five empty paragraphs of the form, then "P-probe" (its top says how high they are):
  <form>-n<N>[-<ctx>][-m14]
  form   bare (<w:p/>), ppr (<w:p><w:pPr/></w:p>), rpr (an empty w:pPr/w:rPr), jc (w:pPr with w:jc left)
  N      Normal's w:sz (16, 24; none: no size in Normal)
  ctx    tbl: after a one-cell table instead of "A 3"
  conv   docx4j-conventions' styles: Normal Arial 8pt, the defaults 10pt after at 1.15 (line 276)
in compatibility mode 15 and 14 (-m14).
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "empty-mark"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')
FORMS = {
    "bare": '<w:p/>',
    "ppr": '<w:p><w:pPr/></w:p>',
    "rpr": '<w:p><w:pPr><w:rPr/></w:pPr></w:p>',
    "jc": '<w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p>',
}

def para(text):
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'

TABLE = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
         '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
         '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>T 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')

def document(form, ctx):
    body = [para("A 1"), para("A 2"), TABLE if ctx == "tbl" else para("A 3")]
    body += [FORMS[form]] * 5
    body += [para("P-probe"), para("C 1")]
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{SECT}</w:body></w:document>'

def styles(normal_sz, conv):
    if conv:
        defaults = ('<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
                    '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>')
        normal = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
    else:
        defaults = ('<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
                    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>')
        normal = f'<w:rPr><w:sz w:val="{normal_sz}"/></w:rPr>' if normal_sz else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>{defaults}</w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>{normal}</w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
            '</w:styles>')

def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')

def write(name, doc, sty, mode):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", sty); z.writestr("word/settings.xml", settings(mode))

count = 0
for mode in [15, 14]:
    suffix = "-m14" if mode == 14 else ""
    for form in FORMS:
        for n in [16, 24, None]:
            write(f"{form}-n{n or 'none'}{suffix}", document(form, None), styles(n, False), mode); count += 1
    for form in ["bare", "ppr"]:
        for n in [16, 24]:
            write(f"{form}-n{n}-tbl{suffix}", document(form, "tbl"), styles(n, False), mode); count += 1
    write(f"bare-conv{suffix}", document("bare", None), styles(None, True), mode); count += 1
print(out, count)
