#!/usr/bin/env python3
"""Fixture family: does a list marker's font make a line taller in Word?

    python3 tools/word-parity/fixtures/list-marker-height.py <out dir>

Hypothesis (docs/layout-parity-flow.md ledger): a line is as tall as its text; the list marker's font
(Symbol for Word's bullet) does not raise it. Each document: 30 one-line paragraphs, Calibri 11pt, line
spacing 1.15 (auto 276), no space before or after - only what is named varies:
  plain            no list (the baseline)
  bullet-symbol    Word's bullet: U+F0B7 in Symbol
  bullet-calibri   U+2022 in Calibri
  bullet-symbol16  the Symbol bullet at 16pt (does the marker's size count?)
  numbered         "1." decimal in Calibri
  text-cambria16   no list, the text in Cambria 16pt (control: the text's font does count)
  bullet-symbol-3lines, bullet-symbol16-3lines   10 list paragraphs of 3 lines each (line breaks):
                   does the marker's taller first line push the paragraph's other lines down?
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "list-marker-height"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def numbering(fmt, text, font, size):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/>' + (f'<w:sz w:val="{size}"/>' if size else '') + '</w:rPr>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}>'
            f'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>{rpr}</w:lvl></w:abstractNum>'
            f'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')

def document(listed, font="Calibri", size=22, count=30, lines=1):
    paras = []
    for i in range(1, count + 1):
        num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if listed else ''
        paras.append(f'<w:p><w:pPr>{num}<w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>'
                     f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{size}"/></w:rPr><w:t>Line {i} of the list</w:t></w:r>'
                     + ''.join(f'<w:r><w:br/><w:t>Line {i} part {k}</w:t></w:r>' for k in range(2, lines + 1)) + '</w:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(paras)}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')

def write(name, doc, numbering_xml=None):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          + ('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' if numbering_xml else '')
          + '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             + ('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' if numbering_xml else '')
             + '</Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc); z.writestr("word/styles.xml", STYLES)
        if numbering_xml: z.writestr("word/numbering.xml", numbering_xml)

write("plain", document(False))
write("bullet-symbol", document(True), numbering("bullet", "", "Symbol", None))
write("bullet-calibri", document(True), numbering("bullet", "•", "Calibri", None))
write("bullet-symbol16", document(True), numbering("bullet", "", "Symbol", 32))
write("numbered", document(True), numbering("decimal", "%1.", "Calibri", None))
write("text-cambria16", document(False, "Cambria", 32))
write("bullet-symbol-3lines", document(True, count=10, lines=3), numbering("bullet", "", "Symbol", None))
write("bullet-symbol16-3lines", document(True, count=10, lines=3), numbering("bullet", "", "Symbol", 32))
print(out)
