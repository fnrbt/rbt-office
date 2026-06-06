#!/usr/bin/env python3
"""Construct a .docx exercising field codes: STYLEREF (nearest heading text),
DATE, TIME. Validates the field-computation additions. -> tests/repros/word-fields/"""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), "../../tests/repros/word-fields")
os.makedirs(OUT, exist_ok=True)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = (f'xmlns:w="{W}" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          f'<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
          '<w:pPr><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr>'
          '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>')
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'

def heading(t):
    return (f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>')

def field(instr, placeholder):
    # complex field: begin / instrText / separate / result run / end
    return (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve">{instr}</w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:t xml:space="preserve">{placeholder}</w:t></w:r>'
            f'<w:r><w:fldChar w:fldCharType="end"/></w:r>')

def para(inner):
    return f'<w:p>{inner}</w:p>'

def run(t):
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'

body = (
    heading("Chapter One") +
    para(run("Current section: ") + field(' STYLEREF "Heading 1" ', "X")) +
    para(run("Body text under chapter one.")) +
    heading("Chapter Two") +
    para(run("Current section: ") + field(' STYLEREF "Heading 1" ', "X")) +
    para(run("Printed on ") + field(' DATE \\@ "MMMM d, yyyy" ', "date") +
         run(" at ") + field(' TIME \\@ "h:mm tt" ', "time") + run(".")))

doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       f'<w:document {NS}><w:body>{body}{SECT}</w:body></w:document>')

path = os.path.join(OUT, "fields-styleref-date-time.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/styles.xml", STYLES)
    z.writestr("word/document.xml", doc)
print("wrote", path)
print("expected: STYLEREF#1 -> 'Chapter One', STYLEREF#2 -> 'Chapter Two', DATE/TIME computed")
