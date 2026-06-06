#!/usr/bin/env python3
"""Construct a .docx with SDT content controls (checkbox, dropdown, date,
plain-text, placeholder) to validate content-control rendering.
-> tests/repros/word-sdt/"""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), "../../tests/repros/word-sdt")
os.makedirs(OUT, exist_ok=True)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
NS = f'xmlns:w="{W}" xmlns:w14="{W14}"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>'

def checkbox(checked):
    val = "1" if checked else "0"
    glyph = "&#x2612;" if checked else "&#x2610;"
    return (f'<w:sdt><w:sdtPr><w14:checkbox><w14:checked w14:val="{val}"/>'
            f'<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
            f'<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>'
            f'<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>{glyph}</w:t></w:r></w:sdtContent></w:sdt>')

def dropdown(selected):
    return (f'<w:sdt><w:sdtPr><w:dropDownList>'
            f'<w:listItem w:displayText="Red" w:value="r"/>'
            f'<w:listItem w:displayText="Green" w:value="g"/>'
            f'<w:listItem w:displayText="Blue" w:value="b"/></w:dropDownList></w:sdtPr>'
            f'<w:sdtContent><w:r><w:t>{selected}</w:t></w:r></w:sdtContent></w:sdt>')

def datectrl(text):
    return (f'<w:sdt><w:sdtPr><w:date w:fullDate="2026-06-06T00:00:00Z">'
            f'<w:dateFormat w:val="MMMM d, yyyy"/></w:date></w:sdtPr>'
            f'<w:sdtContent><w:r><w:t>{text}</w:t></w:r></w:sdtContent></w:sdt>')

def plaintext(text):
    return (f'<w:sdt><w:sdtPr><w:alias w:val="Full name"/><w:tag w:val="name"/><w:text/></w:sdtPr>'
            f'<w:sdtContent><w:r><w:t>{text}</w:t></w:r></w:sdtContent></w:sdt>')

def r(t): return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'
def b(t): return f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
def p(*inner): return '<w:p>' + ''.join(inner) + '</w:p>'

body = (
    p(b("Checkboxes:  ")) +
    p(checkbox(True), r("  Subscribe to newsletter")) +
    p(checkbox(False), r("  Send me promotions")) +
    p(checkbox(True), r("  I agree to the terms")) +
    p(b("Dropdown:  "), r("Favorite colour: "), dropdown("Green")) +
    p(b("Date picker:  "), r("Signed on "), datectrl("June 6, 2026")) +
    p(b("Plain-text control:  "), r("Name: "), plaintext("Ada Lovelace"))
)
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>{body}{SECT}</w:body></w:document>'

path = os.path.join(OUT, "sdt-controls.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", doc)
print("wrote", path)
