#!/usr/bin/env python3
"""Construct minimal .docx fixtures that exercise features our renderer is
known (from code) to handle incompletely, to demonstrate the shortfalls.
Writes to tests/repros/word-audit/. Calibri 11pt defaults so fonts match."""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), "../../tests/repros/word-audit")
os.makedirs(OUT, exist_ok=True)

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>'''

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'

def doc(body):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {W}><w:body>{body}{SECT}</w:body></w:document>'

def heading(t):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'

def para(inner):
    return f'<w:p>{inner}</w:p>'

def write(name, body):
    path = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc(body))
    print("wrote", path)

# ---- 1. Underline styles (we collapse w:u@val to a bool -> all render single) ----
ulines = []
for val, label in [("single","single"),("double","double"),("thick","thick"),
                   ("dotted","dotted"),("dash","dash"),("dotDash","dot-dash"),
                   ("wave","wavy"),("wavyDouble","wavy-double")]:
    ulines.append(para(f'<w:r><w:rPr><w:u w:val="{val}"/></w:rPr><w:t xml:space="preserve">This text uses {label} underline (w:u val={val})</w:t></w:r>'))
# colored underline
ulines.append(para('<w:r><w:rPr><w:u w:val="single" w:color="FF0000"/></w:rPr><w:t xml:space="preserve">Single underline with a red underline color (w:u color=FF0000)</w:t></w:r>'))
write("underline-styles", heading("Underline styles repro") + "".join(ulines))

# ---- 2. Decimal + center tab stops (we only read val=right tabs) ----
rows = []
tabs = '<w:tabs><w:tab w:val="decimal" w:pos="4320"/></w:tabs>'
for n in ["1.5", "22.75", "333.1", "4444.625", "7.0"]:
    rows.append(para(f'<w:pPr>{tabs}</w:pPr><w:r><w:t xml:space="preserve">Item</w:t></w:r><w:r><w:tab/><w:t xml:space="preserve">{n}</w:t></w:r>'))
center = '<w:tabs><w:tab w:val="center" w:pos="4320"/></w:tabs>'
crow = para(f'<w:pPr>{center}</w:pPr><w:r><w:tab/><w:t xml:space="preserve">centered on the tab stop</w:t></w:r>')
# left tab at an explicit position
left = '<w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs>'
lrow = para(f'<w:pPr>{left}</w:pPr><w:r><w:t xml:space="preserve">Name:</w:t></w:r><w:r><w:tab/><w:t xml:space="preserve">value at 2 inch left tab</w:t></w:r>')
write("tab-stops", heading("Decimal / center / left tab stops repro")
      + para('<w:r><w:t>Numbers below should align on the decimal point at 3":</w:t></w:r>')
      + "".join(rows) + crow + lrow)

# ---- 3. RTL / bidi (no rtl/bidi handling in our code) ----
heb = "שלום עולם זהו טקסט בעברית"          # "Hello world this is Hebrew text"
ar  = "مرحبا بالعالم هذا نص عربي"            # Arabic
rtl_rpr = '<w:rPr><w:rtl/></w:rPr>'
body = (heading("RTL / bidi repro")
        + para(f'<w:pPr><w:bidi/></w:pPr><w:r>{rtl_rpr}<w:t xml:space="preserve">{heb}</w:t></w:r>')
        + para(f'<w:pPr><w:bidi/></w:pPr><w:r>{rtl_rpr}<w:t xml:space="preserve">{ar}</w:t></w:r>')
        + para(f'<w:pPr><w:bidi/></w:pPr><w:r>{rtl_rpr}<w:t xml:space="preserve">123 ميكس mixed عربي 456</w:t></w:r>'))
write("rtl-bidi", body)
