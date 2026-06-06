#!/usr/bin/env python3
"""Construct .docx fixtures for features docx-editor (eigenpal) renders that we
don't, to document the gaps. Writes to tests/repros/word-gaps/."""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), "../../tests/repros/word-gaps")
os.makedirs(OUT, exist_ok=True)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = (f'xmlns:w="{W}" xmlns:m="{M}" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:o="urn:schemas-microsoft-com:office:office"')

CT_BASE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
{extra}</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="''' + W + '''"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>'''
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>{hdr}</w:sectPr>'

def doc(body, sect_hdr=''):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>{body}'
            + SECT.format(hdr=sect_hdr) + '</w:body></w:document>')
def heading(t): return f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
def p(inner): return f'<w:p>{inner}</w:p>'
def run(text, rpr=''): return f'<w:r>{("<w:rPr>"+rpr+"</w:rPr>") if rpr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'

def write(name, body, extra_parts=None, extra_ct=''):
    path = os.path.join(OUT, name + ".docx")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT_BASE.format(extra=extra_ct))
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", body)
        for pn, content in (extra_parts or {}).items():
            z.writestr(pn, content)
    print("wrote", path)

# 1. Hidden text (w:vanish) — Word/docx-editor HIDE it; we render it (correctness bug)
write("hidden-text", doc(
    heading("Hidden text (w:vanish) repro") +
    p(run("Visible before. ") + run("THIS IS HIDDEN (w:vanish) and must not render. ", "<w:vanish/>") + run("Visible after.")) +
    p(run("The line above should read “Visible before.  Visible after.” with no hidden text."))))

# 2. Run effects: double-strike, position (raise/lower), emphasis mark, symbol char, kerning
write("run-effects", doc(
    heading("Run effects repro") +
    p(run("Double strikethrough: ") + run("struck twice", "<w:dstrike/>")) +
    p(run("Raised") + run("super", "<w:position w:val=\"12\"/>") + run(" and ") + run("lowered", "<w:position w:val=\"-8\"/>") + run(" baseline (w:position)")) +
    p(run("Emphasis marks (w:em dot): ") + run("emphasized", "<w:em w:val=\"dot\"/>")) +
    p('<w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:t xml:space="preserve">Symbol char (w:sym): </w:t></w:r><w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r>') +
    p(run("Hidden via vanish should be gone: ") + run("[gone]", "<w:vanish/>"))))

# 3. Content-control types: checkbox + dropdown (docx-editor renders the control; we render generic content)
checkbox = ('<w:sdt><w:sdtPr><w:alias w:val="Agree"/><w:tag w:val="agree"/>'
            '<w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:font="MS Gothic" w14:val="2612"/>'
            '<w14:uncheckedState w14:font="MS Gothic" w14:val="2610"/></w14:checkbox></w:sdtPr>'
            '<w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>')
dropdown = ('<w:sdt><w:sdtPr><w:alias w:val="Color"/><w:tag w:val="color"/>'
            '<w:dropDownList><w:listItem w:displayText="Red" w:value="r"/><w:listItem w:displayText="Green" w:value="g"/></w:dropDownList>'
            '</w:sdtPr><w:sdtContent><w:r><w:t>Green</w:t></w:r></w:sdtContent></w:sdt>')
write("sdt-controls", doc(
    heading("Content controls repro") +
    p(run("Checkbox: ") + checkbox + run("  I agree")) +
    p(run("Dropdown: ") + dropdown)))

# 4. Math (OMML) — docx-editor renders the equation; we show "[Equation]"
frac = ('<m:oMathPara><m:oMath>'
        '<m:r><m:t>x</m:t></m:r><m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'
        '<m:r><m:t> + </m:t></m:r>'
        '<m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>'
        '<m:r><m:t> = </m:t></m:r>'
        '<m:rad><m:deg><m:r><m:t>2</m:t></m:r></m:deg><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad>'
        '</m:oMath></m:oMathPara>')
write("math-omml", doc(
    heading("Math (OMML) repro") +
    p(run("An inline equation follows; it should render as math, not “[Equation]”:")) +
    f'<w:p>{frac}</w:p>'))

# 5. Watermark (VML text watermark in the header) — docx-editor renders it; we render nothing
hdr_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
header = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<w:hdr {NS}><w:p><w:r><w:pict>'
          '<v:shape id="wm1" o:spid="_x0000_s2049" type="#_x0000_t136" '
          'style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;rotation:-45;z-index:-1" '
          'fillcolor="silver" stroked="f"><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" '
          'string="CONFIDENTIAL"/></v:shape></w:pict></w:r></w:p></w:hdr>')
doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            '</Relationships>')
wm_body = doc(
    heading("Watermark repro") +
    p(run("This page should show a diagonal grey “CONFIDENTIAL” watermark behind the text.")) +
    "".join(p(run("Body line %d, lorem ipsum dolor sit amet consectetur adipiscing elit." % i)) for i in range(1, 12)),
    sect_hdr='<w:headerReference w:type="default" r:id="rIdH"/>')
write("watermark", wm_body, extra_parts={
    "word/_rels/document.xml.rels": doc_rels,
    "word/header1.xml": header,
    "word/_rels/header1.xml.rels": hdr_rels,
}, extra_ct='<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
