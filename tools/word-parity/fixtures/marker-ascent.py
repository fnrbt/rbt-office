#!/usr/bin/env python3
"""Fixture family: how much a list marker in a taller font raises its line, in Word.

    python3 tools/word-parity/fixtures/marker-ascent.py <out dir>
    node tools/word-parity/page-parity.mjs --dist dist --dir <out dir> --refresh-word   (Word's PDFs in <out dir>/word)
    python3 tools/word-parity/pdf-geometry.py <out dir>/word/*.pdf > <out dir>/word/geometry.tsv

Hypothesis (docs/layout-parity-flow.md ledger): a list paragraph's first line is its text's line (at the
paragraph's line spacing) plus how far the marker's ascent - its font's line less its descent, at the
marker's own size - reaches above the text's ascent (the same, of the line's tallest text); the marker's
descent below the text's adds nothing, nor does the spacing multiple scale the excess; an exact line is not
grown. Each document: 30 one-line list paragraphs (or plain ones: the control), no space before or after,
on a Letter page with 1in margins; what varies is in the name:
  <text>-<marker>[-<spacing>]
  text:    cal11 (Calibri 11), tnr12 (Times New Roman 12), ari10 (Arial 10), cou10 (Courier New 10),
           cam11 (Cambria 11), geo11 (Georgia 11), seg10 (Segoe UI 10), cal20 (Calibri 20),
           mix (Calibri 11 and a 16pt Calibri word)
  marker:  plain (no list), sym (Word's Symbol bullet at the text's size), sym16, sym24 (16pt, 24pt),
           wd20 (Wingdings bullet 20pt), num20 (Calibri "1." 20pt), cour24 (Courier New "o" 24pt),
           cour22 (22pt), tnr8 (Times New Roman "1." 8pt), seg14 (Segoe UI "1." 14pt)
  spacing: (none) 1.15 (auto 276), s1 (auto 240), s2 (auto 480), al15 / al30 (at least 300 / 600),
           ex20 (exactly 400)
Measured: the first line's text top (against the plain control: the marker's ascent above the text's) and
the paragraphs' pitch (the line's whole height). No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "marker-ascent"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

TEXT = {"cal11": ("Calibri", 22), "tnr12": ("Times New Roman", 24), "ari10": ("Arial", 20), "cou10": ("Courier New", 20),
        "cam11": ("Cambria", 22), "geo11": ("Georgia", 22), "seg10": ("Segoe UI", 20), "cal20": ("Calibri", 40), "mix": ("Calibri", 22)}
# marker: (numFmt, lvlText, font, half-points or None for the text's size)
MARKER = {"plain": None, "sym": ("bullet", "", "Symbol", None), "sym16": ("bullet", "", "Symbol", 32),
          "sym24": ("bullet", "", "Symbol", 48), "wd20": ("bullet", "", "Wingdings", 40),
          "num20": ("decimal", "%1.", "Calibri", 40), "cour24": ("bullet", "o", "Courier New", 48),
          "cour22": ("bullet", "o", "Courier New", 44), "tnr8": ("decimal", "%1.", "Times New Roman", 16),
          "seg14": ("decimal", "%1.", "Segoe UI", 28)}
SPACING = {"": 'w:line="276" w:lineRule="auto"', "s1": 'w:line="240" w:lineRule="auto"', "s2": 'w:line="480" w:lineRule="auto"',
           "al15": 'w:line="300" w:lineRule="atLeast"', "al30": 'w:line="600" w:lineRule="atLeast"', "ex20": 'w:line="400" w:lineRule="exact"'}

def numbering(marker):
    fmt, text, font, size = marker
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/>' + (f'<w:sz w:val="{size}"/>' if size else '') + '</w:rPr>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}>'
            f'<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="720"/></w:pPr>{rpr}</w:lvl></w:abstractNum>'
            f'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')

def document(text, listed, spacing):
    font, size = TEXT[text]
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{size}"/></w:rPr>'
    paras = []
    for i in range(1, 31):
        num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if listed else ''
        runs = f'<w:r>{rpr}<w:t xml:space="preserve">Line {i} of the list </w:t></w:r>'
        if text == "mix":
            runs += f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="32"/></w:rPr><w:t>Big</w:t></w:r>'
        paras.append(f'<w:p><w:pPr>{num}<w:spacing w:before="0" w:after="0" {SPACING[spacing]}/>{rpr}</w:pPr>{runs}</w:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(paras)}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')

def write(name, doc, numbering_xml=None):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          + ('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' if numbering_xml else '')
          + '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             + ('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' if numbering_xml else '')
             + '</Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        if numbering_xml: z.writestr("word/numbering.xml", numbering_xml)

CASES = []
for text in ["cal11", "tnr12", "ari10", "cou10", "cam11", "geo11", "seg10"]:
    for marker in ["plain", "sym", "sym16", "num20", "cour24"]:
        CASES.append((text, marker, ""))
for marker in ["wd20", "tnr8", "seg14", "sym24", "cour22"]:
    CASES.append(("cal11", marker, ""))
for text in ["cal20"]:
    for marker in ["plain", "cour22", "sym24", "num20"]:
        CASES.append((text, marker, ""))
for text in ["mix"]:
    for marker in ["plain", "sym16", "sym24"]:
        CASES.append((text, marker, ""))
for spacing in ["s1", "s2", "al15", "al30", "ex20"]:
    for marker in ["plain", "sym16", "num20", "cour24"]:
        CASES.append(("cal11", marker, spacing))
for text, marker, spacing in CASES:
    name = f"{text}-{marker}" + (f"-{spacing}" if spacing else "")
    m = MARKER[marker]
    write(name, document(text, m is not None, spacing), numbering(m) if m else None)
print(out, len(CASES))
