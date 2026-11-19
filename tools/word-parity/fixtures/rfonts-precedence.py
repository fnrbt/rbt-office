#!/usr/bin/env python3
"""Fixture family: which w:rFonts wins - an explicit face lower in the style chain, or a theme face above it?

    python3 tools/word-parity/fixtures/rfonts-precedence.py <out dir>

Hypothesis (docs/layout-parity-flow.md ledger): each level of the chain (document defaults, paragraph style,
character style, run) replaces the inherited face slot by slot; an explicit face at a lower level beats a
theme face inherited from above; only on one and the same element does the theme attribute beat the
explicit one. The theme's minor font is Calibri, its major Cambria; the explicit face is Georgia.
Each document is one paragraph "Specimen"; Word's font of it (Range.Font.Name) is compared with ours.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "rfonts-precedence"
os.makedirs(out, exist_ok=True)
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
THEME_MINOR = 'w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"'
GEORGIA = 'w:ascii="Georgia" w:hAnsi="Georgia"'

def doc(name, defaults="", normal="", charstyle=None, run="", expect=""):
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
              + (f'<w:rFonts {defaults}/>' if defaults else '') + '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>'
              + (f'<w:rFonts {normal}/>' if normal else '') + '</w:rPr></w:style>'
              + (f'<w:style w:type="character" w:styleId="Spec"><w:name w:val="Spec"/><w:rPr><w:rFonts {charstyle}/></w:rPr></w:style>' if charstyle else '')
              + '</w:styles>')
    rpr = ('<w:rStyle w:val="Spec"/>' if charstyle else '') + (f'<w:rFonts {run}/>' if run else '')
    body = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:rPr>{rpr}</w:rPr><w:t>Specimen</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels); z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", body); z.writestr("word/styles.xml", styles); z.writestr("word/theme/theme1.xml", THEME)
    with open(os.path.join(out, "expected.tsv"), "a") as f:
        f.write(f"{name}\t{expect}\n")

open(os.path.join(out, "expected.tsv"), "w").close()
# (The hypothesis's predictions; Word's answers are what count.)
doc("defaults-theme_normal-explicit", defaults=THEME_MINOR, normal=GEORGIA, expect="Georgia")
doc("defaults-explicit_normal-theme", defaults=GEORGIA, normal=THEME_MINOR, expect="Calibri")
doc("normal-theme_run-explicit", normal=THEME_MINOR, run=GEORGIA, expect="Georgia")
doc("defaults-theme_charstyle-explicit", defaults=THEME_MINOR, charstyle=GEORGIA, expect="Georgia")
doc("normal-both", normal=THEME_MINOR + " " + GEORGIA, expect="Calibri")
doc("run-both", run=THEME_MINOR + " " + GEORGIA, expect="Calibri")
doc("defaults-theme_normal-ascii-only", defaults=THEME_MINOR, normal='w:ascii="Georgia"', expect="Georgia")
doc("defaults-theme_normal-hansi-only", defaults=THEME_MINOR, normal='w:hAnsi="Georgia"', expect="Calibri")
doc("defaults-major_run-explicit", defaults='w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"', run=GEORGIA, expect="Georgia")
doc("defaults-theme_only", defaults=THEME_MINOR, expect="Calibri")
print(out)
