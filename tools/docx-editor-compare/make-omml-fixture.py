#!/usr/bin/env python3
"""Construct a .docx with OMML (m:oMath) equations to validate math rendering.
-> tests/repros/word-math/"""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), "../../tests/repros/word-math")
os.makedirs(OUT, exist_ok=True)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = f'xmlns:w="{W}" xmlns:m="{M}"'
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

def r(t): return f'<m:r><m:t>{t}</m:t></m:r>'
def frac(num, den): return f'<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>'
def ssup(e, sup): return f'<m:sSup><m:e>{e}</m:e><m:sup>{sup}</m:sup></m:sSup>'
def ssub(e, sub): return f'<m:sSub><m:e>{e}</m:e><m:sub>{sub}</m:sub></m:sSub>'
def rad(e, deg=None):
    pr = '' if deg else '<m:radPr><m:degHide m:val="1"/></m:radPr>'
    d = f'<m:deg>{deg}</m:deg>' if deg else '<m:deg/>'
    return f'<m:rad>{pr}{d}<m:e>{e}</m:e></m:rad>'
def nary(chr_, sub, sup, e):
    return (f'<m:nary><m:naryPr><m:chr m:val="{chr_}"/></m:naryPr>'
            f'<m:sub>{sub}</m:sub><m:sup>{sup}</m:sup><m:e>{e}</m:e></m:nary>')
def delim(parts, beg='(', end=')'):
    es = ''.join(f'<m:e>{p}</m:e>' for p in parts)
    return f'<m:d><m:dPr><m:begChr m:val="{beg}"/><m:endChr m:val="{end}"/></m:dPr>{es}</m:d>'
def func(name, arg): return f'<m:func><m:fName>{r(name)}</m:fName><m:e>{arg}</m:e></m:func>'

equations = [
    ("fraction", frac(r("a+b"), r("c"))),
    ("superscript", ssup(r("x"), r("2")) + r(" + ") + ssup(r("y"), r("n"))),
    ("subscript", ssub(r("x"), r("i")) + r(" + ") + ssub(r("a"), r("k"))),
    ("radical", rad(r("x"), None) + r(" and ") + rad(ssup(r("b"), r("2")) + r("-4ac"))),
    ("summation", nary("∑", r("i=1"), r("n"), ssup(r("i"), r("2")))),
    ("integral", nary("∫", r("0"), r("1"), func("f", r("x")) + r("dx"))),
    ("delimiters", delim([frac(r("n"), r("k"))])),
    ("quadratic", r("x = ") + frac(r("-b ± ") + rad(ssup(r("b"), r("2")) + r("-4ac")), r("2a"))),
]

def para(label, omml):
    return (f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{label}:  </w:t></w:r>'
            f'<m:oMath>{omml}</m:oMath></w:p>')

body = ''.join(para(lbl, omml) for lbl, omml in equations)
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>{body}{SECT}</w:body></w:document>'

path = os.path.join(OUT, "omml-equations.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", doc)
print("wrote", path, "-", len(equations), "equations")
