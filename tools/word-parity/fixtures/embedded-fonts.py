#!/usr/bin/env python3
"""Fixtures for when and how Word sets text in a font the document embeds (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/embedded-fonts.py [out dir]   (default tests/word/fixtures/embedded-fonts)

needs fontTools (pip install fonttools; it renames and subsets the faces), then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/embedded-fonts \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/embedded-fonts --refresh-word
    python3 tools/word-parity/fixtures/embedded-fonts.py --widths tests/word/fixtures/embedded-fonts

--widths reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/widths.tsv:
fixture, paragraph (what it is set in), word, the word's width in Word's PDF (pdftotext's box; of a
glyphs-* fixture: the paragraph's runs joined by "|" and the line's width), and
the font of Word's PDF that word is in (pdftotext -bbox does not say: the fonts pdffonts lists, by the
widths), plus <dir>/word/fonts.tsv: fixture, the fonts Word's PDF embeds (pdffonts' name, without
the subset prefix). tests/word/embedded-fonts.fsx checks ours against them. page-lines.vbs opens
each document in a Word of its own: in one session for all of them Word can lose a document's
embedded faces (the batch run had subset's missing characters drawn as .notdef boxes).

The embedded faces are Liberation Mono 2.1.5 (fonts/, SIL OFL 1.1; sans-reg: Carlito) renamed - family, full and
PostScript names - so the machine cannot have them, subset to printable ASCII (smaller packages; the
glyphs kept are the font's own), and obfuscated as ECMA-376 Part 2 / Word do: the first 32 bytes
XORed with the 16 bytes of w:fontKey's GUID, last byte first. Every character of Liberation Mono is
1229/2048 em (6.0pt at 10pt), Courier New's too: a word in the embedded face is 6.0pt a character,
in Word's fallback for a family it has not got (no PANOSE given) another width. Every paragraph is
10pt, single, nothing before or after; the words "MMMMMMMMMM iiiiiiiiii Quiz" (60pt each in the
embedded face). What varies:

  reg                 "Rbt Embed Mono": embedRegular only; w:embedTrueTypeFonts on. Paragraphs in its
                      regular, bold, italic and bold italic (w:b, w:i)
  reg-noset           the same without w:embedTrueTypeFonts in the settings
  all4                all four faces embedded (embedRegular, embedBold, embedItalic, embedBoldItalic)
  bold-only           only embedBold
  calibri             all four faces embedded under the name "Calibri" (the machine has Calibri);
                      paragraphs in Calibri
  names               the font table's w:name "Rbt Outer", the face's own family "Rbt Inner";
                      paragraphs in Rbt Outer and Rbt Inner
  subset              embedRegular w:subsetted="1", w:saveSubsetFonts on, the face without "Q", "z"
                      and "e" in its cmap (as a subset of the characters used would be)
  restricted          as reg, the face's OS/2 fsType 2 (restricted licence embedding)
  noprint             as reg, fsType 4 (preview and print)
  editable            as reg, fsType 8 (editable)
  sans-reg            Carlito 2023 (fonts/) as "Rbt Embed Sans", embedRegular only: whether a bold or
                      italic made of the regular face (synthetic) has the regular face's advances
  subset-default-georgia, subset-theme-georgia, subset-roman (the table's PANOSE and family Times New
  Roman's), subset-alt-courier (w:altName Courier New), subset-all4: as subset - which font sets the
                      characters the subset lacks
  glyphs-R, -B, -RB, -RI, -all4   the subset face(s) of those styles embedded; paragraphs of runs
                      "Quiz", "uiQ", "uuQuu", "Q uu", "uu", "MQM", "uuzuu", "uu" + "Quu" (two runs), and
                      bold "Quiz", "uuQuu": which characters of a run go to the other font
  restricted-theme-georgia, restricted-roman   as restricted with a theme of Georgia / the table's
                      PANOSE of Times New Roman: the font Word sets a family it cannot use in
"""
import io
import os
import re
import subprocess
import sys
import uuid
import zipfile

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
WORDS = 'MMMMMMMMMM iiiiiiiiii Quiz'
FACES = ['Regular', 'Bold', 'Italic', 'BoldItalic']
STYLE_OF = {'Regular': '', 'Bold': '<w:b/>', 'Italic': '<w:i/>', 'BoldItalic': '<w:b/><w:i/>'}

STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="DEFAULT" w:hAnsi="DEFAULT" w:cs="DEFAULT"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''
TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>'''
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
FONT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'


def face_bytes(repo, face, family, fs_type=0, drop='', source='LiberationMono'):
    """`source`'s face (fonts/) renamed to `family`, subset to printable ASCII less `drop`."""
    from fontTools.ttLib import TTFont
    from fontTools import subset
    font = TTFont(os.path.join(repo, 'fonts', f'{source}-{face}.ttf'))
    keep = [c for c in range(0x20, 0x7f) if chr(c) not in drop]
    opts = subset.Options()
    opts.name_IDs = ['*']
    opts.notdef_outline = True
    opts.layout_features = ['*']
    sub = subset.Subsetter(opts)
    sub.populate(unicodes=keep)
    sub.subset(font)
    sub_name = {'Regular': 'Regular', 'Bold': 'Bold', 'Italic': 'Italic', 'BoldItalic': 'Bold Italic'}[face]
    name = font['name']
    for rec in list(name.names):
        if rec.nameID in (16, 17, 21, 22):
            name.removeNames(nameID=rec.nameID)
    for pid, eid, lid in ((3, 1, 0x409), (1, 0, 0)):
        name.setName(family, 1, pid, eid, lid)
        name.setName(sub_name, 2, pid, eid, lid)
        name.setName(f'{family} {sub_name}', 4, pid, eid, lid)
        name.setName(f'{family}:{sub_name}', 3, pid, eid, lid)
        name.setName(f'{family.replace(" ", "")}-{face}', 6, pid, eid, lid)
    font['OS/2'].fsType = fs_type
    out = io.BytesIO()
    font.save(out)
    return out.getvalue()


def obfuscate(data, key):
    """ECMA-376 Part 2 font obfuscation (its own inverse): the first 32 bytes XOR the GUID's 16 bytes, last first."""
    k = bytes.fromhex(key.replace('-', ''))[::-1]
    b = bytearray(data)
    for i in range(32):
        b[i] ^= k[i % 16]
    return bytes(b)


def para(family, face, runs=None):
    fonts = f'<w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:cs="{family}"/>'
    return '<w:p>' + ''.join(f'<w:r><w:rPr>{fonts}{STYLE_OF[face]}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
                             for text in (runs or [WORDS])) + '</w:p>'


MONO_TABLE = '<w:charset w:val="00"/><w:family w:val="modern"/><w:pitch w:val="fixed"/>'


def theme(minor, major):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">'
            '<a:themeElements><a:clrScheme name="C"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
            '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
            '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
            '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
            f'<a:fontScheme name="F"><a:majorFont><a:latin typeface="{major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            f'<a:minorFont><a:latin typeface="{minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '<a:fmtScheme name="M"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
            '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
            '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
            '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
            '</a:themeElements></a:theme>')


def write(path, repo, table_name, inner, faces, paras, embed_setting=True, subset_setting=False,
          fs_type=0, drop='', source='LiberationMono', default_font='Times New Roman', theme_fonts=None,
          table_props=MONO_TABLE):
    """`faces`: the faces embedded; `paras`: (family, face) per paragraph."""
    entries, rels, parts = [], [], []
    for i, face in enumerate(faces, 1):
        key = str(uuid.uuid5(uuid.NAMESPACE_URL, f'{os.path.basename(path)}/{face}')).upper()
        subsetted = ' w:subsetted="1"' if drop else ''
        entries.append(f'<w:embed{face} r:id="rId{i}" w:fontKey="{{{key}}}"{subsetted}/>')
        rels.append(f'<Relationship Id="rId{i}" Type="{FONT_REL}" Target="fonts/font{i}.odttf"/>')
        parts.append((f'word/fonts/font{i}.odttf', obfuscate(face_bytes(repo, face, inner, fs_type, drop, source), key)))
    table = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>'
             '<w:font w:name="Times New Roman"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
             f'<w:font w:name="{table_name}">{table_props}'
             f'{"".join(entries)}</w:font></w:fonts>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
                + ('<w:embedTrueTypeFonts/>' if embed_setting else '')
                + ('<w:saveSubsetFonts/>' if subset_setting else '')
                + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
                '</w:compat></w:settings>')
    body = ''.join(para(*p) for p in paras)
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        types, doc_rels = TYPES, DOC_RELS
        if theme_fonts:
            types = types.replace('</Types>', '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>')
            doc_rels = doc_rels.replace('</Relationships>', '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>')
            z.writestr('word/theme/theme1.xml', theme(*theme_fonts))
        z.writestr('[Content_Types].xml', types)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES.replace('DEFAULT', default_font))
        z.writestr('word/settings.xml', settings)
        z.writestr('word/fontTable.xml', table)
        z.writestr('word/_rels/fontTable.xml.rels',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + ''.join(rels) + '</Relationships>')
        for name, data in parts:
            z.writestr(name, data)


MONO = 'Rbt Embed Mono'
SANS = 'Rbt Embed Sans'
# The glyphs-* fixtures: which characters of a run Word sets in another font when the subset face lacks
# Q and z (one paragraph per run list; "uu|Quu" two runs).
GLYPH_RUNS = [['Quiz'], ['uiQ'], ['uuQuu'], ['Q uu'], ['uu'], ['MQM'], ['uuzuu'], ['uu', 'Quu']]
GLYPH_PARAS = [(MONO, 'Regular', runs) for runs in GLYPH_RUNS] + [(MONO, 'Bold', ['Quiz']), (MONO, 'Bold', ['uuQuu'])]


def four(family):
    return [(family, f) for f in FACES]


FIXTURES = {
    'reg': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO)),
    'reg-noset': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), embed_setting=False),
    'all4': dict(table_name=MONO, inner=MONO, faces=FACES, paras=four(MONO)),
    'bold-only': dict(table_name=MONO, inner=MONO, faces=['Bold'], paras=four(MONO)),
    'calibri': dict(table_name='Calibri', inner='Calibri', faces=FACES, paras=four('Calibri')),
    'names': dict(table_name='Rbt Outer', inner='Rbt Inner', faces=['Regular'],
                  paras=[('Rbt Outer', 'Regular'), ('Rbt Inner', 'Regular')]),
    'subset': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), subset_setting=True,
                   drop='Qze'),
    'subset-default-georgia': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), subset_setting=True,
                                   drop='Qze', default_font='Georgia'),
    'subset-theme-georgia': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), subset_setting=True,
                                 drop='Qze', theme_fonts=('Georgia', 'Verdana')),
    'subset-roman': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), subset_setting=True, drop='Qze',
                         table_props='<w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/>'),
    'subset-alt-courier': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), subset_setting=True, drop='Qze',
                               table_props='<w:altName w:val="Courier New"/>' + MONO_TABLE),
    'subset-all4': dict(table_name=MONO, inner=MONO, faces=FACES, paras=four(MONO), subset_setting=True, drop='Qze'),
    'glyphs-R': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=GLYPH_PARAS, subset_setting=True, drop='Qze'),
    'glyphs-B': dict(table_name=MONO, inner=MONO, faces=['Bold'], paras=GLYPH_PARAS, subset_setting=True, drop='Qze'),
    'glyphs-RB': dict(table_name=MONO, inner=MONO, faces=['Regular', 'Bold'], paras=GLYPH_PARAS, subset_setting=True, drop='Qze'),
    'glyphs-RI': dict(table_name=MONO, inner=MONO, faces=['Regular', 'Italic'], paras=GLYPH_PARAS, subset_setting=True, drop='Qze'),
    'glyphs-all4': dict(table_name=MONO, inner=MONO, faces=FACES, paras=GLYPH_PARAS, subset_setting=True, drop='Qze'),
    'restricted-theme-georgia': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), fs_type=2,
                                     theme_fonts=('Georgia', 'Verdana')),
    'restricted-roman': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), fs_type=2,
                             table_props='<w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/>'),
    'restricted': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), fs_type=2),
    'noprint': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), fs_type=4),
    'editable': dict(table_name=MONO, inner=MONO, faces=['Regular'], paras=four(MONO), fs_type=8),
    'sans-reg': dict(table_name=SANS, inner=SANS, faces=['Regular'], paras=four(SANS), source='Carlito'),
}


def widths(folder):
    word = os.path.join(folder, 'word')
    rows, fonts = [], []
    for name, spec in sorted(FIXTURES.items()):
        pdf = os.path.join(word, name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        listed = subprocess.run(['pdffonts', pdf], capture_output=True, text=True, check=True).stdout.splitlines()[2:]
        names = sorted({re.sub(r'^[A-Z]{6}\+', '', l.split()[0]) for l in listed if l.strip()})
        fonts.append(f'{name}\t{" ".join(names)}')
        out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        # (A word set in two fonts - a subset face's missing characters in another - comes in pieces that
        # touch: joined again.)
        pieces = [(m.group(4), float(m.group(1)), float(m.group(2)), float(m.group(3))) for m in
                  re.finditer(r'xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', out)]
        if name.startswith('glyphs-'):
            # (Each paragraph one line: its width from its first character's left to its last's advance.)
            lines = []
            for text, x0, x1, y1 in pieces:
                if lines and abs(lines[-1][2] - y1) < 1.0:
                    lines[-1] = (lines[-1][0], x1, lines[-1][2])
                else:
                    lines.append((x0, x1, y1))
            if len(lines) != len(spec['paras']):
                print(name, 'lines:', len(lines), file=sys.stderr)
                continue
            for (family, face, runs), (x0, x1, _) in zip(spec['paras'], lines):
                rows.append(f'{name}\t{family}/{face}\t{"|".join(runs)}\t{x1 - x0:.3f}')
            continue
        joined = []
        for text, x0, x1, y1 in pieces:
            if joined and abs(joined[-1][3] - y1) < 4 and -0.3 < x0 - joined[-1][2] < 0.3:
                t, a, _, y = joined[-1]
                joined[-1] = (t + text, a, x1, y)
            else:
                joined.append((text, x0, x1, y1))
        found = [(t, x1 - x0) for t, x0, x1, _ in joined]
        expected = [(f'{family}/{face}', w) for family, face in spec['paras'] for w in WORDS.split()]
        if [w for w, _ in found] != [w for _, w in expected]:
            print(name, 'words not as written:', [w for w, _ in found][:12], file=sys.stderr)
            continue
        for (what, w), (_, width) in zip(expected, found):
            rows.append(f'{name}\t{what}\t{w}\t{width:.3f}')
    with open(os.path.join(word, 'widths.tsv'), 'w') as f:
        f.write('\n'.join(rows) + '\n')
    with open(os.path.join(word, 'fonts.tsv'), 'w') as f:
        f.write('\n'.join(fonts) + '\n')
    print(f'{len(rows)} widths')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--widths':
        widths(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/embedded-fonts')
    os.makedirs(out_dir, exist_ok=True)
    for name, spec in FIXTURES.items():
        write(os.path.join(out_dir, name + '.docx'), repo, **spec)
    print(f'{len(FIXTURES)} fixtures in {out_dir}')
