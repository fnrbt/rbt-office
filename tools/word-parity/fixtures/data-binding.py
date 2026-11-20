#!/usr/bin/env python3
"""Fixtures for content controls bound to custom XML (w:sdtPr/w:dataBinding; docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/data-binding.py [out dir]   (default tests/word/fixtures/data-binding)

then, to ask Word (winvm; a Word of its own for each document):

    tools/word-parity/fixtures/data-binding.sh tests/word/fixtures/data-binding

which runs tools/word-parity/content-controls.vbs on the VM and brings back <dir>/word/<name>.tsv (each
paragraph's text in every story, each content control's type, text and mapping) and <dir>/word/<name>.saved.docx
(the document as Word saves it after opening it) and <dir>/word/<name>.edited.docx (each mapped control's
text set to "edit <n>" through the object model, then saved: what Word writes back to the custom XML).

Word's text when it opens a document comes from the custom XML part a control is bound to, not from the
text document.xml caches. Each paragraph of `cases` is "<label>: [control] :end", the control's cached text
"cached <label>", every label a case:
  plain      a plain text control, the store's value different
  nomatch    an XPath that matches no node
  nostore    a storeItemID that no part has
  lowerid    the storeItemID in lower case, without the braces' case
  noid       no storeItemID at all
  attr       an XPath to an attribute (/root[1]/attr[1]/@x)
  ns         a namespaced store (prefixMappings xmlns:ns0='urn:fixture:ns'), the item with prefix q:
  nsdefault  the same store, the item's namespace the default one in a second store
  nopred     /root/a without [1]
  second     /root[1]/many[1]/v[2]
  empty      an empty element (<empty/>)
  multiline  a value with a line break, w:text w:multiLine="1"
  oneline    the same value, w:text without multiLine
  date1-3    a date control (dateFormat "d MMMM yyyy", en-GB) bound to 2024-03-05T00:00:00, 2024-03-05, "not a date"
  dd / ddx   a drop-down list (One/one, Two/two) bound to "two" / to "unlisted"
  combo      a combo box bound to "two"
  cb1-4      a check box bound to true, false, 1, 0 (cached unchecked, cb3 / cb4 cached checked)
  ws         "  padded   value  " (spaces at the ends and three inside)
  mixed      an element with a child element: a<b>b</b>c
  fmt        cached in two runs (bold "cached", italic " fmt"), w:sdtPr/w:rPr red
  rich       a rich text control (no type) with a w:dataBinding
  plc        cached as its placeholder (w:showingPlcHdr), the store's value given
  block      a block-level plain text control around a paragraph
`cases2`, alike:
  noidns / nostorens   no storeItemID / one no part has, the XPath matching in the second store only
  wrongstore           the storeItemID of the second store, the XPath matching in the first only
  fmt2                 cached bold "cached" and italic " fmt2", no w:sdtPr/w:rPr
  fmt3                 cached italic, w:sdtPr/w:rPr bold
  emptyplc             an empty value, the control cached as its placeholder (w:showingPlcHdr)
  emptyrpr             an empty value, w:sdtPr/w:rPr bold
  tab / amp            a value with a tab / with &amp; &lt; &gt;
  nocontent            an empty w:sdtContent
  date4-9              formats "dddd, d MMM yy", "dd/MM/yyyy HH:mm" (14:07:09), "M/d/yyyy h:mm am/pm" (en-US), de-DE
                       "MMMM d, yyyy", quoted text, a UTC time (23:30Z)
  cbx                  a check box bound to "yes"
  samefmt              the store's value the cached text, cached bold (no w:sdtPr/w:rPr)
  samefmt2             the same in two runs (bold, italic), w:sdtPr/w:rPr red
  ddsame               a drop-down list cached bold as the value's display text
  ddempty              a drop-down list bound to an empty element
Further documents: `header` (a bound control in the header and in the body), `repeat` (a w15:repeatingSection
bound to a node set of 3 items, the document holding 1), `repeat1` (its XPath
ending in item[1], as Word saves it; the document holding 2), `repeatfewer` (a set of 2, the document holding 3).
No dependencies: the package is written as XML.
"""
import os
import sys
import zipfile

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = (f'xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15"')
STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style></w:styles>'''
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W_NS}"><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>')
CT = 'application/vnd.openxmlformats-officedocument'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ID1 = '{6A2B9C10-1111-4D2E-8F00-00000000A001}'
ID2 = '{6A2B9C10-1111-4D2E-8F00-00000000A002}'
ID3 = '{6A2B9C10-1111-4D2E-8F00-00000000A003}'

STORE1 = ('<root><a>bound value A</a><attr x="attribute value"/><many><v>first</v><v>second</v></many><empty/>'
          '<ml>line one\nline two</ml><date1>2024-03-05T00:00:00</date1><date2>2024-03-05</date2><date3>not a date</date3>'
          '<dd>two</dd><ddx>unlisted</ddx><cb1>true</cb1><cb2>false</cb2><cb3>1</cb3><cb4>0</cb4>'
          '<ws>  padded   value  </ws><mixed>a<b>b</b>c</mixed><fmt>new text</fmt><rich>rich bound</rich>'
          '<plc>value over placeholder</plc><block>block value</block><hdr>header value</hdr>'
          '<tab>a\tb</tab><amp>x &amp; y &lt;z&gt;</amp><date5>2024-03-05T14:07:09</date5><date9>2024-03-05T23:30:00Z</date9><cbx>yes</cbx><same>same text</same>'
          '<items><item><name>first item</name></item><item><name>second item</name></item><item><name>third item</name></item></items>'
          '</root>')
STORE2 = '<q:data xmlns:q="urn:fixture:ns"><q:v>namespaced value</q:v></q:data>'
STORE3 = '<data xmlns="urn:fixture:other"><v>default-namespace value</v></data>'
OTHER_NS = "xmlns:ns0='urn:fixture:other'"
FIX_NS = "xmlns:ns0='urn:fixture:ns'"


def run(text, rpr=''):
    return f'<w:r>{rpr and "<w:rPr>" + rpr + "</w:rPr>"}<w:t xml:space="preserve">{text}</w:t></w:r>'


def binding(xpath, store=ID1, prefixes='', tag='w:dataBinding'):
    sid = '' if store is None else f' w:storeItemID="{store}"'
    return f'<{tag} w:prefixMappings="{prefixes}" w:xpath="{xpath}"{sid}/>'


def sdt(n, label, kind, bind, content=None, pr_rpr='', extra_pr=''):
    pr_rpr = pr_rpr and f'<w:rPr>{pr_rpr}</w:rPr>'
    content = content if content is not None else run(f'cached {label}')
    return (f'<w:sdt><w:sdtPr>{pr_rpr}<w:alias w:val="{label}"/><w:tag w:val="{label}"/><w:id w:val="{1000 + n}"/>{extra_pr}'
            f'{bind}{kind}</w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>')


def para(label, control):
    return f'<w:p>{run(label + ": ")}{control}{run(" :end")}</w:p>'


def date(fmt='d MMMM yyyy'):
    return (f'<w:date w:fullDate="2020-01-01T00:00:00Z"><w:dateFormat w:val="{fmt}"/><w:lid w:val="en-GB"/>'
            '<w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>')


LIST = '<w:listItem w:displayText="One" w:value="one"/><w:listItem w:displayText="Two" w:value="two"/>'
GOTHIC = '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic"/>'


def checkbox(checked):
    return (f'<w14:checkbox><w14:checked w14:val="{1 if checked else 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
            '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>')


def cases():
    out = []
    n = 0

    def add(label, kind, bind, **kw):
        nonlocal n
        n += 1
        out.append(para(label, sdt(n, label, kind, bind, **kw)))

    text = '<w:text/>'
    add('plain', text, binding('/root[1]/a[1]'))
    add('nomatch', text, binding('/root[1]/nothing[1]'))
    add('nostore', text, binding('/root[1]/a[1]', store='{6A2B9C10-1111-4D2E-8F00-0000000000FF}'))
    add('lowerid', text, binding('/root[1]/a[1]', store=ID1.lower()))
    add('noid', text, binding('/root[1]/a[1]', store=None))
    add('attr', text, binding('/root[1]/attr[1]/@x'))
    add('ns', text, binding('/ns0:data[1]/ns0:v[1]', store=ID2, prefixes=FIX_NS))
    add('nsdefault', text, binding('/ns0:data[1]/ns0:v[1]', store=ID3, prefixes=OTHER_NS))
    add('nopred', text, binding('/root/a'))
    add('second', text, binding('/root[1]/many[1]/v[2]'))
    add('empty', text, binding('/root[1]/empty[1]'))
    add('multiline', '<w:text w:multiLine="1"/>', binding('/root[1]/ml[1]'))
    add('oneline', text, binding('/root[1]/ml[1]'))
    for i in (1, 2, 3):
        add(f'date{i}', date(), binding(f'/root[1]/date{i}[1]'))
    add('dd', f'<w:dropDownList>{LIST}</w:dropDownList>', binding('/root[1]/dd[1]'))
    add('ddx', f'<w:dropDownList>{LIST}</w:dropDownList>', binding('/root[1]/ddx[1]'))
    add('combo', f'<w:comboBox>{LIST}</w:comboBox>', binding('/root[1]/dd[1]'))
    for i, cached in ((1, False), (2, False), (3, True), (4, True)):
        add(f'cb{i}', checkbox(cached), binding(f'/root[1]/cb{i}[1]'),
            content=run('☒' if cached else '☐', GOTHIC))
    add('ws', text, binding('/root[1]/ws[1]'))
    add('mixed', text, binding('/root[1]/mixed[1]'))
    add('fmt', text, binding('/root[1]/fmt[1]'), content=run('cached', '<w:b/>') + run(' fmt', '<w:i/>'),
        pr_rpr='<w:color w:val="FF0000"/>')
    add('rich', '', binding('/root[1]/rich[1]'))
    add('plc', text, binding('/root[1]/plc[1]'), content=run('Click here to enter text.', '<w:rStyle w:val="PlaceholderText"/>'),
        extra_pr='<w:showingPlcHdr/>')
    n += 1
    out.append(sdt(n, 'block', '<w:text/>', binding('/root[1]/block[1]'),
                   content=f'<w:p>{run("cached block")}</w:p>'))
    return ''.join(out)


def cases2():
    out = []
    n = 100

    def add(label, kind, bind, **kw):
        nonlocal n
        n += 1
        out.append(para(label, sdt(n, label, kind, bind, **kw)))

    text = '<w:text/>'
    add('noidns', text, binding('/ns0:data[1]/ns0:v[1]', store=None, prefixes=FIX_NS))
    add('nostorens', text, binding('/ns0:data[1]/ns0:v[1]', store='{6A2B9C10-1111-4D2E-8F00-0000000000FF}', prefixes=FIX_NS))
    add('wrongstore', text, binding('/root[1]/a[1]', store=ID2))
    add('fmt2', text, binding('/root[1]/fmt[1]'), content=run('cached', '<w:b/>') + run(' fmt2', '<w:i/>'))
    add('fmt3', text, binding('/root[1]/fmt[1]'), content=run('cached fmt3', '<w:i/>'), pr_rpr='<w:b/>')
    add('emptyplc', text, binding('/root[1]/empty[1]'), content=run('Click here to enter text.', '<w:rStyle w:val="PlaceholderText"/>'),
        extra_pr='<w:showingPlcHdr/>')
    add('emptyrpr', text, binding('/root[1]/empty[1]'), pr_rpr='<w:b/>')
    add('tab', text, binding('/root[1]/tab[1]'))
    add('amp', text, binding('/root[1]/amp[1]'))
    add('nocontent', text, binding('/root[1]/a[1]'), content='')
    add('date4', date('dddd, d MMM yy'), binding('/root[1]/date1[1]'))
    add('date5', date('dd/MM/yyyy HH:mm'), binding('/root[1]/date5[1]'))
    add('date6', date('M/d/yyyy h:mm am/pm').replace('en-GB', 'en-US'), binding('/root[1]/date5[1]'))
    add('date7', date('MMMM d, yyyy').replace('en-GB', 'de-DE'), binding('/root[1]/date1[1]'))
    add('date8', date("'Day' d 'of' MMMM"), binding('/root[1]/date1[1]'))
    add('date9', date(), binding('/root[1]/date9[1]'))
    add('cbx', checkbox(False), binding('/root[1]/cbx[1]'), content=run('\u2610', GOTHIC))
    add('samefmt', text, binding('/root[1]/same[1]'), content=run('same text', '<w:b/>'))
    add('samefmt2', text, binding('/root[1]/same[1]'), content=run('same', '<w:b/>') + run(' text', '<w:i/>'), pr_rpr='<w:color w:val="FF0000"/>')
    add('ddsame', f'<w:dropDownList>{LIST}</w:dropDownList>', binding('/root[1]/dd[1]'), content=run('Two', '<w:b/>'))
    add('ddempty', f'<w:dropDownList>{LIST}</w:dropDownList>', binding('/root[1]/empty[1]'), content=run('One'))
    return ''.join(out)


def header_doc():
    body = para('body', sdt(1, 'body', '<w:text/>', binding('/root[1]/a[1]')))
    header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
              + para('hdr', sdt(2, 'hdr', '<w:text/>', binding('/root[1]/hdr[1]'))) + '</w:hdr>')
    return body, header


def repeat_doc(section_xpath='/root[1]/items[1]/item', item_xpath='/root[1]/items[1]/item[{k}]/name[1]', held=1):
    def item(k):
        inner = sdt(10 + k, 'name', '<w:text/>', binding(item_xpath.format(k=k)), content=run(f'cached item {k}'))
        return f'<w:sdt><w:sdtPr><w:id w:val="{1100 + k}"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:p>{run("Item: ")}{inner}</w:p></w:sdtContent></w:sdt>'
    section = (f'<w:sdt><w:sdtPr><w:alias w:val="items"/><w:id w:val="1001"/>'
               f'{binding(section_xpath, tag="w15:dataBinding")}<w15:repeatingSection/></w:sdtPr>'
               f'<w:sdtContent>{"".join(item(k) for k in range(1, held + 1))}</w:sdtContent></w:sdt>')
    return f'<w:p>{run("Before")}</w:p>{section}<w:p>{run("After")}</w:p>'


def write(path, body, header=None):
    sect = '<w:sectPr>' + ('<w:headerReference w:type="default" r:id="rId9"/>' if header else '') + \
        '<w:pgSz w:w="11907" w:h="16839"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
    stores = [(ID1, STORE1), (ID2, STORE2), (ID3, STORE3)]
    types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
             f'<Override PartName="/word/document.xml" ContentType="{CT}.wordprocessingml.document.main+xml"/>'
             f'<Override PartName="/word/styles.xml" ContentType="{CT}.wordprocessingml.styles+xml"/>'
             f'<Override PartName="/word/settings.xml" ContentType="{CT}.wordprocessingml.settings+xml"/>'
             + (f'<Override PartName="/word/header1.xml" ContentType="{CT}.wordprocessingml.header+xml"/>' if header else '')
             + ''.join(f'<Override PartName="/customXml/itemProps{i}.xml" ContentType="{CT}.customXmlProperties+xml"/>' for i in range(1, len(stores) + 1))
             + '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>')
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/>'
                + ''.join(f'<Relationship Id="rId{2 + i}" Type="{REL}/customXml" Target="../customXml/item{i}.xml"/>' for i in range(1, len(stores) + 1))
                + (f'<Relationship Id="rId9" Type="{REL}/header" Target="header1.xml"/>' if header else '')
                + '</Relationships>')
    parts = [('[Content_Types].xml', types), ('_rels/.rels', rels), ('word/_rels/document.xml.rels', doc_rels),
             ('word/document.xml', doc), ('word/styles.xml', STYLES), ('word/settings.xml', SETTINGS)]
    if header:
        parts.append(('word/header1.xml', header))
    for i, (sid, xml) in enumerate(stores, 1):
        parts.append((f'customXml/item{i}.xml', xml))
        parts.append((f'customXml/itemProps{i}.xml', '<?xml version="1.0" encoding="UTF-8" standalone="no"?>'
                      f'<ds:datastoreItem ds:itemID="{sid}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>'))
        parts.append((f'customXml/_rels/item{i}.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                      f'<Relationship Id="rId1" Type="{REL}/customXmlProps" Target="itemProps{i}.xml"/></Relationships>'))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            z.writestr(zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0)), data, zipfile.ZIP_DEFLATED)


def docs():
    body, header = header_doc()
    return {'cases': (cases(), None), 'cases2': (cases2(), None), 'header': (body, header), 'repeat': (repeat_doc(), None),
            'repeat1': (repeat_doc('/root[1]/items[1]/item[1]', held=2), None),
            'repeatfewer': (repeat_doc('/root[1]/many[1]/v', '/root[1]/many[1]/v[{k}]', held=3), None)}


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/data-binding')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    ds = docs()
    for name, (body, header) in ds.items():
        write(os.path.join(out_dir, name + '.docx'), body, header)
    print(f'{len(ds)} fixtures in {out_dir}')
