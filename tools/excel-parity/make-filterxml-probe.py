#!/usr/bin/env python3
"""Probe cells for FILTERXML (FnWeb.fs), as eval-cells-on-vm.sh takes them.

    python3 tools/excel-parity/make-filterxml-probe.py > filterxml-cells.txt

One line per cell: sheet, address, formula. Each case is an XML text and an XPath 1.0 expression; for each
the probe asks FILTERXML itself (the first node in a cell of Excel 2019, which has no spilling), ROWS and
COLUMNS of the result and INDEX of each of its first nodes: which nodes the expression selects, in which
order, how a node becomes a cell value (the text of an element, an attribute's value, numbers converted),
what a number, string or boolean result of an expression is, and the errors (no node, XML or XPath that does
not parse, namespaces, DTDs). Deterministic: filterxml-excel.tsv holds Excel 2019's values;
function-parity-engine.fsx compares the engine.
"""
import re

cells = []


def cell(f):
    cells.append(f)


def q(s):
    """s as a formula's text: pieces of at most 200 characters (a literal is at most 255), line feeds and
    tabs as CHAR(10) and CHAR(9) (a probe line is one line)."""
    parts = []
    for chunk in re.split(r'([\t\n])', s):
        if chunk == '\t':
            parts.append('CHAR(9)')
        elif chunk == '\n':
            parts.append('CHAR(10)')
        elif chunk:
            parts += ['"' + chunk[j:j + 200].replace('"', '""') + '"' for j in range(0, len(chunk), 200)]
    return '&'.join(parts) or '""'


DOC = "<r><a id='1' n='x'>10</a><a id='2'>2.5</a><a id='3' n='y'>abc</a><b><a id='4'>-7</a><c>t</c></b><a id='5'/></r>"
BOOKS = ("<lib><book cat='f' year='2001'><title lang='en'>Dune</title><price>9.99</price></book>"
         "<book cat='s' year='1999'><title lang='de'>Faust</title><price>12</price></book>"
         "<book cat='f' year='2010'><title lang='en'>Hyperion</title><price>7.5</price></book>"
         "<mag><title>Wired</title><price>5</price></mag></lib>")
CASES = [
    (DOC, '//a'), (DOC, '/r/a'), (DOC, '//a/@id'), (DOC, '//@n'), (DOC, '//a[@id>2]'), (DOC, '//a[2]'), (DOC, '(//a)[2]'),
    (DOC, '//a[last()]'), (DOC, '(//a)[last()]'), (DOC, '//a[position()<3]'), (DOC, '//b/a'), (DOC, '//b'), (DOC, '/r'),
    (DOC, '//c|//a[1]'), (DOC, '//a[@n]'), (DOC, '//a[not(@n)]'), (DOC, '//a[@n="y"]'), (DOC, "//a[@n='x']/@id"),
    (DOC, '//a/text()'), (DOC, '//b//text()'), (DOC, '//*'), (DOC, '/*/*'), (DOC, '//a[.=10]'), (DOC, '//a[. > 0]'),
    (DOC, '//c/..'), (DOC, '//c/parent::*'), (DOC, '//c/ancestor::*'), (DOC, '//c/ancestor-or-self::*'),
    (DOC, '//a[@id=2]/following-sibling::*'), (DOC, '//a[@id=3]/preceding-sibling::a'), (DOC, '//b/following::a'),
    (DOC, '//c/preceding::a'), (DOC, '//b/descendant::*'), (DOC, '//b/descendant-or-self::*'), (DOC, '//a/self::a'),
    (DOC, '//a[contains(.,"b")]'), (DOC, '//a[starts-with(.,"1")]'), (DOC, '//a[string-length(.)=3]'),
    (DOC, '//a[@id mod 2=1]'), (DOC, '//a[@id div 2=1]'), (DOC, '//a[@id*2=4]'), (DOC, '//a[@id+1=4]'), (DOC, '//a[@id-1=3]'),
    (DOC, '//a[@id=1 or @id=5]'), (DOC, '//a[@id>1 and @id<4]'), (DOC, '//a[@id!=1]'), (DOC, '//a[@id<=2]'),
    (DOC, '//a[@id>=4]'), (DOC, '//a[-@id=-2]'), (DOC, '//a[number(.)=2.5]'), (DOC, '//a[name()="a"][1]'),
    (DOC, '//*[local-name()="c"]'), (DOC, '//a[translate(.,"abc","ABC")="ABC"]'), (DOC, '//a[substring(.,2,1)="b"]'),
    (DOC, '//a[substring-before(@n,"y")=""]'), (DOC, '//a[concat(@id,@n)="1x"]'), (DOC, '//a[normalize-space(.)!=""]'),
    (DOC, '//a[boolean(@n)]'), (DOC, '//a[true()]'), (DOC, '//a[false()]'), (DOC, '//a[floor(.)=2]'),
    (DOC, '//a[ceiling(.)=3]'), (DOC, '//a[round(.)=3]'), (DOC, '//a[count(../a)=4]'), (DOC, '//b[count(*)=2]'),
    (DOC, '//a[sum(../a)>0]'), (DOC, '//node()'), (DOC, '//a[1]/@*'), (DOC, '//a[@id="5"]'), (DOC, '//a[.=""]'),
    (DOC, '//x'), (DOC, '//a[@id=9]'), (DOC, 'count(//a)'), (DOC, 'sum(//a/@id)'), (DOC, 'string(//a)'), (DOC, '1+1'),
    (DOC, 'true()'), (DOC, '//a[1]=10'), (DOC, 'concat("a","b")'), (DOC, '//'), (DOC, '//a['), (DOC, 'a'), (DOC, 'r/a'),
    (DOC, '/'), (DOC, '.'), (DOC, '//a[@id=1]/@id'), (DOC, '//comment()'), (DOC, "//processing-instruction()"),
    (DOC, '//a[position()=last()-1]'), (DOC, "//a[@id='1']/following::*[1]"), (DOC, '//b/*[2]'), (DOC, '//b/a/@id|//a[1]'),
    (DOC, ' //a[1] '), (DOC, '//A'), (DOC, '//a[id(1)]'), (DOC, '//a[lang("en")]'),
    (BOOKS, '//title'), (BOOKS, '//book[price>10]/title'), (BOOKS, '//book/@year'), (BOOKS, '//title[@lang="en"]'),
    (BOOKS, '//book[@cat="f"][2]/title'), (BOOKS, '//book[last()]/price'), (BOOKS, '//*[price<8]/title'),
    (BOOKS, '/lib/*/title'), (BOOKS, '//book[title="Faust"]/@year'), (BOOKS, '//price'), (BOOKS, '//book[1]'),
    (BOOKS, '//mag'), (BOOKS, 'sum(//price)'), (BOOKS, '//book[price=//book/price[.>10]]/title'),
    ("<a>1<b>2</b>3</a>", '/a'), ("<a>1<b>2</b>3</a>", '/a/text()'), ("<a> 12 </a>", '/a'), ("<a>\n 12\n</a>", '/a'),
    ("<a>1e3</a>", '/a'), ("<a>1E-2</a>", '/a'), ("<a>TRUE</a>", '/a'), ("<a>true</a>", '/a'), ("<a>$5</a>", '/a'),
    ("<a>50%</a>", '/a'), ("<a>1,5</a>", '/a'), ("<a>1,000</a>", '/a'), ("<a>2020-01-02</a>", '/a'), ("<a>1/2/2020</a>", '/a'),
    ("<a>12:30</a>", '/a'), ("<a>-0</a>", '/a'), ("<a>+5</a>", '/a'), ("<a>.5</a>", '/a'), ("<a>5.</a>", '/a'),
    ("<a>0x10</a>", '/a'), ("<a>(5)</a>", '/a'), ("<a>1 2</a>", '/a'), ("<a>12345678901234567890</a>", '/a'),
    ("<a>0.1234567890123456789</a>", '/a'), ("<a>1e400</a>", '/a'), ("<a>#N/A</a>", '/a'), ("<a>Infinity</a>", '/a'),
    ("<a>NaN</a>", '/a'), ("<a>00012</a>", '/a'), ("<a></a>", '/a'), ("<a/>", '/a'), ("<a><b/></a>", '/a'),
    ("<a>&amp;&lt;&gt;&quot;&apos;</a>", '/a'), ("<a>&#65;&#x42;</a>", '/a'), ("<a><![CDATA[<x>]]></a>", '/a'),
    ("<a><!-- c -->x</a>", '/a'), ("<a><!-- c -->x</a>", '//comment()'), ("<?xml version='1.0'?><a>1</a>", '/a'),
    ("<?xml version='1.0' encoding='UTF-8'?><a>1</a>", '/a'), ("<a>1</a><b>2</b>", '/a'), ("<a>1", '/a'), ("a", '/a'),
    ("", '/a'), ("<a><b></a></b>", '/a'), ("<a x='1' x='2'/>", '/a'), ("<a xmlns='urn:x'><b>1</b></a>", '//b'),
    ("<a xmlns='urn:x'><b>1</b></a>", "//*[local-name()='b']"), ("<a xmlns:p='urn:p'><p:b>1</p:b></a>", '//p:b'),
    ("<a xmlns:p='urn:p'><p:b>1</p:b></a>", "//*[local-name()='b']"), ("<a xmlns:p='urn:p'><p:b>1</p:b></a>", "//*[name()='p:b']"),
    ("<p:a xmlns:p='urn:p'>1</p:a>", '/*'), ("<a><b>x</b><b>y</b><b>z</b></a>", '//b'),
    ("<a><b>1</b><b>x</b><b>3</b></a>", '//b'), ("<!DOCTYPE a [<!ENTITY e 'hi'>]><a>&e;</a>", '/a'), ("<!DOCTYPE a><a>1</a>", '/a'),
    ("<a>café 中</a>", '/a'), ("<a b='x &amp; y'/>", '/a/@b'), ("<a b='1'/>", '/a/@b'), ("<a b=' 2 '/>", '/a/@b'),
    ("<a>\t<b>1</b>\n</a>", '/a'), ("<a>\t<b>1</b>\n</a>", '/a/node()'), ("<a>  x  y  </a>", '/a'),
    ("<a>x</a>", '/a/@nope'), ("<a><b>1</b></a>", 'name(/a/b)'), ("<a><b>1</b></a>", '/a/b[1]/text()[1]'),
    ("<root><i>1</i><i>2</i><i>3</i><i>4</i><i>5</i><i>6</i><i>7</i></root>", '//i'),
    ("<root><i>1</i><i>2</i><i>3</i><i>4</i><i>5</i><i>6</i><i>7</i></root>", '//i[position() mod 2 = 0]'),
    ("<root><i>1</i><i>2</i></root>", '//i[. = 2] | //i[. = 1]'), ("<r><a>x</a></r>", '//a/following-sibling::a'),
    ("<a> <b>1</b> </a>", '/a/node()'), ("<a> </a>", '/a'), ("<a>1</a>", '/a/text()/text()'),
]
for xml, xp in CASES:
    x, p = q(xml), q(xp)
    cell('=FILTERXML(%s,%s)' % (x, p))
    cell('=ROWS(FILTERXML(%s,%s))' % (x, p))
    cell('=COLUMNS(FILTERXML(%s,%s))' % (x, p))
    for k in range(1, 5):
        cell('=INDEX(FILTERXML(%s,%s),%d)' % (x, p, k))
for f in ['FILTERXML(1,"/a")', 'FILTERXML("<a>1</a>",1)', 'FILTERXML(TRUE,"/a")', 'FILTERXML("<a>1</a>",1/0)',
          'FILTERXML(1/0,"/a")', 'FILTERXML(Z1,"/a")', 'FILTERXML("<a>1</a>",Z1)', 'FILTERXML("<a>1</a>","")',
          'FILTERXML("<a>TRUE</a>","/a")=TRUE', 'ISTEXT(FILTERXML("<a>1</a>","/a"))', 'FILTERXML("<a>5</a>","/a")+1',
          'FILTERXML(REPT("<a>1</a>",1),"//a")', 'SUM(FILTERXML("<r><a>1</a><a>2</a></r>","//a"))',
          'COUNT(FILTERXML("<r><a>1</a><a>x</a></r>","//a"))', 'FILTERXML("<a>"&REPT("x",300)&"</a>","/a")=REPT("x",300)',
          'LEN(FILTERXML("<a>"&REPT("x",300)&"</a>","/a"))', 'FILTERXML("<a>"&REPT("x",40000)&"</a>","/a")',
          'FILTERXML({"<a>1</a>","<a>2</a>"},"/a")', 'INDEX(FILTERXML({"<a>1</a>","<a>2</a>"},"/a"),1,2)']:
    cell('=' + f)

for i, f in enumerate(cells):
    print('S\tA%d\t%s' % (i + 1, f))
