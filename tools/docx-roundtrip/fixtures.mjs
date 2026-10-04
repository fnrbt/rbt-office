// Synthetic documents for roundtrip.mjs: one small package per feature the
// DOCX writer has to carry through a save. `expect` entries are checked by
// page.js: {part, has|lacks} are regexes over a written part, {readText} /
// {savedText} look at the text the reader produced before / after the save,
// {mutate} is JS run on the exported package before writing (simulates an edit).

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');
const PNG2 = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

const p = (text, pPr = '', rPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

const styles = `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const numbering = `<w:numbering ${W}>
<w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="11"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="12"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="3"><w:abstractNumId w:val="10"/></w:num><w:num w:numId="4"><w:abstractNumId w:val="11"/></w:num><w:num w:numId="5"><w:abstractNumId w:val="12"/></w:num><w:num w:numId="6"><w:abstractNumId w:val="10"/></w:num>
</w:numbering>`;

const li = (text, numId, ilvl = 0) => p(text, `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
const cell = (content, tcPr = '') => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>${tcPr}</w:tcPr>${content}</w:tc>`;
const drawing = (rid, cx, cy) => `<w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="pic"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="pic"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
// A wps shape in mc:AlternateContent with the VML fallback Word writes next to it.
const shapeAnchor = (id, name, o, spPr, text) =>
    `<w:r><mc:AlternateContent><mc:Choice Requires="wps" xmlns:wps="${WPS}"><w:drawing><wp:anchor distT="${o.dist}" distB="${o.dist}" distL="${o.dist}" distR="${o.dist}" simplePos="0" relativeHeight="${o.z}" behindDoc="${o.behind}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${o.relH}"><wp:posOffset>${o.x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${o.relV}"><wp:posOffset>${o.y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${o.cx}" cy="${o.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${o.wrap}<wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr>${spPr}</wps:spPr>${text ? `<wps:txbx><w:txbxContent>${text}</w:txbxContent></wps:txbx>` : ''}<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="vml${id}" type="#_x0000_t202" style="position:absolute;margin-left:10pt;margin-top:10pt;width:50pt;height:30pt"><v:textbox><w:txbxContent>${text || '<w:p/>'}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

const docs = [
    {
        name: 'synthetic/units',
        body:
            p('twelve point', '', '<w:sz w:val="24"/>') +
            p('eleven point', '', '<w:sz w:val="22"/>') +
            p('ten and a half', '', '<w:sz w:val="21"/>') +
            p('nine point condensed', '', '<w:spacing w:val="-7"/><w:sz w:val="18"/>') +
            p('raised', '', '<w:position w:val="6"/>') +
            p('indented', '<w:ind w:left="720" w:right="367" w:firstLine="283"/>') +
            p('hanging', '<w:ind w:left="1134" w:hanging="567"/>') +
            p('spaced', '<w:spacing w:before="120" w:after="200"/>') +
            p('odd spacing', '<w:spacing w:before="113" w:after="57"/>') +
            p('one and a half lines', '<w:spacing w:line="360" w:lineRule="auto"/>') +
            p('multiple 1.15', '<w:spacing w:line="276" w:lineRule="auto"/>') +
            p('exactly 14pt', '<w:spacing w:line="280" w:lineRule="exact"/>') +
            p('at least 15pt', '<w:spacing w:line="300" w:lineRule="atLeast"/>', '<w:sz w:val="28"/>') +
            p('no size at all'),
        expect: [
            { part: 'word/document.xml', has: '<w:sz w:val="24"/>' }, { part: 'word/document.xml', has: '<w:sz w:val="21"/>' },
            { part: 'word/document.xml', has: 'w:left="720"' }, { part: 'word/document.xml', has: 'w:right="367"' },
            { part: 'word/document.xml', has: 'w:line="280" w:lineRule="exact"' }, { part: 'word/document.xml', has: 'w:line="300" w:lineRule="atLeast"' },
            { part: 'word/document.xml', has: 'w:before="113"' }, { part: 'word/document.xml', has: '<w:position w:val="6"/>' },
            { part: 'word/document.xml', has: '<w:spacing w:val="-7"/>' },
        ],
    },
    {
        name: 'synthetic/lists',
        body: li('one', 3) + li('two', 3) + li('nested a', 3, 1) + p('between') + li('III', 4) + li('IV', 4) + li('bullet', 5) + li('restarts at one', 6) + li('two again', 6),
        expect: [
            { part: 'word/document.xml', has: '<w:numId w:val="4"/>' }, { part: 'word/document.xml', has: '<w:numId w:val="6"/>' },
            { part: 'word/numbering.xml', has: 'upperRoman' }, { part: 'word/numbering.xml', has: 'lowerLetter' },
            { part: 'word/numbering.xml', has: '<w:start w:val="3"/>' },
        ],
    },
    {
        name: 'synthetic/tables',
        body:
            `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblCellSpacing w:w="30" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="12" w:color="FF0000"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="double" w:sz="6" w:color="0000FF"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="dashed" w:sz="8" w:color="00FF00"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="150" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
            `<w:tr><w:trPr><w:trHeight w:val="567" w:hRule="exact"/></w:trPr>${cell(p('a1'), '<w:tcMar><w:top w:w="45" w:type="dxa"/><w:left w:w="300" w:type="dxa"/><w:bottom w:w="45" w:type="dxa"/><w:right w:w="300" w:type="dxa"/></w:tcMar>')}${cell(p('b1'))}</w:tr>` +
            `<w:tr>${cell(p('before nested') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>${p('inner')}</w:tc></w:tr></w:tbl>` + p('after nested'))}${cell(p('b2'))}</w:tr></w:tbl>` + p('after table'),
        expect: [
            { part: 'word/document.xml', has: '<w:tbl>.*<w:tbl>' }, { savedText: 'inner' },
            { part: 'word/document.xml', has: '<w:left w:w="300" w:type="dxa"/>' }, { part: 'word/document.xml', has: '<w:tblCellMar>' },
            { part: 'word/document.xml', has: '<w:top w:val="single" w:sz="12"' },
        ],
    },
    {
        name: 'synthetic/tabs-breaks',
        body:
            `<w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs></w:pPr><w:r><w:t>left</w:t><w:tab/><w:t>middle</w:t><w:tab/><w:t>right</w:t></w:r></w:p>` +
            `<w:p><w:r><w:t>line one</w:t><w:br/><w:t>line two</w:t></w:r></w:p>` +
            `<w:p><w:r><w:t>e</w:t><w:noBreakHyphen/><w:t>mail co</w:t><w:softHyphen/><w:t>operate</w:t></w:r><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r><w:r><w:t>x</w:t><w:cr/><w:t>y</w:t></w:r></w:p>`,
        expect: [
            { part: 'word/document.xml', has: '<w:t[^>]*>left</w:t></w:r><w:r>(<w:rPr>.*?</w:rPr>)?<w:tab/>' }, { part: 'word/document.xml', has: '<w:br/>' },
            { part: 'word/document.xml', has: '<w:tab w:val="center" w:pos="4513"/>' }, { part: 'word/document.xml', has: '<w:tab w:val="right" w:leader="dot" w:pos="9026"/>' },
            { readText: 'e‑mail co­operate' }, { savedText: 'e‑mail co­operate' }, { readText: 'x\ny' },
        ],
    },
    {
        name: 'synthetic/run-props',
        body:
            p('hidden', '', '<w:vanish/>') + p('double strike', '', '<w:dstrike/>') +
            p('wavy red underline', '', '<w:u w:val="wave" w:color="FF0000"/>') + p('double underline', '', '<w:u w:val="double"/>') +
            p('caps', '', '<w:caps/>') + p('small caps', '', '<w:smallCaps/>') + p('highlight super', '', '<w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/>') +
            p('right to left', '<w:bidi/>') +
            p('boxed', '<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/><w:left w:val="single" w:sz="8" w:space="4" w:color="FF0000"/><w:bottom w:val="double" w:sz="12" w:space="1" w:color="0000FF"/><w:right w:val="single" w:sz="8" w:space="4" w:color="FF0000"/></w:pBdr>') +
            p('shaded', '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>') + p('kept', '<w:keepNext/><w:keepLines/>'),
        expect: [
            { part: 'word/document.xml', has: '<w:vanish/>' }, { part: 'word/document.xml', has: '<w:dstrike/>' },
            { part: 'word/document.xml', has: '<w:u w:val="wave" w:color="FF0000"/>' }, { part: 'word/document.xml', has: '<w:bidi/>' },
            { part: 'word/document.xml', has: '<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/>' },
            // keepNext / keepLines set on the paragraph itself survive the save, once.
            { part: 'word/document.xml', has: '<w:pPr><w:keepNext/><w:keepLines/>' }, { part: 'word/document.xml', lacks: '<w:keepNext/>.*<w:keepNext/>' },
        ],
    },
    {
        name: 'synthetic/revisions',
        body: `<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:ins w:id="7" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="8" w:author="Bob" w:date="2024-01-02T03:04:05Z"><w:r><w:rPr><w:b/></w:rPr><w:delText>removed</w:delText></w:r></w:del><w:moveTo w:id="9" w:author="Cy"><w:r><w:t xml:space="preserve"> moved here</w:t></w:r></w:moveTo></w:p>`,
        expect: [
            { part: 'word/document.xml', has: '<w:del [^>]*w:author="Bob"[^>]*><w:r>(<w:rPr>.*?</w:rPr>)?<w:delText[^>]*>removed</w:delText>' },
            { part: 'word/document.xml', has: '<w:ins [^>]*w:author="Ann"' }, { readText: 'moved here' },
        ],
    },
    {
        name: 'synthetic/control-chars',
        body: p('PLACEHOLDER') + p('plain'),
        expect: [
            { mutate: 'pkg.Paragraphs[0].Runs[0].Text = "bell\\u0007 vt\\u000B ff\\u000C esc\\u001B nul\\u0000 ok \\uFFFE end";' },
            { savedText: 'bell vt ff esc nul ok  end' },
        ],
    },
    {
        name: 'synthetic/page-setup',
        body: p('first section') + `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:pPr></w:p>` + p('second section'),
        sect: '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1417" w:left="1701" w:header="709" w:footer="567" w:gutter="0"/><w:cols w:num="2" w:space="708" w:sep="1"/><w:titlePg/></w:sectPr>',
        expect: [
            { part: 'word/document.xml', has: '<w:pgSz w:w="11906" w:h="16838"/>' },
            { part: 'word/document.xml', has: 'w:bottom="1417" w:left="1701" w:header="709" w:footer="567"' },
            { part: 'word/document.xml', has: '<w:cols w:num="2" w:sep="1" w:space="708"/>' },
        ],
    },
    {
        name: 'synthetic/images',
        body: `<w:p>${drawing('rId1', 1371600, 685800)}</w:p>` + p('text'),
        sect: `<w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`,
        rels: `<Relationship Id="rId1" Type="${REL}/image" Target="media/body.png"/><Relationship Id="rId9" Type="${REL}/header" Target="header1.xml"/>`,
        extra: {
            'word/media/body.png': PNG,
            'word/media/head.png': PNG2,
            'word/header1.xml': `<w:hdr ${W}><w:p>${drawing('rId1', 457200, 457200)}</w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p('header table cell')}</w:tc></w:tr></w:tbl></w:hdr>`,
            'word/_rels/header1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/image" Target="media/head.png"/></Relationships>`,
        },
        types: '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
        expect: [
            { part: 'word/document.xml', has: '<wp:extent cx="1371600" cy="685800"/>' },
            { part: 'word/header1.xml', has: '<wp:extent cx="457200" cy="457200"/>' },
            { part: 'word/header1.xml', has: '<w:tbl>' }, { savedText: 'header table cell' },
            { read: 'pkg.Paragraphs.0.Runs.0.ImageSrc.length', equals: 22 + Math.ceil(PNG.length / 3) * 4 },
            { read: 'pkg.HeaderParts.0.Paragraphs.0.Runs.0.ImageSrc.length', equals: 22 + Math.ceil(PNG2.length / 3) * 4 },
        ],
    },
    {
        // Direct formatting that switches a style's formatting off must be written
        // (w:val="0", the 11pt size, w:jc left); formatting a run merely inherits must not.
        name: 'synthetic/explicit-off',
        body:
            `<w:p><w:pPr><w:pStyle w:val="Loud"/></w:pPr><w:r><w:t xml:space="preserve">inherits </w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:sz w:val="22"/></w:rPr><w:t>quiet</w:t></w:r></w:p>` +
            p('left under centred', '<w:pStyle w:val="Loud"/><w:jc w:val="left"/>') + p('plain'),
        extra: {
            'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Loud"><w:name w:val="Loud"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:i/><w:sz w:val="40"/></w:rPr></w:style></w:styles>`,
        },
        expect: [
            { part: 'word/document.xml', has: '<w:b w:val="0"/>[^>]*(<[^t][^>]*>)*<w:sz w:val="22"/>' },
            { part: 'word/document.xml', has: '<w:jc w:val="left"/>' },
            { part: 'word/document.xml', lacks: '<w:i( w:val="0")?/>' }, { part: 'word/document.xml', lacks: '<w:b/>' },
            { part: 'word/document.xml', lacks: '<w:sz w:val="40"/>' }, { part: 'word/document.xml', lacks: '<w:jc w:val="center"/>' },
            { part: 'word/document.xml', lacks: '<w:u ' }, { part: 'word/document.xml', lacks: '<w:caps' },
            { part: 'word/styles.xml', has: '<w:b/><w:i/>' }, { part: 'word/styles.xml', has: '<w:jc w:val="center"/>' },
            { read: 'pkg.Paragraphs.0.Runs.1.ExplicitBold', equals: null }, { read: 'pkg.Paragraphs.0.Runs.1.Bold', equals: null },
            { read: 'pkg.Paragraphs.0.Runs.0.Bold', equals: true }, { read: 'pkg.Paragraphs.0.Runs.0.FontSize', equals: 26.667 },
            { read: 'pkg.Paragraphs.0.Runs.1.FontSize', equals: 14.667 }, { read: 'pkg.Paragraphs.0.Alignment', equals: 'center' },
            { read: 'pkg.Paragraphs.1.Alignment', equals: 'left' },
        ],
    },
    {
        name: 'synthetic/reader-structure',
        body:
            `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
            `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent>${p('sdt in cell')}</w:sdtContent></w:sdt></w:tc><w:sdt><w:sdtContent>${cell(p('sdt cell'))}</w:sdtContent></w:sdt></w:tr>` +
            `<w:sdt><w:sdtContent><w:tr>${cell(p('sdt row a'))}${cell(p('sdt row b'))}</w:tr></w:sdtContent></w:sdt></w:tbl>` +
            `<w:p><w:smartTag w:uri="u" w:element="e"><w:r><w:t>smart</w:t></w:r></w:smartTag><w:customXml w:element="x"><w:r><w:t xml:space="preserve"> custom</w:t></w:r></w:customXml><w:dir w:val="rtl"><w:r><w:t xml:space="preserve"> dir</w:t></w:r></w:dir><w:bdo w:val="ltr"><w:r><w:t xml:space="preserve"> bdo</w:t></w:r></w:bdo><mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t xml:space="preserve"> choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t xml:space="preserve"> fallback</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>`,
        expect: [
            { readText: 'sdt in cell' }, { readText: 'sdt cell' }, { readText: 'sdt row a' }, { readText: 'sdt row b' },
            { readText: 'smart custom dir bdo fallback' },
        ],
    },
    {
        // Footnote, endnote and comment bodies are blocks: formatting, a hyperlink
        // (in the part's own relationships) and a table survive the save.
        name: 'synthetic/rich-notes',
        body:
            `<w:p><w:r><w:t xml:space="preserve">Body with a note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>` +
            `<w:r><w:t xml:space="preserve"> and an endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r>` +
            `<w:r><w:t xml:space="preserve"> and a comment</w:t></w:r><w:r><w:commentReference w:id="5"/></w:r></w:p>`,
        rels:
            `<Relationship Id="rId20" Type="${REL}/footnotes" Target="footnotes.xml"/><Relationship Id="rId21" Type="${REL}/endnotes" Target="endnotes.xml"/>` +
            `<Relationship Id="rId22" Type="${REL}/comments" Target="comments.xml"/>`,
        types:
            '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
            '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>' +
            '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>',
        extra: {
            'word/footnotes.xml': `<w:footnotes ${W}>` +
                `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
                `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
                `<w:footnote w:id="2"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> See </w:t></w:r>` +
                `<w:r><w:rPr><w:b/></w:rPr><w:t>the bold source</w:t></w:r><w:r><w:t xml:space="preserve"> at </w:t></w:r>` +
                `<w:hyperlink r:id="rId1"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>the site</w:t></w:r></w:hyperlink></w:p>` +
                `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>` +
                `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>${p('note cell a')}</w:tc><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>${p('note cell b')}</w:tc></w:tr></w:tbl>` +
                `${p('after the note table')}</w:footnote></w:footnotes>`,
            'word/_rels/footnotes.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.com/note" TargetMode="External"/></Relationships>`,
            'word/endnotes.xml': `<w:endnotes ${W}><w:endnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> italic endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`,
            'word/comments.xml': `<w:comments ${W}><w:comment w:id="5" w:author="Ann" w:date="2024-03-04T05:06:07Z" w:initials="A">` +
                `<w:p><w:r><w:t xml:space="preserve">first </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold remark</w:t></w:r></w:p>${p('second paragraph')}</w:comment></w:comments>`,
        },
        expect: [
            { part: 'word/footnotes.xml', has: '<w:footnote w:type="separator" w:id="-1">.*<w:separator/>' },
            { part: 'word/footnotes.xml', has: '<w:footnote w:id="2"><w:p><w:pPr/?>(.*?</w:pPr>)?<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' },
            { part: 'word/footnotes.xml', has: '<w:b/>.*the bold source' }, { part: 'word/footnotes.xml', has: '<w:hyperlink r:id="[^"]+"[^>]*>.*the site' },
            { part: 'word/footnotes.xml', has: '<w:tbl>.*note cell a.*note cell b.*</w:tbl>.*after the note table' },
            { part: 'word/footnotes.xml', lacks: '<w:footnoteRef/>.*<w:footnoteRef/>' },
            { part: 'word/_rels/footnotes.xml.rels', has: 'Target="https://example.com/note" TargetMode="External"' },
            { part: 'word/endnotes.xml', has: '<w:endnoteRef/>.*<w:i/>.*italic endnote' },
            { part: 'word/comments.xml', has: '<w:comment w:id="5" w:author="Ann" w:date="2024-03-04T05:06:07Z" w:initials="A">' },
            { part: 'word/comments.xml', has: '<w:b/>.*bold remark.*second paragraph' },
            { part: 'word/document.xml', has: '<w:footnoteReference w:id="2"/>' }, { part: 'word/document.xml', has: '<w:commentReference w:id="5"/>' },
            { read: 'pkg.Footnotes.2.Blocks.1.Type', equals: 'Table' }, { read: 'pkg.Footnotes.2.Blocks.0.Runs.3.Hyperlink', equals: 'https://example.com/note' },
            { read: 'pkg.Comments.0.Text', equals: 'first bold remark second paragraph' },
        ],
    },
    {
        // Three sections, each with its own header: a section's references stay in
        // the sectPr of the paragraph that ends it, the final section's in the body sectPr.
        name: 'synthetic/three-sections',
        body:
            p('section one') + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId11"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:pPr></w:p>` +
            p('section two') + p('still section two') + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId12"/><w:footerReference w:type="default" r:id="rId14"/><w:pgSz w:w="15840" w:h="12240"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:start="7"/></w:sectPr></w:pPr></w:p>` +
            p('section three'),
        sect: `<w:sectPr><w:headerReference w:type="default" r:id="rId13"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:start="3"/></w:sectPr>`,
        rels: [['rId11', 'header', 'header1.xml'], ['rId12', 'header', 'header2.xml'], ['rId13', 'header', 'header3.xml'], ['rId14', 'footer', 'footer1.xml']]
            .map(([id, kind, target]) => `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${target}"/>`).join(''),
        extra: {
            'word/header1.xml': `<w:hdr ${W}>${p('header one')}</w:hdr>`, 'word/header2.xml': `<w:hdr ${W}>${p('header two')}</w:hdr>`,
            'word/header3.xml': `<w:hdr ${W}>${p('header three')}</w:hdr>`, 'word/footer1.xml': `<w:ftr ${W}>${p('footer of section two')}</w:ftr>`,
        },
        types: [1, 2, 3].map((n) => `<Override PartName="/word/header${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`).join('') +
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
        expect: [
            { part: 'word/document.xml', has: '<w:sectPr><w:headerReference w:type="default" r:id="rId11"/>(<w:type w:val="nextPage"/>)?<w:pgSz w:w="12240" w:h="15840"/>.*</w:sectPr></w:pPr>.*section two' },
            { part: 'word/document.xml', has: 'still section two.*<w:sectPr><w:headerReference w:type="default" r:id="rId12"/><w:footerReference w:type="default" r:id="rId14"/>(<w:type w:val="nextPage"/>)?<w:pgSz w:w="15840" w:h="12240"/>.*<w:pgNumType w:start="7"/></w:sectPr></w:pPr>' },
            { part: 'word/document.xml', has: 'section three.*</w:p><w:sectPr><w:headerReference w:type="default" r:id="rId13"/><w:pgSz w:w="12240" w:h="15840"/>.*<w:pgNumType w:start="3"/></w:sectPr></w:body>' },
            // Each reference is written once, and the final section does not pick up section two's footer.
            { part: 'word/document.xml', lacks: 'rId13.*rId13' }, { part: 'word/document.xml', lacks: 'rId14.*rId14' }, { part: 'word/document.xml', lacks: 'rId12.*rId12' },
            { part: 'word/document.xml', lacks: '<w:sectPr>.*<w:sectPr>.*<w:sectPr>.*<w:sectPr>' },
            { read: 'pkg.PageSetup.HeaderReferences.0.RelationshipId', equals: 'rId13' }, { read: 'pkg.PageSetup.FooterReferences', equals: null }, { read: 'pkg.PageSetup.PageNumberStart', equals: 3 },
            { read: 'pkg.Paragraphs.1.HeaderReferences.0.RelationshipId', equals: 'rId11' }, { read: 'pkg.Paragraphs.2.HeaderReferences', equals: null },
            { read: 'pkg.Paragraphs.4.HeaderReferences.0.RelationshipId', equals: 'rId12' }, { read: 'pkg.Paragraphs.4.SectionPageNumberStart', equals: 7 },
        ],
    },
    {
        // Shapes are written as DrawingML (wp:anchor / wp:inline + wps:wsp in an
        // mc:AlternateContent with a VML fallback): geometry, rotation, fills, line and
        // wrap / position survive, and each shape stays in the paragraph it is anchored in.
        name: 'synthetic/shapes',
        body:
            p('before the shapes') +
            `<w:p><w:r><w:t>anchor paragraph</w:t></w:r>` +
            shapeAnchor(1, 'Oval', { relH: 'page', relV: 'page', x: 571500, y: 952500, cx: 1333500, cy: 857250, wrap: '<wp:wrapSquare wrapText="bothSides"/>', behind: 0, z: 5, dist: 114300 },
                `<a:xfrm rot="900000"><a:off x="0" y="0"/><a:ext cx="1333500" cy="857250"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill><a:prstDash val="dash"/></a:ln>`,
                p('in the oval', '', '<w:b/>')) +
            shapeAnchor(2, 'Gradient box', { relH: 'margin', relV: 'paragraph', x: 1905000, y: 0, cx: 952500, cy: 476250, wrap: '<wp:wrapNone/>', behind: 1, z: 3, dist: 0 },
                `<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:ln w="9525"><a:noFill/></a:ln>`, '') +
            `</w:p>` + p('between') +
            `<w:p><w:r><w:t>second anchor</w:t></w:r>` +
            shapeAnchor(3, 'Triangle', { relH: 'column', relV: 'paragraph', x: 95250, y: 190500, cx: 762000, cy: 762000, wrap: '<wp:wrapTopAndBottom/>', behind: 0, z: 9, dist: 0 },
                `<a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="762000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="50" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></a:ln>`, '') +
            `</w:p>` + p('after the shapes'),
        expect: [
            { part: 'word/document.xml', has: 'anchor paragraph</w:t></w:r><w:r><mc:AlternateContent[^>]*><mc:Choice Requires="wps"><w:drawing><wp:anchor ' },
            { part: 'word/document.xml', has: '<wp:positionH relativeFrom="page"><wp:posOffset>571500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>952500</wp:posOffset>' },
            { part: 'word/document.xml', has: '<wp:extent cx="1333500" cy="857250"/>.*?<wp:wrapSquare wrapText="bothSides"/>' },
            { part: 'word/document.xml', has: '<a:xfrm rot="900000">.*?<a:prstGeom prst="ellipse">.*?<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill><a:prstDash val="dash"/></a:ln>' },
            { part: 'word/document.xml', has: '<wps:txbx><w:txbxContent>.*?<w:b/>.*?in the oval' },
            { part: 'word/document.xml', has: 'behindDoc="1".*?<a:prstGeom prst="roundRect">.*?<a:gradFill[^>]*><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:ln w="9525"><a:noFill/>' },
            { part: 'word/document.xml', has: 'second anchor</w:t></w:r><w:r><mc:AlternateContent.*?<wp:wrapTopAndBottom/>.*?<a:custGeom>.*?<a:path w="100" h="100"><a:moveTo><a:pt x="50" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path>.*?<a:noFill/><a:ln w="12700">' },
            { part: 'word/document.xml', has: '<mc:Fallback><w:pict><v:shape ' }, { part: 'word/document.xml', has: 'relativeHeight="9"' },
            // Three shapes in, three shapes out: the VML fallback is not read as a second copy.
            { read: 'pkg.Paragraphs.length', equals: 8 }, { read: 'pkg.Paragraphs.5.AnchorBlock', equals: 1 }, { read: 'pkg.Paragraphs.7.AnchorBlock', equals: 3 },
            { read: 'pkg.Paragraphs.5.Rotation', equals: 15 }, { read: 'pkg.Paragraphs.6.Gradient.Angle', equals: 90 }, { read: 'pkg.Paragraphs.7.NoFill', equals: true },
        ],
    },
    {
        // A floating picture: crop, wrap mode and side, alignment-based position,
        // relative-from, behind-text and z-order survive.
        name: 'synthetic/floating-image',
        body:
            `<w:p><w:r><w:t>text</w:t></w:r><w:r><w:drawing><wp:anchor distT="45720" distB="91440" distL="114300" distR="228600" simplePos="0" relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="0">` +
            `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>` +
            `<wp:extent cx="1371600" cy="685800"/><wp:wrapTight wrapText="left"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>` +
            `<wp:docPr id="7" name="Picture 7" descr="alt text" title="the title"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture 7"/><pic:cNvPicPr/></pic:nvPicPr>` +
            `<pic:blipFill><a:blip r:embed="rId1"/><a:srcRect l="10000" t="5000" r="20000" b="15000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1371600" cy="685800"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>` +
            `<w:p>${drawing('rId1', 457200, 457200)}<w:r><w:t>inline picture</w:t></w:r></w:p>`,
        rels: `<Relationship Id="rId1" Type="${REL}/image" Target="media/body.png"/>`,
        extra: { 'word/media/body.png': PNG },
        types: '<Default Extension="png" ContentType="image/png"/>',
        expect: [
            { part: 'word/document.xml', has: '<wp:anchor [^>]*distT="45720" distB="91440" distL="114300" distR="228600"[^>]*relativeHeight="251659264" behindDoc="1"[^>]*allowOverlap="0"' },
            { part: 'word/document.xml', has: '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset>' },
            { part: 'word/document.xml', has: '<wp:wrapTight wrapText="left">' }, { part: 'word/document.xml', has: '<a:srcRect l="10000" t="5000" r="20000" b="15000"/>' },
            { part: 'word/document.xml', has: 'descr="alt text" title="the title"' }, { part: 'word/document.xml', has: '<wp:inline[^>]*><wp:extent cx="457200" cy="457200"/>' },
        ],
    },
];

export async function syntheticFixtures(JSZip) {
    const out = [];
    for (const doc of docs) {
        const zip = new JSZip();
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${doc.types ?? ''}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`);
        zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
        zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId101" Type="${REL}/numbering" Target="numbering.xml"/>${doc.rels ?? ''}</Relationships>`);
        zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${styles}`);
        zip.file('word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${numbering}`);
        zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${doc.body}${doc.sect ?? sect}</w:body></w:document>`);
        for (const [name, data] of Object.entries(doc.extra ?? {})) zip.file(name, data);
        out.push({ name: doc.name + '.docx', data: await zip.generateAsync({ type: 'nodebuffer' }), expect: doc.expect });
    }
    return out;
}
