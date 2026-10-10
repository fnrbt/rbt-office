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

// A package with every kind of part the editor does not model.
function preservedParts(name, expect) {
    const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
    return {
        name,
        body: p('styled paragraph', '<w:pStyle w:val="Kept"/>') + p('plain paragraph'),
        rels: `<Relationship Id="rId40" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rId41" Type="${REL}/theme" Target="theme/theme1.xml"/>` +
            `<Relationship Id="rId42" Type="${REL}/fontTable" Target="fontTable.xml"/><Relationship Id="rId43" Type="${REL}/webSettings" Target="webSettings.xml"/>` +
            `<Relationship Id="rId44" Type="${REL}/customXml" Target="../customXml/item1.xml"/>`,
        rootRels: `<Relationship Id="rId2" Type="${PKG}/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
            `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>`,
        extra: {
            'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings ${W}><w:zoom w:percent="120"/><w:attachedTemplate r:id="rId1"/><w:trackRevisions/><w:defaultTabStop w:val="708"/><w:footnotePr><w:footnote w:id="-1"/></w:footnotePr>` +
                `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:rsids><w:rsidRoot w:val="00A1B2C3"/></w:rsids><w:themeFontLang w:val="en-US"/></w:settings>`,
            'word/_rels/settings.xml.rels': `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/attachedTemplate" Target="file:///C:/Templates/Normal.dotm" TargetMode="External"/></Relationships>`,
            'word/theme/theme1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Kept theme"><a:themeElements><a:clrScheme name="Kept"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1></a:clrScheme><a:fontScheme name="Kept"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
            'word/fontTable.xml': `<w:fonts ${W}><w:font w:name="Verdana"><w:panose1 w:val="020B0604030504040204"/></w:font></w:fonts>`,
            'word/webSettings.xml': `<w:webSettings ${W}><w:optimizeForBrowser/></w:webSettings>`,
            'customXml/item1.xml': '<root xmlns="urn:example:data"><value>42</value></root>',
            'customXml/itemProps1.xml': '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{11111111-2222-3333-4444-555555555555}"/>',
            'customXml/_rels/item1.xml.rels': `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
            'docProps/core.xml': '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Kept title</dc:title><dc:creator>Ann Author</dc:creator></cp:coreProperties>',
            'docProps/app.xml': '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Some Writer</Application><Company>Example</Company></Properties>',
            'docProps/custom.xml': '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Case"><vt:lpwstr>A-17</vt:lpwstr></property></Properties>',
            'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault></w:docDefaults>` +
                `<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:count="2"><w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/></w:latentStyles>` +
                `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
                `<w:style w:type="paragraph" w:styleId="Kept"><w:name w:val="Kept"/><w:basedOn w:val="Normal"/><w:pPr><w:suppressAutoHyphens/><w:spacing w:after="120"/></w:pPr><w:rPr><w:i/><w:lang w:val="de-DE"/></w:rPr></w:style>` +
                `<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="Numbered"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>` +
                `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style></w:styles>`,
        },
        types: '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
            '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
            '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' +
            '<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>' +
            '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
            '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>',
        expect,
    };
}

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
            p('section two') + p('still section two') + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId12"/><w:footerReference w:type="default" r:id="rId14"/><w:pgSz w:w="15840" w:h="12240"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="7"/></w:sectPr></w:pPr></w:p>` +
            p('section three'),
        sect: `<w:sectPr><w:headerReference w:type="default" r:id="rId13"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:pgNumType w:fmt="upperLetter" w:start="3"/></w:sectPr>`,
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
            { part: 'word/document.xml', has: 'still section two.*<w:sectPr><w:headerReference w:type="default" r:id="rId12"/><w:footerReference w:type="default" r:id="rId14"/>(<w:type w:val="nextPage"/>)?<w:pgSz w:w="15840" w:h="12240"/>.*<w:pgNumType w:fmt="lowerRoman" w:start="7"/></w:sectPr></w:pPr>' },
            { part: 'word/document.xml', has: 'section three.*</w:p><w:sectPr><w:headerReference w:type="default" r:id="rId13"/><w:pgSz w:w="12240" w:h="15840"/>.*<w:pgNumType w:fmt="upperLetter" w:start="3"/></w:sectPr></w:body>' },
            // Each reference is written once, and the final section does not pick up section two's footer.
            { part: 'word/document.xml', lacks: 'rId13.*rId13' }, { part: 'word/document.xml', lacks: 'rId14.*rId14' }, { part: 'word/document.xml', lacks: 'rId12.*rId12' },
            { part: 'word/document.xml', lacks: '<w:sectPr>.*<w:sectPr>.*<w:sectPr>.*<w:sectPr>' },
            { read: 'pkg.PageSetup.HeaderReferences.0.RelationshipId', equals: 'rId13' }, { read: 'pkg.PageSetup.FooterReferences', equals: null }, { read: 'pkg.PageSetup.PageNumberStart', equals: 3 }, { read: 'pkg.PageSetup.PageNumberFormat', equals: 'upperLetter' },
            { read: 'pkg.Paragraphs.1.HeaderReferences.0.RelationshipId', equals: 'rId11' }, { read: 'pkg.Paragraphs.2.HeaderReferences', equals: null },
            { read: 'pkg.Paragraphs.4.HeaderReferences.0.RelationshipId', equals: 'rId12' }, { read: 'pkg.Paragraphs.4.SectionPageNumberStart', equals: 7 }, { read: 'pkg.Paragraphs.4.SectionPageNumberFormat', equals: 'lowerRoman' },
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
            `<wp:extent cx="1371600" cy="685800"/><wp:wrapTight wrapText="left"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="10800"/><wp:lineTo x="10800" y="-300"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>` +
            `<wp:docPr id="7" name="Picture 7" descr="alt text" title="the title"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture 7"/><pic:cNvPicPr/></pic:nvPicPr>` +
            `<pic:blipFill><a:blip r:embed="rId1"/><a:srcRect l="10000" t="5000" r="20000" b="15000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1371600" cy="685800"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>` +
            `<w:p>${drawing('rId1', 457200, 457200)}<w:r><w:t>inline picture</w:t></w:r></w:p>`,
        rels: `<Relationship Id="rId1" Type="${REL}/image" Target="media/body.png"/>`,
        extra: { 'word/media/body.png': PNG },
        types: '<Default Extension="png" ContentType="image/png"/>',
        expect: [
            { part: 'word/document.xml', has: '<wp:anchor [^>]*distT="45720" distB="91440" distL="114300" distR="228600"[^>]*relativeHeight="251659264" behindDoc="1"[^>]*allowOverlap="0"' },
            { part: 'word/document.xml', has: '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset>' },
            { part: 'word/document.xml', has: '<wp:wrapTight wrapText="left"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="10800"/><wp:lineTo x="10800" y="-300"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>' }, { part: 'word/document.xml', has: '<a:srcRect l="10000" t="5000" r="20000" b="15000"/>' },
            { part: 'word/document.xml', has: 'descr="alt text" title="the title"' }, { part: 'word/document.xml', has: '<wp:inline[^>]*><wp:extent cx="457200" cy="457200"/>' },
        ],
    },
    {
        // Objects the editor does not model are kept as they are: the drawing XML in
        // the text, and the parts behind it (chart + embedded workbook, diagram parts,
        // OLE binary and its picture) with their relationships and content types.
        name: 'synthetic/opaque-objects',
        root: 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" mc:Ignorable="wp14"',
        body:
            `<w:p><w:r><w:t xml:space="preserve">chart: </w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="12345678"><wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="Chart 1"/><wp:cNvGraphicFramePr/>` +
            `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId30"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
            `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="2" name="Diagram 2"/><wp:cNvGraphicFramePr/>` +
            `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rId31" r:lo="rId32" r:qs="rId33" r:cs="rId34"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
            `<w:p><w:r><w:object w:dxaOrig="1440" w:dyaOrig="720"><v:shape id="_x0000_i1025" style="width:72pt;height:36pt" o:ole=""><v:imagedata r:id="rId35" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="rId36"/></w:object></w:r></w:p>` +
            `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="cx1"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="3" name="Chart 3"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx1:chart r:id="rId37"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>` +
            `<mc:Fallback>${drawing('rId38', 914400, 914400).replace('<w:r>', '').replace('</w:r>', '')}</mc:Fallback></mc:AlternateContent></w:r></w:p>` +
            `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="4" name="Group 4"/>` +
            `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/><wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp><wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
            `<mc:Fallback><w:pict><v:group style="position:absolute;width:72pt;height:36pt"><v:rect style="width:36pt;height:36pt"/><v:oval style="width:36pt;height:36pt"/></v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>after the group</w:t></w:r></w:p>` +
            p('end'),
        sect: `<w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`,
        rels:
            `<Relationship Id="rId30" Type="${REL}/chart" Target="charts/chart1.xml"/>` +
            ['Data', 'Layout', 'QuickStyle', 'Colors'].map((kind, i) => `<Relationship Id="rId${31 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagram${kind}" Target="diagrams/${kind.toLowerCase()}1.xml"/>`).join('') +
            `<Relationship Id="rId35" Type="${REL}/image" Target="media/ole1.png"/><Relationship Id="rId36" Type="${REL}/package" Target="embeddings/Sheet1.xlsx"/>` +
            `<Relationship Id="rId37" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="charts/chartEx1.xml"/><Relationship Id="rId38" Type="${REL}/image" Target="media/fallback.png"/>` +
            `<Relationship Id="rId9" Type="${REL}/header" Target="header1.xml"/>`,
        extra: {
            'word/charts/chart1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/></c:barChart></c:plotArea></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>',
            'word/charts/_rels/chart1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/package" Target="../embeddings/Workbook1.xlsx"/></Relationships>`,
            'word/embeddings/Workbook1.xlsx': Buffer.from([80, 75, 3, 4, 0, 1, 2, 3, 250, 251, 252, 253, 254, 255]),
            'word/embeddings/Sheet1.xlsx': Buffer.from([80, 75, 5, 6, 9, 8, 7]),
            'word/charts/chartEx1.xml': '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart/></cx:chartSpace>',
            'word/diagrams/data1.xml': '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/></dgm:dataModel>',
            'word/diagrams/layout1.xml': '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
            'word/diagrams/quickstyle1.xml': '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
            'word/diagrams/colors1.xml': '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
            'word/media/ole1.png': PNG, 'word/media/fallback.png': PNG2,
            // A WordArt watermark: VML in the header, with nothing the reader models.
            'word/header1.xml': `<w:hdr ${W} xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251658752" fillcolor="silver" stroked="f"><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r><w:r><w:t>header text</w:t></w:r></w:p></w:hdr>`,
        },
        types: '<Default Extension="png" ContentType="image/png"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>' +
            '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
            '<Override PartName="/word/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>' +
            '<Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>' +
            '<Override PartName="/word/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>' +
            '<Override PartName="/word/diagrams/quickstyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>' +
            '<Override PartName="/word/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>' +
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
        expect: [
            { part: 'word/document.xml', has: 'chart: </w:t></w:r><w:r><w:drawing[^>]*><wp:inline [^>]*wp14:anchorId="12345678"[^>]*><wp:extent cx="2743200" cy="1828800"/>.*?<c:chart r:id="rId30"/>' },
            { part: 'word/document.xml', has: '<dgm:relIds r:cs="rId34" r:dm="rId31" r:lo="rId32" r:qs="rId33"/>' },
            { part: 'word/document.xml', has: '<w:object [^>]*>.*?<v:imagedata [^>]*r:id="rId35"[^>]*/>.*?<o:OLEObject [^>]*ProgID="Excel.Sheet.12"[^>]*r:id="rId36"' },
            { part: 'word/document.xml', has: '<mc:Choice Requires="cx1"><w:drawing>.*?<cx1:chart r:id="rId37"/>.*?<mc:Fallback>.*?r:embed="rId38"' },
            { part: 'word/document.xml', has: '<mc:Choice Requires="wpg">.*?<wpg:wgp>.*?</wpg:wgp>.*?<mc:Fallback><w:pict><v:group ' },
            { part: 'word/document.xml', lacks: '>Chart<' }, { part: 'word/document.xml', lacks: 'wps:txbx' },
            { part: 'word/_rels/document.xml.rels', has: '<Relationship Id="rId30" Type="[^"]*/chart" Target="/word/charts/chart1.xml"/>' },
            { part: 'word/_rels/document.xml.rels', has: 'Id="rId31" Type="[^"]*/diagramData" Target="/word/diagrams/data1.xml"' },
            { part: 'word/_rels/document.xml.rels', has: 'Id="rId36" Type="[^"]*/package" Target="/word/embeddings/Sheet1.xlsx"' },
            { part: 'word/_rels/document.xml.rels', has: 'Id="rId38" Type="[^"]*/image" Target="/word/media/fallback.png"' },
            { part: '[Content_Types].xml', has: 'PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart\\+xml"' },
            { part: '[Content_Types].xml', has: 'Extension="xlsx"' }, { part: '[Content_Types].xml', has: 'PartName="/word/diagrams/colors1.xml"' },
            { part: 'word/header1.xml', has: '<w:pict[^>]*><v:shape [^>]*id="PowerPlusWaterMarkObject1"[^>]*><v:textpath [^>]*string="DRAFT"[^>]*/></v:shape></w:pict>' },
            { same: 'word/charts/chart1.xml' }, { same: 'word/charts/_rels/chart1.xml.rels' }, { same: 'word/embeddings/Workbook1.xlsx' }, { same: 'word/embeddings/Sheet1.xlsx' },
            { same: 'word/diagrams/data1.xml' }, { same: 'word/diagrams/colors1.xml' }, { same: 'word/media/ole1.png' }, { same: 'word/media/fallback.png' }, { same: 'word/charts/chartEx1.xml' },
            { read: 'pkg.Paragraphs.0.Runs.1.OpaqueKind', equals: 'chart' }, { read: 'pkg.Paragraphs.0.Runs.1.ImageSrc', equals: 'opaque:chart' }, { read: 'pkg.Paragraphs.0.Runs.1.ImageWidth', equals: 288 },
            { read: 'pkg.Paragraphs.1.Runs.0.OpaqueKind', equals: 'diagram' }, { read: 'pkg.Paragraphs.2.Runs.0.OpaqueKind', equals: 'ole' },
            // The chart with a fallback picture shows that picture.
            { read: 'pkg.Paragraphs.3.Runs.0.ImageSrc.length', equals: 22 + Math.ceil(PNG2.length / 3) * 4 },
            { read: 'pkg.Paragraphs.4.Runs.0.OpaqueKind', equals: 'drawing' }, { read: 'pkg.Paragraphs.length', equals: 6 },
            { read: 'pkg.HeaderParts.0.Paragraphs.0.Runs.0.OpaqueKind', equals: 'shape' },
        ],
    },
    {
        // Empty and whitespace-only paragraphs at the very end of the body are
        // content: they come through import and export, and through a save.
        name: 'synthetic/trailing-empty-paragraphs',
        body: p('first') + '<w:p/>' + p('last text') + '<w:p/>' + '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>' + p('   ') + '<w:p/>',
        expect: [
            { read: 'pkg.Paragraphs.length', equals: 7 },
            { part: 'word/document.xml', has: 'last text</w:t></w:r></w:p><w:p>.*<w:jc w:val="center"/>.*<w:t xml:space="preserve">   </w:t>.*<w:p>(<w:pPr/>)?</w:p><w:sectPr' },
        ],
    },
    {
        // Formatting a paragraph inherits (from its style, the document defaults or its
        // list level) or has nowhere at all is not written onto it: only what its own
        // w:pPr states. The first save is a fixed point for unset and for zero.
        name: 'synthetic/unset-spacing',
        body:
            p('nothing set') +
            p('indent only', '<w:ind w:left="720"/>') +
            p('inherits from its style', '<w:pStyle w:val="Spaced"/>') +
            p('explicit zero under the style', '<w:pStyle w:val="Spaced"/><w:spacing w:before="0" w:after="0"/>') +
            `<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>` +
            li('list item, indent from its level', 7) + li('second item', 7) +
            p('after'),
        extra: {
            'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
                `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
                `<w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/><w:tabs><w:tab w:val="right" w:pos="5000"/></w:tabs><w:spacing w:before="100" w:after="200" w:line="360" w:lineRule="auto"/><w:ind w:left="360" w:hanging="180"/><w:contextualSpacing/></w:pPr></w:style></w:styles>`,
            'word/numbering.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}><w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num></w:numbering>`,
        },
        expect: [
            { part: 'word/document.xml', has: '<w:p><w:pPr/><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>nothing set' },
            { part: 'word/document.xml', has: '<w:pPr><w:ind w:left="720"/></w:pPr><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>indent only' },
            { part: 'word/document.xml', has: '<w:pPr><w:pStyle w:val="Spaced"/></w:pPr><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>inherits from its style' },
            { part: 'word/document.xml', has: '<w:pStyle w:val="Spaced"/><w:spacing w:before="0" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0"/></w:pPr>' },
            { part: 'word/document.xml', has: '<w:p><w:pPr><w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr></w:p>' },
            // One w:spacing (the explicit zero), one w:ind (the indent-only paragraph), no shading / tabs / contextual spacing copied from the style.
            { part: 'word/document.xml', lacks: '<w:spacing .*<w:spacing ' }, { part: 'word/document.xml', lacks: '<w:ind .*<w:ind ' },
            { part: 'word/document.xml', lacks: '<w:shd ' }, { part: 'word/document.xml', lacks: '<w:tabs>' }, { part: 'word/document.xml', lacks: 'contextualSpacing' },
            { part: 'word/numbering.xml', has: '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0">.*?<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' },
            { read: 'pkg.Paragraphs.0.SpacingAfter', equals: null }, { read: 'pkg.Paragraphs.0.ExplicitSpacingAfter', equals: null },
            { read: 'pkg.Paragraphs.1.SpacingAfter', equals: 0 }, { read: 'pkg.Paragraphs.1.ExplicitSpacingAfter', equals: null }, { read: 'pkg.Paragraphs.1.ExplicitIndentLeft', equals: 48 },
            { read: 'pkg.Paragraphs.2.SpacingAfter', equals: 10 }, { read: 'pkg.Paragraphs.2.ExplicitSpacingAfter', equals: null }, { read: 'pkg.Paragraphs.2.LineSpacing', equals: 1.5 },
            { read: 'pkg.Paragraphs.2.FirstLineIndent', equals: -12 }, { read: 'pkg.Paragraphs.2.ExplicitFirstLineIndent', equals: null },
            { read: 'pkg.Paragraphs.3.ExplicitSpacingAfter', equals: 0 }, { read: 'pkg.Paragraphs.3.ExplicitSpacingBefore', equals: 0 },
            { read: 'pkg.Paragraphs.4.HasMarkProperties', equals: true }, { read: 'pkg.Paragraphs.4.ExplicitMarkFontSize', equals: 18.667 },
            { read: 'pkg.Paragraphs.5.IndentLeft', equals: 48 }, { read: 'pkg.Paragraphs.5.ExplicitIndentLeft', equals: null }, { read: 'pkg.Paragraphs.5.ListLevelFirstLineIndent', equals: -24 },
        ],
    },
    {
        // Tracked changes as Word writes them: a deletion inside another author's
        // insertion, tracked paragraph marks, and a move with its range markers.
        name: 'synthetic/nested-revisions',
        body:
            `<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:ins w:id="10" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t xml:space="preserve">added by Ann </w:t></w:r>` +
            `<w:del w:id="11" w:author="Bob" w:date="2024-02-03T04:05:06Z"><w:r><w:delText xml:space="preserve">then cut by Bob </w:delText></w:r></w:del><w:r><w:t>still added</w:t></w:r></w:ins></w:p>` +
            `<w:p><w:pPr><w:rPr><w:ins w:id="12" w:author="Ann" w:date="2024-01-02T03:04:05Z"/></w:rPr></w:pPr><w:r><w:t>split here by Enter</w:t></w:r></w:p>` +
            `<w:p><w:pPr><w:rPr><w:del w:id="13" w:author="Bob" w:date="2024-02-03T04:05:06Z"/></w:rPr></w:pPr><w:r><w:t>joined with the next</w:t></w:r></w:p>` +
            `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r><w:moveFromRangeStart w:id="14" w:author="Cy" w:date="2024-03-04T05:06:07Z" w:name="move1"/><w:moveFrom w:id="15" w:author="Cy" w:date="2024-03-04T05:06:07Z"><w:r><w:t xml:space="preserve">the moved words </w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="14"/><w:r><w:t>after</w:t></w:r></w:p>` +
            `<w:p><w:r><w:t xml:space="preserve">destination: </w:t></w:r><w:moveToRangeStart w:id="16" w:author="Cy" w:date="2024-03-04T05:06:07Z" w:name="move1"/><w:moveTo w:id="17" w:author="Cy" w:date="2024-03-04T05:06:07Z"><w:r><w:t xml:space="preserve">the moved words </w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="16"/></w:p>` +
            p('end'),
        expect: [
            { part: 'word/document.xml', has: '<w:ins w:id="10" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>added by Ann </w:t></w:r></w:ins>' },
            { part: 'word/document.xml', has: '<w:ins w:id="10" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:del w:id="11" w:author="Bob" w:date="2024-02-03T04:05:06Z"><w:r>(<w:rPr>.*?</w:rPr>)?<w:delText[^>]*>then cut by Bob </w:delText></w:r></w:del></w:ins>' },
            { part: 'word/document.xml', has: '<w:pPr><w:rPr><w:ins w:id="12" w:author="Ann" w:date="2024-01-02T03:04:05Z"/></w:rPr></w:pPr><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>split here by Enter' },
            { part: 'word/document.xml', has: '<w:pPr><w:rPr><w:del w:id="13" w:author="Bob" w:date="2024-02-03T04:05:06Z"/></w:rPr></w:pPr>' },
            { part: 'word/document.xml', has: '<w:moveFromRangeStart w:id="(\\d+)" w:author="Cy" w:date="2024-03-04T05:06:07Z" w:name="move1"/><w:moveFrom w:id="15" w:author="Cy" w:date="2024-03-04T05:06:07Z"><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>the moved words </w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="\\1"/>' },
            { part: 'word/document.xml', has: '<w:moveToRangeStart w:id="(\\d+)" [^>]*w:name="move1"/><w:moveTo w:id="17" [^>]*><w:r>(<w:rPr>.*?</w:rPr>)?<w:t[^>]*>the moved words </w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="\\1"/>' },
            // The moved text is not turned into a deletion and an insertion.
            { part: 'word/document.xml', lacks: '<w:delText[^>]*>the moved words' }, { part: 'word/document.xml', lacks: '<w:ins [^>]*><w:r>(<w:rPr>((?!</w:r>).)*?</w:rPr>)?<w:t[^>]*>the moved words' },
            { read: 'pkg.Paragraphs.0.Runs.2.RevisionType', equals: 'delete' }, { read: 'pkg.Paragraphs.0.Runs.2.RevisionAuthor', equals: 'Bob' }, { read: 'pkg.Paragraphs.0.Runs.2.InsertedRevisionAuthor', equals: 'Ann' },
            { read: 'pkg.Paragraphs.0.Runs.1.RevisionType', equals: 'insert' }, { read: 'pkg.Paragraphs.0.Runs.1.InsertedRevisionAuthor', equals: null },
            { read: 'pkg.Paragraphs.1.MarkRevisionType', equals: 'insert' }, { read: 'pkg.Paragraphs.2.MarkRevisionType', equals: 'delete' }, { read: 'pkg.Paragraphs.2.MarkRevisionAuthor', equals: 'Bob' },
            { read: 'pkg.Paragraphs.3.Runs.1.RevisionType', equals: 'moveFrom' }, { read: 'pkg.Paragraphs.3.Runs.1.RevisionMoveName', equals: 'move1' },
            { read: 'pkg.Paragraphs.4.Runs.1.RevisionType', equals: 'moveTo' }, { read: 'pkg.Paragraphs.4.Runs.1.RevisionMoveName', equals: 'move1' },
        ],
    },
    {
        // Hyperlink targets a document may carry. Only http, https, mailto, tel and
        // anchors in the document are links; anything else is kept (the file saves as
        // it was) but is inactive. A linked picture keeps its address and is marked linked.
        name: 'synthetic/hostile-links',
        body:
            `<w:p>` + ['rId50', 'rId51', 'rId52', 'rId53', 'rId54', 'rId55', 'rId56', 'rId57'].map((id, i) => `<w:hyperlink r:id="${id}"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">link${i} </w:t></w:r></w:hyperlink>`).join('') +
            `<w:hyperlink w:anchor="chapter2"><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>to the bookmark</w:t></w:r></w:hyperlink></w:p>` +
            `<w:p><w:bookmarkStart w:id="1" w:name="chapter2"/><w:r><w:t>Chapter 2</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>` +
            `<w:p>${drawing('rId58', 457200, 457200).replace('r:embed', 'r:link')}${drawing('rId59', 457200, 457200)}</w:p>`,
        rels: [['rId50', 'javascript:alert(1)'], ['rId51', 'https://example.com/ok'], ['rId52', 'mailto:someone@example.com'], ['rId53', 'tel:+15551234'],
               ['rId54', 'file:///etc/passwd'], ['rId55', 'data:text/html,&lt;script&gt;alert(1)&lt;/script&gt;'], ['rId56', ' JavaScript:alert(2)'], ['rId57', 'vbscript:msgbox(1)']]
            .map(([id, target]) => `<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${target}" TargetMode="External"/>`).join('') +
            `<Relationship Id="rId58" Type="${REL}/image" Target="https://tracker.example.com/pixel.png" TargetMode="External"/>` +
            `<Relationship Id="rId59" Type="${REL}/image" Target="http://tracker.example.com/embedded-but-external.png" TargetMode="External"/>`,
        expect: [
            { read: 'pkg.Paragraphs.0.Runs.0.Hyperlink', equals: null }, { read: 'pkg.Paragraphs.0.Runs.0.InactiveHyperlink', equals: 'javascript:alert(1)' },
            { read: 'pkg.Paragraphs.0.Runs.1.Hyperlink', equals: 'https://example.com/ok' }, { read: 'pkg.Paragraphs.0.Runs.1.InactiveHyperlink', equals: null },
            { read: 'pkg.Paragraphs.0.Runs.2.Hyperlink', equals: 'mailto:someone@example.com' }, { read: 'pkg.Paragraphs.0.Runs.3.Hyperlink', equals: 'tel:+15551234' },
            { read: 'pkg.Paragraphs.0.Runs.4.Hyperlink', equals: null }, { read: 'pkg.Paragraphs.0.Runs.4.InactiveHyperlink', equals: 'file:///etc/passwd' },
            { read: 'pkg.Paragraphs.0.Runs.5.Hyperlink', equals: null }, { read: 'pkg.Paragraphs.0.Runs.6.Hyperlink', equals: null }, { read: 'pkg.Paragraphs.0.Runs.6.InactiveHyperlink', equals: ' JavaScript:alert(2)' },
            { read: 'pkg.Paragraphs.0.Runs.7.Hyperlink', equals: null }, { read: 'pkg.Paragraphs.0.Runs.8.Hyperlink', equals: '#chapter2' },
            { read: 'pkg.Paragraphs.2.Runs.0.ImageLinked', equals: true }, { read: 'pkg.Paragraphs.2.Runs.1.ImageLinked', equals: true },
            // The targets are still in the saved file, where they were.
            { part: 'word/_rels/document.xml.rels', has: 'Type="[^"]*/hyperlink" Target="javascript:alert\\(1\\)" TargetMode="External"' },
            { part: 'word/_rels/document.xml.rels', has: 'Target="file:///etc/passwd" TargetMode="External"' }, { part: 'word/_rels/document.xml.rels', has: 'Target="https://example.com/ok"' },
            { part: 'word/_rels/document.xml.rels', has: 'Type="[^"]*/image" Target="https://tracker.example.com/pixel.png" TargetMode="External"' },
            { part: 'word/document.xml', has: '<w:hyperlink w:anchor="chapter2" w:history="1">' }, { part: 'word/document.xml', has: '<a:blip r:link="[^"]+"/>.*<a:blip r:link="[^"]+"/>' },
        ],
    },
    preservedParts('synthetic/preserved-parts', [
        // Everything the editor does not model comes back as it was.
        { same: 'word/theme/theme1.xml' }, { same: 'word/fontTable.xml' }, { same: 'word/webSettings.xml' },
        { same: 'customXml/item1.xml' }, { same: 'customXml/itemProps1.xml' }, { same: 'customXml/_rels/item1.xml.rels' },
        { same: 'docProps/core.xml' }, { same: 'docProps/app.xml' }, { same: 'docProps/custom.xml' },
        { part: '_rels/.rels', has: 'Id="rId2" Type="[^"]*/core-properties" Target="/docProps/core.xml"' }, { part: '_rels/.rels', has: 'custom-properties' },
        { part: 'word/_rels/document.xml.rels', has: 'Type="[^"]*/theme" Target="/word/theme/theme1.xml"' }, { part: 'word/_rels/document.xml.rels', has: 'Type="[^"]*/customXml" Target="/customXml/item1.xml"' },
        { part: '[Content_Types].xml', has: 'PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme\\+xml"' },
        { part: '[Content_Types].xml', has: 'PartName="/docProps/custom.xml"' },
        // settings.xml: compatibility, default tab stop, track-revisions flag, footnote properties, attached template.
        { part: 'word/settings.xml', has: '<w:trackRevisions/><w:defaultTabStop w:val="708"/>' }, { part: 'word/settings.xml', has: '<w:footnotePr><w:footnote w:id="-1"/></w:footnotePr>' },
        { part: 'word/settings.xml', has: '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' },
        { part: 'word/settings.xml', has: '<w:attachedTemplate r:id="rId1"/>' }, { part: 'word/settings.xml', lacks: 'evenAndOddHeaders' },
        { part: 'word/_rels/settings.xml.rels', has: 'Id="rId1" Type="[^"]*/attachedTemplate" Target="file:///C:/Templates/Normal.dotm" TargetMode="External"' },
        // styles.xml: latent styles, a table style, a numbering-linked style and properties the editor has no model for.
        { part: 'word/styles.xml', has: '<w:latentStyles [^>]*w:count="2"[^>]*><w:lsdException [^>]*w:name="Normal"' },
        { part: 'word/styles.xml', has: '<w:style w:styleId="TableGrid" w:type="table">.*?<w:tblBorders>.*?</w:tblBorders>.*?<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>' },
        { part: 'word/styles.xml', has: '<w:style w:styleId="Numbered" w:type="paragraph">.*?<w:numPr><w:numId w:val="3"/></w:numPr>' },
        { part: 'word/styles.xml', has: '<w:style w:styleId="Kept" w:type="paragraph">.*?<w:suppressAutoHyphens/>.*?<w:lang w:val="de-DE"/>' },
        { part: 'word/styles.xml', has: '<w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/><w:lang w:eastAsia="ja-JP" w:val="en-US"/>' },
        { part: 'word/document.xml', has: '<w:pStyle w:val="Kept"/>' },
    ]),
    preservedParts('synthetic/preserved-parts-edited', [
        // Edits to what the editor owns are merged into the kept parts: even / odd
        // headers go into the original settings (in schema position), a changed style
        // replaces its original definition and the other styles stay untouched.
        { mutate: 'pkg.PageSetup.EvenAndOddHeaders = true; const kept = pkg.Styles.find((s) => s.StyleId === "Kept"); kept.StyleBold = true; kept.StyleFontSize = 32; pkg.ChangedStyles = ["paragraph:Kept"];' },
        { part: 'word/settings.xml', has: '<w:trackRevisions/><w:defaultTabStop w:val="708"/><w:evenAndOddHeaders/><w:footnotePr>' },
        { part: 'word/settings.xml', has: '<w:compat>' },
        { part: 'word/styles.xml', has: '<w:style w:type="paragraph" w:styleId="Kept">.*?<w:b/>.*?<w:sz w:val="48"/>' }, { part: 'word/styles.xml', lacks: 'suppressAutoHyphens' },
        { part: 'word/styles.xml', has: '<w:style w:styleId="TableGrid" w:type="table">.*?<w:tblStylePr w:type="firstRow">' }, { part: 'word/styles.xml', has: '<w:latentStyles ' },
        { part: 'word/styles.xml', has: '<w:style w:styleId="Numbered" w:type="paragraph">.*?<w:numPr><w:numId w:val="3"/></w:numPr>' },
        { same: 'word/theme/theme1.xml' }, { same: 'customXml/item1.xml' }, { same: 'docProps/core.xml' },
    ]),
];

export async function syntheticFixtures(JSZip) {
    const out = [];
    for (const doc of docs) {
        const zip = new JSZip();
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${doc.types ?? ''}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`);
        zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>${doc.rootRels ?? ''}</Relationships>`);
        zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId101" Type="${REL}/numbering" Target="numbering.xml"/>${doc.rels ?? ''}</Relationships>`);
        zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${styles}`);
        zip.file('word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${numbering}`);
        zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}${doc.root ? ' ' + doc.root : ''}><w:body>${doc.body}${doc.sect ?? sect}</w:body></w:document>`);
        for (const [name, data] of Object.entries(doc.extra ?? {})) zip.file(name, data);
        out.push({ name: doc.name + '.docx', data: await zip.generateAsync({ type: 'nodebuffer' }), expect: doc.expect });
    }
    // A document made from scratch (page.js builds it): no kept parts, still a valid package.
    out.push({ name: 'scratch/new-document.docx', data: Buffer.alloc(0), expect: [
        { part: '[Content_Types].xml', has: 'PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main\\+xml"' },
        { part: '_rels/.rels', has: 'Target="word/document.xml"' }, { part: 'word/_rels/document.xml.rels', has: 'Target="styles.xml"' },
        { part: 'word/styles.xml', has: '<w:docDefaults>' }, { part: 'word/document.xml', has: 'New document' }, { savedText: 'typed bold text' },
        { absent: 'word/settings.xml' },
    ] });
    return out;
}
