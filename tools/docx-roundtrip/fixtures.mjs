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
