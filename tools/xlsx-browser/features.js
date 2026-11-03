// Browser side of run.mjs, fourth part: features of the Spreadsheet component that are
// persisted in XLSX the way Excel writes them (dynamic arrays, ...). Each check returns a
// list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, bytesOf } from '/component.js';
import { reopened } from '/sheet-api.js';

const text = async (bytes, path) => { const zip = await window.JSZip.loadAsync(bytes); return zip.files[path] ? zip.files[path].async('string') : null; };
// A package from its parts ({ path: text }).
export const zipOf = async (parts) => {
    const zip = new window.JSZip();
    for (const [path, content] of Object.entries(parts)) zip.file(path, content);
    return zip.generateAsync({ type: 'uint8array' });
};
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
// A one-sheet workbook as Excel lays it out: `sheetBody` is what the worksheet element holds,
// `extra` further parts as { path: [content type, relationship type or null (none of the workbook), text] }.
export const excelPackage = (sheetBody, extra = {}) => {
    const parts = {
        '[Content_Types].xml': DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
            + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            + Object.entries(extra).map(([path, [type]]) => `<Override PartName="/${path}" ContentType="${type}"/>`).join('') + '</Types>',
        '_rels/.rels': DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        'xl/_rels/workbook.xml.rels': DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            + Object.entries(extra).filter(([, [, relType]]) => relType).map(([path, [, relType]], i) => `<Relationship Id="rId${i + 10}" Type="${relType}" Target="${path.replace(/^xl\//, '')}"/>`).join('') + '</Relationships>',
        'xl/worksheets/sheet1.xml': DECL + `<worksheet xmlns="${NS}">${sheetBody}</worksheet>`,
    };
    for (const [path, [, , content]] of Object.entries(extra)) parts[path] = content;
    return zipOf(parts);
};
const cellTags = (sheetXml, cell) => (new RegExp(`<c r="${cell}"[^>]*>(<f[^>]*>)?`).exec(sheetXml) ?? [''])[0];
const METADATA_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const METADATA_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
const EXCEL_METADATA = DECL + `<metadata xmlns="${NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;

// Dynamic arrays are written the way Excel writes them (cm + xl/metadata.xml), legacy arrays stay legacy.
checks['dynamic-arrays'] = async () => {
    const { lines, check } = makeLog();
    try {
        // ---- a workbook made here ----
        const made = mountGrid();
        try {
            made.grid.SetCell(0, 0, '=SEQUENCE(3)'); made.grid.SetCell(0, 2, '=1+1');
            check('the formula spills', made.grid.GetRange('A1:A3').map((row) => row[0].value), [1, 2, 3]);
            const bytes = await made.grid.ExportXlsx();
            const sheet = await text(bytes, 'xl/worksheets/sheet1.xml');
            check('written as a dynamic array', [/ cm="1"/.test(cellTags(sheet, 'A1')), /<f t="array" ref="A1:A3">/.test(cellTags(sheet, 'A1')), / cm=/.test(cellTags(sheet, 'C1'))], [true, true, false]);
            const metadata = await text(bytes, 'xl/metadata.xml');
            check('with the metadata part Excel writes', metadata && metadata.replace(/\r\n/g, '\n'), EXCEL_METADATA.replace(/\r\n/g, '\n'));
            check('its relationship and content type', [((await text(bytes, 'xl/_rels/workbook.xml.rels')).match(/\/sheetMetadata"/g) ?? []).length, ((await text(bytes, '[Content_Types].xml')).match(/PartName="\/xl\/metadata.xml"/g) ?? []).length], [1, 1]);
            const again = await reopened(made.grid);
            try {
                check('opened again: the values', again.grid.GetRange('A1:A3').map((row) => [row[0].formula, row[0].value]), [['=SEQUENCE(3)', 1], [null, 2], [null, 3]]);
                again.grid.EnterText(0, 0, '=SEQUENCE(2)');
                check('edited: the array spills anew', again.grid.GetRange('A1:A3').map((row) => row[0].value), [1, 2, null]);
                again.grid.Undo();
                check('undone', again.grid.GetRange('A1:A3').map((row) => row[0].value), [1, 2, 3]);
                again.grid.Redo();
                check('redone', again.grid.GetRange('A1:A3').map((row) => row[0].value), [1, 2, null]);
                const second = await text(await again.grid.ExportXlsx(), 'xl/worksheets/sheet1.xml');
                check('saved again as a dynamic array of two', [/ cm="1"/.test(cellTags(second, 'A1')), /ref="A1:A2"/.test(cellTags(second, 'A1'))], [true, true]);
            } finally { again.close(); }
        } finally { made.close(); }

        // ---- a file as Excel writes it: dynamic and legacy arrays side by side ----
        const rows = '<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="C1" cm="1"><f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f><v>1</v></c>'
            + '<c r="E1"><f t="array" ref="E1:E3">A1:A3*2</f><v>6</v></c><c r="G1"><f t="array" ref="G1">SUM(A1:A3*A1:A3)</f><v>14</v></c>'
            + '<c r="H1" cm="1"><f t="array" ref="H1">_xlfn.SEQUENCE(1)</f><v>1</v></c></row>'
            + '<row r="2"><c r="A2"><v>1</v></c><c r="C2"><v>2</v></c><c r="E2"><v>2</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="C3"><v>3</v></c><c r="E3"><v>4</v></c></row></sheetData>';
        const file = mountGrid();
        try {
            await file.grid.ImportXlsx(await excelPackage(rows, { 'xl/metadata.xml': [METADATA_TYPE, METADATA_REL, EXCEL_METADATA] }));
            check('the values of the file', [file.grid.GetRange('C1:C3').map((row) => row[0].value), file.grid.GetRange('E1:E3').map((row) => row[0].value), file.grid.GetCell(0, 6).value], [[1, 2, 3], [6, 2, 4], 14]);
            file.grid.SetCell(1, 0, '5');
            check('an edit recalculates both kinds', [file.grid.GetRange('C1:C3').map((row) => row[0].value), file.grid.GetRange('E1:E3').map((row) => row[0].value), file.grid.GetCell(0, 6).value], [[2, 3, 5], [6, 10, 4], 38]);
            const bytes = await file.grid.ExportXlsx();
            const sheet = await text(bytes, 'xl/worksheets/sheet1.xml');
            check('dynamic arrays name the metadata, legacy arrays do not', ['C1', 'E1', 'G1', 'H1'].map((cell) => [/ cm="1"/.test(cellTags(sheet, cell)), (/ref="([^"]*)"/.exec(cellTags(sheet, cell)) ?? [])[1]]), [[true, 'C1:C3'], [false, 'E1:E3'], [false, 'G1'], [true, 'H1']]);
            check("Excel's metadata part is kept as it was", await text(bytes, 'xl/metadata.xml'), EXCEL_METADATA);
            check('one relationship, one content type', [((await text(bytes, 'xl/_rels/workbook.xml.rels')).match(/\/sheetMetadata"/g) ?? []).length, ((await text(bytes, '[Content_Types].xml')).match(/PartName="\/xl\/metadata.xml"/g) ?? []).length], [1, 1]);
        } finally { file.close(); }

        // ---- a legacy file has no metadata after a save ----
        const legacy = mountGrid();
        try {
            await legacy.grid.ImportXlsx(await excelPackage(rows.replace(/ cm="1"/g, '')));
            const bytes = await legacy.grid.ExportXlsx();
            const sheet = await text(bytes, 'xl/worksheets/sheet1.xml');
            check('legacy arrays stay legacy', [(sheet.match(/<f t="array"/g) ?? []).length, / cm=/.test(sheet), await text(bytes, 'xl/metadata.xml')], [4, false, null]);
        } finally { legacy.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    return lines;
};

// ---- x14 conditional formats and validations (Excel 2010+) ----
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
// Two sheets as Excel lays them out: "Data" with `dataBody` in its worksheet element, "Lists" with A1:A3 = red, green, blue and B1 = 3, B2 = 40.
export const twoSheetPackage = (dataBody) => zipOf({
    '[Content_Types].xml': DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        + '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': DECL + `<workbook xmlns="${NS}" xmlns:r="${REL_NS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Lists" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': DECL + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL_NS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': DECL + `<worksheet xmlns="${NS}" xmlns:r="${REL_NS}">${dataBody}</worksheet>`,
    'xl/worksheets/sheet2.xml': DECL + `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>red</t></is></c><c r="B1"><v>3</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>green</t></is></c><c r="B2"><v>40</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>blue</t></is></c></row></sheetData></worksheet>`,
});
const numberRows = (values) => '<sheetData>' + values.map((cells, r) => `<row r="${r + 1}">` + cells.map((v, c) => `<c r="${'ABCDEFGH'[c]}${r + 1}"><v>${v}</v></c>`).join('') + '</row>').join('') + '</sheetData>';
const X14_SHEET = numberRows([[-50, 1, 5, 0, 1], [-25, 2, 4, 0, 2], [0, 3, 3, 0, 3], [50, 4, 2, 0, 4], [100, 5, 1, 0, 5]])
    + `<conditionalFormatting sqref="A1:A5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14}"><x14:id>{C0DA50BC-9888-411D-B3C3-DCE2CC2ECCB4}</x14:id></ext></extLst></cfRule></conditionalFormatting>`
    + `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}"><x14:conditionalFormattings>`
    + `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="dataBar" id="{C0DA50BC-9888-411D-B3C3-DCE2CC2ECCB4}"><x14:dataBar minLength="0" maxLength="100" border="1" gradient="0" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF1F4E79"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FF990000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>`
    + `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="iconSet" priority="2" id="{AAAAAAAA-0000-4000-8000-000000000001}"><x14:iconSet iconSet="3Triangles" showValue="0" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>2</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>4</xm:f></x14:cfvo><x14:cfIcon iconSet="3Arrows" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>C1:C5</xm:sqref></x14:conditionalFormatting>`
    + `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="expression" priority="3" id="{AAAAAAAA-0000-4000-8000-000000000002}"><xm:f>Lists!$B$1&gt;=$E1</xm:f><x14:dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>E1:E5</xm:sqref></x14:conditionalFormatting>`
    + `</x14:conditionalFormattings></ext>`
    + `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="${XM}"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>H1:H5</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`;

checks['x14-rules'] = async () => {
    const { lines, check } = makeLog();
    try {
        const file = mountGrid();
        try {
            const { grid, host } = file;
            await grid.ImportXlsx(await twoSheetPackage(X14_SHEET));
            await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
            const rulesOf = (g) => g.GetConditionalFormats().map((entry) => [entry.range, entry.rule]);
            const expected = [
                ['A1:A5', { type: 'dataBar', min: { type: 'autoMin', value: null }, max: { type: 'autoMax', value: null }, color: '#638EC6', gradient: false, borderColor: '#1F4E79', negativeColor: '#FF0000', negativeBorderColor: '#990000',
                            axisPosition: 'automatic', axisColor: '#000000', direction: 'context', showValue: true, minLength: 0, maxLength: 100, format: null }],
                ['C1:C5', { type: 'iconSet', iconSet: '3Triangles', thresholds: [{ type: 'percent', value: 0 }, { type: 'number', value: 2 }, { type: 'number', value: 4 }], reverse: false, showValue: false,
                            icons: [{ iconSet: '3Arrows', index: 0 }, { iconSet: 'NoIcons', index: 0 }, { iconSet: '3Stars', index: 2 }], format: null }],
            ];
            const read = rulesOf(grid);
            check('the data bar and the icon set of the file, with their x14 options', read.slice(0, 2), expected);
            check('the rule that reads another sheet', [read[2][0], read[2][1].type, read[2][1].formula, read[2][1].format.fill, read[2][1].format.bold], ['E1:E5', 'expression', 'Lists!$B$1>=$E1', '#FFC7CE', true]);

            // ---- the DOM grid ----
            const barOf = (row) => { const bar = node(row, 0).querySelector('.spreadsheet-cf-bar'); return bar && [bar.style.left, bar.style.width, bar.getAttribute('data-cf-bar-color'), bar.style.borderColor !== '']; };
            check('bars: negative left of the axis in red, positive right of it, none for zero', [0, 2, 4].map(barOf), [['0%', '33.33%', '#FF0000', true], null, ['33.33%', '66.67%', '#638EC6', true]]);
            check('bars: the axis in every cell of the rule', [0, 1, 2, 3, 4].map((row) => node(row, 0).querySelector('.spreadsheet-cf-axis')?.style.left), ['33.33%', '33.33%', '33.33%', '33.33%', '33.33%']);
            const box = (element) => { const r = element.getBoundingClientRect(); return [Math.round(r.left), Math.round(r.right)]; };
            const cellBox = box(node(4, 0)), barBox = box(node(4, 0).querySelector('.spreadsheet-cf-bar')), negBox = box(node(0, 0).querySelector('.spreadsheet-cf-bar'));
            check('bars: on screen the positive bar reaches the right edge, the negative bar the left', [Math.abs(barBox[1] - cellBox[1]) <= 2, Math.abs(negBox[0] - box(node(0, 0))[0]) <= 2, Math.abs(negBox[1] - barBox[0]) <= 1], [true, true, true]);
            // Pixels: the middle of the positive bar is its colour.
            const hostBox = host.getBoundingClientRect(), bar = node(4, 0).querySelector('.spreadsheet-cf-bar').getBoundingClientRect();
            const png = await window.screenshotRegion(Math.round(bar.left + bar.width / 2 - 2), Math.round(bar.top + 3), 4, 2);
            const image = new Image(); await new Promise((resolve) => { image.onload = resolve; image.src = 'data:image/png;base64,' + png; });
            const canvas = document.createElement('canvas'); canvas.width = 4; canvas.height = 2; canvas.getContext('2d').drawImage(image, 0, 0);
            const pixel = [...canvas.getContext('2d').getImageData(1, 1, 1, 1).data].slice(0, 3);
            check('bars: the pixels of the positive bar are its colour', pixel, [0x63, 0x8E, 0xC6]);
            const iconOf = (row) => { const icon = node(row, 2).querySelector('[data-cf-icon]'); return icon ? [icon.getAttribute('data-cf-icon-set'), icon.getAttribute('data-cf-icon')] : null; };
            // C1:C5 = 5, 4, 3, 2, 1 with steps at 2 and 4.
            check('icons: the custom icon of each step, none for NoIcons', [0, 2, 4].map(iconOf), [['3Stars', '2'], null, ['3Arrows', '0']]);
            check('icons: the values are hidden', [0, 4].map((row) => /\d/.test(node(row, 2).textContent)), [false, false]);
            const fillOf = (row) => getComputedStyle(node(row, 4)).backgroundColor;
            check('the cross-sheet rule formats E1:E3 (Lists!B1 = 3)', [0, 2, 3].map(fillOf), ['rgb(255, 199, 206)', 'rgb(255, 199, 206)', 'rgb(255, 255, 255)']);
            // The other sheet edited: the rule follows.
            grid.SwitchSheet(1); grid.EnterText(0, 1, '1'); grid.SwitchSheet(0);
            await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            check('... and follows an edit of that sheet (Lists!B1 = 1)', [0, 1, 2].map(fillOf), ['rgb(255, 199, 206)', 'rgb(255, 255, 255)', 'rgb(255, 255, 255)']);
            grid.SwitchSheet(1); grid.Undo(); grid.SwitchSheet(0);
            await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            check('... and its undo', [0, 1, 2].map(fillOf), ['rgb(255, 199, 206)', 'rgb(255, 199, 206)', 'rgb(255, 199, 206)']);
            check('the list of the other sheet', grid.GetValidationList(0, 7), ['red', 'green', 'blue']);
            check('the validation', (({ type, source, range }) => ({ type, source, range }))(grid.GetValidationAt(2, 7)), { type: 'list', source: 'Lists!$A$1:$A$3', range: 'H1:H5' });

            // ---- through the host API, with undo ----
            const id = grid.AddConditionalFormat('B1:B5', { type: 'dataBar', color: '#63C384', gradient: true, axisPosition: 'middle', negativeColor: null, borderColor: '#2E7D32', showValue: false });
            await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            const added = grid.GetConditionalFormats().find((entry) => entry.id === id).rule;
            check('a data bar added with options: the others have their defaults', added, { type: 'dataBar', min: { type: 'min', value: null }, max: { type: 'max', value: null }, color: '#63C384', gradient: true, borderColor: '#2E7D32', negativeColor: null, negativeBorderColor: null,
                axisPosition: 'middle', axisColor: '#000000', direction: 'context', showValue: false, minLength: 0, maxLength: 100, format: null });
            check('... drawn from the middle, as a gradient, without the value', [node(4, 1).querySelector('.spreadsheet-cf-bar')?.style.left, /gradient/.test(node(4, 1).querySelector('.spreadsheet-cf-bar')?.style.background ?? ''), /\d/.test(node(4, 1).textContent)], ['50%', true, false]);
            grid.Undo(); await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            check('undone', [grid.GetConditionalFormats().length, node(4, 1).querySelector('.spreadsheet-cf-bar'), node(4, 1).textContent], [3, null, '5']);
            grid.Redo(); await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            check('redone', [grid.GetConditionalFormats().length, node(4, 1).querySelector('.spreadsheet-cf-bar')?.style.left], [4, '50%']);
            const stars = grid.AddConditionalFormat('D1:D5', { type: 'iconSet', iconSet: '3Stars' });
            check('an icon set of Excel 2010+', grid.GetConditionalFormats().find((entry) => entry.id === stars).rule, { type: 'iconSet', iconSet: '3Stars', thresholds: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }], reverse: false, showValue: true, icons: null, format: null });
            const errors = [];
            for (const rule of [{ type: 'dataBar', axisPosition: 'left' }, { type: 'dataBar', negativeColor: 'red' }, { type: 'iconSet', icons: [{ iconSet: '3Arrows', index: 0 }] }, { type: 'dataBar', min: { type: 'autoMin' } }])
                try { grid.AddConditionalFormat('F1:F5', rule); errors.push(null); } catch (e) { errors.push(String(e.message).slice(0, 60)); }
            check('what is refused', errors.map((message) => message !== null), [true, true, true, true]);
            check('a plain data bar still has its three keys', (() => { const plain = grid.AddConditionalFormat('F1:F5', { type: 'dataBar' }); const rule = grid.GetConditionalFormats().find((entry) => entry.id === plain).rule; grid.RemoveConditionalFormat(plain); return Object.keys(rule); })(), ['type', 'min', 'max', 'color', 'format']);

            // ---- saved and opened again ----
            const before = rulesOf(grid);
            const bytes = await grid.ExportXlsx();
            const sheet = await text(bytes, 'xl/worksheets/sheet1.xml');
            const count = (pattern) => (sheet.match(pattern) ?? []).length;
            check('written: two 2007 data bars naming their x14 twins, the others in the x14 list only', [count(/<cfRule type="dataBar"/g), count(/<x14:id>/g), count(/<x14:cfRule type="dataBar"/g), count(/<x14:cfRule type="iconSet"/g), count(/<x14:cfRule type="expression"/g), count(/<cfRule type="(iconSet|expression)"/g)], [2, 2, 2, 2, 1, 0]);
            check('written: the validation in the x14 list', [count(/<x14:dataValidation /g), count(/<dataValidation /g), /<x14:formula1><xm:f>Lists!\$A\$1:\$A\$3<\/xm:f><\/x14:formula1><xm:sqref>H1:H5<\/xm:sqref>/.test(sheet)], [1, 0, true]);
            const again = await reopened(grid);
            try {
                check('opened again: the same rules', rulesOf(again.grid), before);
                check('opened again: the list', again.grid.GetValidationList(0, 7), ['red', 'green', 'blue']);
            } finally { again.close(); }
        } finally { file.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    return lines;
};

// ---- tables: typing in and around them, the look of the built-in styles ----
checks['table-entry'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    const frames = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    try {
        const formula = (a1) => grid.GetRange(a1)[0][0].formula, value = (a1) => grid.GetRange(a1)[0][0].value;
        const table = (g = grid) => g.GetTables().map((t) => `${t.name}!${t.range} [${t.columns.join('|')}]`);
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const look = (row, col) => { const style = getComputedStyle(node(row, col)); return [style.backgroundColor, style.color, style.fontWeight === '700' || style.fontWeight === 'bold']; };
        [['Region', 'Amount', 'Qty', 'Total'], ['East', '10', '1', '=[@Amount]*[@Qty]'], ['West', '20', '2', '=[@Amount]*[@Qty]'], ['East', '30', '3', '=[@Amount]*[@Qty]']]
            .forEach((row, r) => row.forEach((text, c) => grid.SetCell(r, c, text)));
        grid.SetCell(1, 6, '=SUM(Sales[Amount])'); grid.SetCell(2, 6, '=SUM(Sales[Total])');
        grid.AddSheet('Other'); grid.SetCell(0, 0, '=SUM(Sales[Qty])'); grid.SwitchSheet(0);
        grid.DefineTable('A1:D4', 'Sales', true, 'TableStyleMedium2');
        check('the table', [table(), value('G2'), value('G3'), value('D4')], [['Sales!A1:D4 [Region|Amount|Qty|Total]'], 60, 140, 90]);
        const revision = grid.GetRevision();

        // ---- a header cell names its column ----
        check('typed in a header cell', [grid.EnterText(0, 1, 'Price'), table(), formula('G2'), formula('D2'), value('G2'), value('D4')], [true, ['Sales!A1:D4 [Region|Price|Qty|Total]'], '=SUM(Sales[Price])', '=[@Price]*[@Qty]', 60, 90]);
        check('the revision moved', grid.GetRevision() > revision, true);
        grid.Undo();
        check('one undo: the name, the cell and the formulas', [table(), value('B1'), formula('G2'), formula('D2'), value('G2')], [['Sales!A1:D4 [Region|Amount|Qty|Total]'], 'Amount', '=SUM(Sales[Amount])', '=[@Amount]*[@Qty]', 60]);
        grid.Redo();
        check('redo', [table(), formula('G2'), formula('D3')], [['Sales!A1:D4 [Region|Price|Qty|Total]'], '=SUM(Sales[Price])', '=[@Price]*[@Qty]']);
        grid.Undo();
        check('a name another column has, a blank one', [grid.EnterText(0, 2, 'region'), value('C1'), grid.GetRange('Other!A1')[0][0].formula, (grid.Undo(), grid.EnterText(0, 2, ''), value('C1')), (grid.Undo(), table())],
            [true, 'region2', '=SUM(Sales[region2])', 'Column3', ['Sales!A1:D4 [Region|Amount|Qty|Total]']]);

        // ---- typing below and right of the table ----
        check('typed below the last row: the table grows, the calculated column with it', [grid.EnterText(4, 0, 'North'), table(), formula('D5'), value('D5')], [true, ['Sales!A1:D5 [Region|Amount|Qty|Total]'], '=[@Amount]*[@Qty]', 0]);
        grid.EnterText(4, 1, '40'); grid.EnterText(4, 2, '2');
        check('the new row is the table\'s', [value('D5'), value('G2'), value('G3')], [80, 100, 220]);
        await frames();
        check('... and has the band of the style', [look(4, 0)[0], look(3, 0)[0]], [look(2, 0)[0], look(1, 0)[0]]);
        grid.Undo(); grid.Undo();
        grid.Undo();
        check('one undo takes the row out again', [table(), value('A5'), formula('D5'), value('G2')], [['Sales!A1:D4 [Region|Amount|Qty|Total]'], null, null, 60]);
        grid.SetCell(4, 2, 'note');
        check('not over a cell that holds something', [grid.EnterText(4, 0, 'North'), table(), value('A5')], [true, ['Sales!A1:D4 [Region|Amount|Qty|Total]'], 'North']);
        grid.Undo(); grid.SetCell(4, 2, '');
        check('a formula typed right of the table: a calculated column', [grid.EnterText(1, 4, '=[@Amount]*0.1'), table(), value('E1'), [formula('E3'), value('E3')], value('E4')], [true, ['Sales!A1:E4 [Region|Amount|Qty|Total|Column5]'], 'Column5', ['=[@Amount]*0.1', 2], 3]);
        grid.Undo();
        check('undone in one step', [table(), value('E1'), value('E2'), value('E3')], [['Sales!A1:D4 [Region|Amount|Qty|Total]'], null, null, null]);
        check('a header typed right of the table', [grid.EnterText(0, 4, 'Tax'), table()], [true, ['Sales!A1:E4 [Region|Amount|Qty|Total|Tax]']]);
        grid.EnterText(2, 4, '=[@Total]/10');
        check('a formula in the empty column fills it', ['E2', 'E3', 'E4'].map(value), [1, 4, 9]);

        // ---- the look of the built-in styles ----
        await frames();
        const white = 'rgb(255, 255, 255)', black = 'rgb(0, 0, 0)';
        check('Medium2: header in the accent with white bold text, banded rows', [look(0, 0), look(1, 0)[0], look(2, 0)[0]], [['rgb(68, 114, 196)', white, true], 'rgb(218, 227, 243)', white]);
        grid.SetTableStyle('Sales', 'TableStyleLight1'); await frames();
        // (A cell draws the line above it - the one under the header is the first row's -, a thin one as its grid line.)
        check('Light1: a bold header without a fill, grey bands, lines above and below the header', [look(0, 0), look(1, 0)[0], [node(0, 0), node(1, 0)].every((cell) => cell.querySelector('.spreadsheet-edge') !== null || getComputedStyle(cell).boxShadow.includes('rgb(0, 0, 0)'))], [[white, black, true], 'rgb(217, 217, 217)', true]);
        grid.SetTableStyle('Sales', 'TableStyleDark2'); await frames();
        check('Dark2: dark fills, white text', [look(0, 0), look(1, 0).slice(0, 2), look(2, 0).slice(0, 2)], [[black, white, true], ['rgb(47, 85, 151)', white], ['rgb(68, 114, 196)', white]]);
        grid.SetTableStyle('Sales', { name: 'TableStyleMedium9', showFirstColumn: true }); grid.SetTableTotalRow('Sales', true); await frames();
        check('Medium9: bands in two tints, the first column and the total row in the accent', [look(1, 1)[0], look(2, 1)[0], look(1, 0), look(4, 1)], ['rgb(180, 199, 231)', 'rgb(218, 227, 243)', ['rgb(68, 114, 196)', white, true], ['rgb(68, 114, 196)', white, true]]);
        grid.SetTableTotalRow('Sales', false); grid.SetTableStyle('Sales', 'TableStyleMedium2');

        // ---- saved and opened again ----
        const again = await reopened(grid);
        try {
            check('opened again: the table, its columns, its formulas', [table(again.grid), again.grid.GetRange('E3')[0][0].formula, again.grid.GetRange('E3')[0][0].value, again.grid.GetRange('G2')[0][0].value], [['Sales!A1:E4 [Region|Amount|Qty|Total|Tax]'], '=[@Total]/10', 4, 60]);
            check('... and typing below it still grows it', [again.grid.EnterText(4, 1, '5'), table(again.grid), again.grid.GetRange('E5')[0][0].formula], [true, ['Sales!A1:E5 [Region|Amount|Qty|Total|Tax]'], '=[@Total]/10']);
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// A workbook opens on the sheet it was saved on (workbook.xml's activeTab), and is then
// exported as the file that was opened; ImportXlsx never shows a hidden sheet.
checks['opens-on-saved-sheet'] = async () => {
    const { lines, check } = makeLog();
    for (const [file, saved] of [['tests/real-world/dashboard-ecommerce-sales.xlsx', 2], ['tests/real-world/dcf-stock-analysis.xlsx', 1], ['tests/real-world/milk-street-office.xlsx', 0]]) {
        const { grid, close } = mountGrid();
        try {
            const bytes = await bytesOf(file);
            await grid.ImportXlsx(bytes.slice());
            const label = file.split('/').pop();
            const loaded = grid.GetSheetNames().map((_, i) => grid.IsSheetLoaded(i));
            check(`${label}: the sheet the file was saved on is shown`, [grid.GetActiveSheetIndex(), grid.GetSheetInfos()[saved].isActive, loaded.filter(Boolean).length, loaded[saved]], [saved, true, 1, true]);
            const out = await grid.ExportXlsx();
            check(`${label}: unedited, it is exported as the bytes that were opened`, [out.length, out.every((b, i) => b === bytes[i])], [bytes.length, true]);
            // An edit on that sheet: the saved file still opens on it.
            grid.SetCell(40, 40, 'edited');
            const again = mountGrid();
            try {
                await again.grid.ImportXlsx(await grid.ExportXlsx());
                check(`${label}: saved after an edit, it opens on that sheet again`, again.grid.GetActiveSheetIndex(), saved);
            } finally { again.close(); }
        } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    return lines;
};

// GetPrintPages: the pages a sheet prints on, as the PDF export lays them out (the numbers
// below are what tests/print/helpers.fsx counts in the PDFs of the same workbooks).
checks['print-pages'] = async () => {
    const { lines, check } = makeLog();
    const expected = [
        ['tests/real-world/milk-street-office.xlsx', 'Annual', 6, 38, 'A1:L63'],
        ['tests/real-world/milk-street-office.xlsx', 'Monthly', 40, 33, 'A1:R58'],
        ['tests/real-world/lyric-multifamily.xlsx', 'Lyric-PF-DCF', 7, 32, 'A1:R57'],
        ['tests/real-world/lyric-multifamily.xlsx', 'Comp-Sales', 1, 40, 'A1:O24'],
        ['tests/real-world/lyric-multifamily.xlsx', 'Apt-Comps', 1, 41, 'A1:Q20'],
        ['tests/real-world/lyric-multifamily.xlsx', 'Replacement-Cost', 1, 54, 'A1:I39'],
        ['tests/real-world/dcf-stock-analysis.xlsx', 'Main', 2, 100, 'B2:J39'],
        ['tests/real-world/dcf-stock-analysis.xlsx', 'Financial Model', 10, 100, 'A1:I53'],
    ];
    for (const file of [...new Set(expected.map((e) => e[0]))]) {
        const { grid, close } = mountGrid();
        try {
            await grid.ImportXlsx(await bytesOf(file));
            const names = grid.GetSheetNames();
            for (const [, name, count, scale, first] of expected.filter((e) => e[0] === file)) {
                const index = names.indexOf(name), shown = grid.GetActiveSheetIndex();
                const print = grid.GetPrintPages(index);
                check(`${name}: pages, scale and the first page as in the PDF`, [print.pageCount, print.pages.length, print.scale, print.pages[0].range, print.sheetName, print.sheetIndex], [count, count, scale, first, name, index]);
                check(`${name}: asking does not switch sheets`, grid.GetActiveSheetIndex(), shown);
                check(`${name}: the setup is GetPageSetup's`, JSON.stringify(print.setup), JSON.stringify(grid.GetPageSetup(index)));
            }
        } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    // A sheet of the editor's own: breaks, fitting, titles, headers.
    const { grid, close } = mountGrid();
    try {
        for (let row = 0; row < 120; row++) for (let col = 0; col < 4; col++) grid.SetCell(row, col, String(row * 4 + col));
        const plain = grid.GetPrintPages();
        check('a new sheet: its used range in pages down the sheet', [plain.printAreas, plain.scale, plain.pageCount > 1, plain.pages[0].firstRow, plain.pages[plain.pageCount - 1].lastRow, plain.pages.every((p) => p.firstCol === 0 && p.lastCol === 3 && p.area === 0), plain.pages[0].header, plain.pages[0].number],
            [['A1:D120'], 100, true, 0, 119, true, null, 1]);
        check('the pages follow one another', plain.pages.every((p, i) => i === 0 || p.firstRow === plain.pages[i - 1].lastRow + 1), true);
        grid.AddPageBreak(10);
        const broken = grid.GetPrintPages();
        check('a manual break ends a page', [broken.pageCount, broken.pages[0].lastRow + 1 === broken.pages[1].firstRow, broken.breaks.rows, [9, 10].includes(broken.pages[0].lastRow)], [plain.pageCount + 1, true, grid.GetPageBreaks().rows, true]);
        grid.SetPageSetup({ header: '&LPage &P of &N&C&A&R&F', footer: 'Total &N', firstPageNumber: 5, printTitles: { rows: '1:2' } });
        const titled = grid.GetPrintPages(null, 'book.xlsx');
        check('print titles ride along: the body starts below them', [titled.printTitles, titled.pages[0].firstRow], [{ rows: { first: 0, last: 1 }, columns: null }, 2]);
        check('headers and footers as printed', [titled.pages[1].header, titled.pages[1].footer, titled.pages[1].number, titled.pages[1].index],
            [{ left: `Page 6 of ${titled.pageCount}`, center: 'Sheet1', right: 'book.xlsx' }, { left: '', center: `Total ${titled.pageCount}`, right: '' }, 6, 1]);
        grid.SetPageSetup({ fitToWidth: 1, fitToHeight: 1 });
        const fitted = grid.GetPrintPages();
        check('fitted to one page: one page per manual break, at a smaller scale', [fitted.pageCount <= 2, fitted.scale < 100, fitted.setup.fitToWidth], [true, true, 1]);
        grid.ClearPageBreaks();
        check('... and one page without the break', grid.GetPrintPages().pageCount, 1);
        grid.SetPageSetup({ scale: 100, orientation: 'landscape' });
        check('landscape: fewer rows to a page', grid.GetPrintPages().pageCount > plain.pageCount - 1 && grid.GetPrintPages().pages[0].lastRow < plain.pages[0].lastRow, true);
        let thrown = null; try { grid.GetPrintPages(7); } catch (e) { thrown = String(e.message || e); }
        check('a sheet that is not there', thrown !== null && /GetPrintPages/.test(thrown), true);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// A chart sheet has no cells: what was typed into one was shown and then lost at the save.
// It takes no entry now (the grid is a viewer while it is on display), and says what it is.
checks['chart-sheet-read-only'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const original = await bytesOf('tests/test-documents-excel/pandas-chartsheet.xlsx');
        await grid.ImportXlsx(original.slice());
        const names = grid.GetSheetNames(), kinds = names.map((_, i) => grid.GetSheetKind(i));
        const chart = kinds.indexOf('chartsheet'), sheet = kinds.indexOf('worksheet');
        check('the kinds of the sheets', [chart >= 0, sheet >= 0, grid.GetSheetInfos().map((s) => s.kind), grid.GetSheetKind()], [true, true, kinds, kinds[grid.GetActiveSheetIndex()]]);
        grid.SwitchSheet(chart);
        const values = () => grid.GetRange('A1:C3').flat().map((cell) => cell.value ?? null).filter((v) => v !== null && v !== '');
        check('on display: a chart sheet, not the host\'s viewer mode', [grid.GetSheetKind(), grid.IsReadOnly()], ['chartsheet', false]);
        check('EnterText is refused, SetCell changes nothing', [grid.EnterText(0, 0, 'typed'), (grid.SetCell(1, 1, 'set'), values()), grid.IsModified()], [false, [], false]);
        // The user's keys: no entry starts.
        grid.SelectCell(2, 2);
        host.querySelector('.spreadsheet-content').focus();
        if (window.realType) { await window.realType('abc'); await window.realKey('Enter'); }
        check('typing starts no entry', [values(), !!grid.GetModel().EditingCell], [[], false]);
        // Saved: the chart sheet is the part it was read from.
        const saved = await grid.ExportXlsx();
        const zip0 = await window.JSZip.loadAsync(original), zip1 = await window.JSZip.loadAsync(saved);
        const part = Object.keys(zip0.files).find((path) => /^xl\/chartsheets\/[^/]+\.xml$/.test(path));
        // (But for tabSelected: it is the sheet on display now.)
        const text = async (zip) => (await zip.files[part].async('string')).replace(/\s+tabSelected="[^"]*"/, '');
        check('saved: the chart sheet\'s part is the file\'s', [!!part, part && zip1.files[part] ? await text(zip1) === await text(zip0) : null], [true, true]);
        // A worksheet takes entries as before, also right after the chart sheet was shown.
        grid.SwitchSheet(sheet);
        check('a worksheet is edited as before', [grid.GetSheetKind(), grid.EnterText(30, 5, 'typed'), grid.GetRange('F31')[0][0].value], ['worksheet', true, 'typed']);
        grid.SetReadOnly(true); grid.SwitchSheet(chart); grid.SwitchSheet(sheet);
        check('the host\'s viewer mode stays through a chart sheet', grid.IsReadOnly(), true);
        grid.SetReadOnly(false);
        check('... and ends when the host ends it', [grid.IsReadOnly(), grid.EnterText(31, 5, 'again')], [false, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
