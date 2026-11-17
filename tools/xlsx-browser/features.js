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
        check('Medium2: header in the accent with white bold text, banded rows', [look(0, 0), look(1, 0)[0], look(2, 0)[0]], [['rgb(68, 114, 196)', white, true], 'rgb(217, 225, 242)', white]);
        grid.SetTableStyle('Sales', 'TableStyleLight1'); await frames();
        // (A cell draws the line above it - the one under the header is the first row's -, a thin one as its grid line.)
        check('Light1: a bold header without a fill, grey bands, lines above and below the header', [look(0, 0), look(1, 0)[0], [node(0, 0), node(1, 0)].every((cell) => cell.querySelector('.spreadsheet-edge') !== null || getComputedStyle(cell).boxShadow.includes('rgb(0, 0, 0)'))], [[white, black, true], 'rgb(217, 217, 217)', true]);
        grid.SetTableStyle('Sales', 'TableStyleDark2'); await frames();
        check('Dark2: dark fills, white text', [look(0, 0), look(1, 0).slice(0, 2), look(2, 0).slice(0, 2)], [[black, white, true], ['rgb(48, 84, 150)', white], ['rgb(68, 114, 196)', white]]);
        grid.SetTableStyle('Sales', { name: 'TableStyleMedium9', showFirstColumn: true }); grid.SetTableTotalRow('Sales', true); await frames();
        check('Medium9: bands in two tints, the first column and the total row in the accent', [look(1, 1)[0], look(2, 1)[0], look(1, 0), look(4, 1)], ['rgb(180, 198, 231)', 'rgb(217, 225, 242)', ['rgb(68, 114, 196)', white, true], ['rgb(68, 114, 196)', white, true]]);
        // (Its thin white lines inside the table are the cells' grid lines, which cover the corners where they
        // cross: a node at every corner made scrolling a long table a node a cell slower - grid-scroll-frames.)
        check('Medium9: the white lines inside drawn as grid lines, no node at the corners', [2, 3].flatMap((row) => [1, 2].map((col) => [node(row, col).querySelectorAll('.spreadsheet-edge').length, getComputedStyle(node(row, col)).boxShadow.includes(white)])), [[0, true], [0, true], [0, true], [0, true]]);
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

// The web grid draws every cell with the look the painters (canvas, PDF, native hosts) draw it
// with (`Paint.cellLook`), mapped to CSS: where the two disagreed, Excel's way. A workbook as
// Excel writes it: a number format with a [Red] section, text centred across the selection, a
// wrapped cell with a line break of its own, a tall row (bottom, Excel's default), General
// numbers too long for their column, a conditional format's underline over a struck font, an
// indent; then a plain data bar, an icon set and a table without a style through the host API.
const STYLES_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const LOOK_STYLES = DECL + `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="0;[Red]-0"/></numFmts>`
    + '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><strike/><sz val="11"/><name val="Calibri"/></font></fonts>'
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="centerContinuous"/></xf>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment indent="2"/></xf>'
    + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><u/></font></dxf></dxfs></styleSheet>';
const ICON_RULE = { type: 'iconSet', iconSet: '3TrafficLights1' };
const inlineCell = (ref, text, style) => `<c r="${ref}"${style ? ` s="${style}"` : ''} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`;
const LOOK_SHEET = '<sheetFormatPr defaultRowHeight="15"/><sheetData>'
    + '<row r="1"><c r="A1" s="1"><v>-5</v></c><c r="B1" s="1"><v>5</v></c></row>'
    + `<row r="2">${inlineCell('A2', 'Title', 2)}<c r="B2" s="2"/><c r="C2" s="2"/></row>`
    + `<row r="3" ht="45" customHeight="1">${inlineCell('A3', 'one&#10;two', 3)}${inlineCell('B3', 'low')}</row>`
    + '<row r="4"><c r="A4"><v>0.123456789123</v></c><c r="B4"><v>123456789012345678</v></c></row>'
    + `<row r="5">${inlineCell('A5', 'x', 4)}${inlineCell('B5', 'in', 5)}</row>`
    + '</sheetData><conditionalFormatting sqref="A5"><cfRule type="expression" dxfId="0" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>';
// A text in runs draws each run in its look (Excel's "OPEN vs. CLOSED" titles), and an accounting
// format's fill puts the currency sign at the cell's left edge and the number at its right.
const RUNS_STYLES = DECL + `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* \\(#,##0.00\\);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)"/></numFmts>`
    + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
const RUNS_SHEET = '<cols><col min="1" max="1" width="40" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/></cols><sheetData><row r="1">'
    + '<c r="A1" t="inlineStr"><is><r><rPr><b/><sz val="11"/><color rgb="FF00B0F0"/><rFont val="Calibri"/></rPr><t>OPEN</t></r><r><rPr><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve"> vs. </t></r>'
    + '<r><rPr><i/><sz val="14"/><color rgb="FFFFC000"/><rFont val="Calibri"/></rPr><t>CLOSED</t></r></is></c>'
    + '<c r="B1" s="1"><v>1234.5</v></c></row></sheetData>';
checks['cell-runs-and-accounting'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx(await excelPackage(RUNS_SHEET, { 'xl/styles.xml': [STYLES_TYPE, STYLES_REL, RUNS_STYLES] }));
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const spans = [...node(0, 0).querySelectorAll('span span')].map((span) => { const css = getComputedStyle(span); return [span.textContent, css.color, Number(css.fontWeight) >= 600, css.fontStyle, Math.round(parseFloat(css.fontSize) * 72 / 96)]; });
        check('each run of a text in its own look', spans,
            [['OPEN', 'rgb(0, 176, 240)', true, 'normal', 11], [' vs. ', 'rgb(0, 0, 0)', false, 'normal', 11], ['CLOSED', 'rgb(255, 192, 0)', false, 'italic', 14]]);
        const cell = node(0, 1).getBoundingClientRect(), [sign, number] = [...node(0, 1).querySelectorAll('span')].map((span) => span.getBoundingClientRect());
        // SetFontLoader: a family the system lacks (Calibri) is asked of the host when a cell in it is
        // drawn, and drawn in the face it hands.
        const response = await fetch('/font?name=calibri.ttf');
        if (!response.ok) note('SetFontLoader: no calibri.ttf (RB_WORD_FONT_DIR, artifacts/fonts/windows): not checked');
        else if (document.fonts.check('11px Calibri') && [...document.fonts].length === 0) note('SetFontLoader: the system has Calibri: not checked');
        else {
            const data = new Uint8Array(await response.arrayBuffer()), asked = [];
            grid.SetFontLoader(async (family, bold, italic) => { asked.push([family, bold, italic]); return family === 'Calibri' && !bold && !italic ? data : null; });
            for (let i = 0; i < 100 && ![...document.fonts].some((face) => face.family.replace(/"/g, '') === 'Calibri' && face.status === 'loaded'); i++) await new Promise((resolve) => setTimeout(resolve, 20));
            await new Promise((resolve) => setTimeout(resolve, 80));
            await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            const context = document.createElement('canvas').getContext('2d'); context.font = `${11 * 96 / 72}px Calibri`;
            const range = document.createRange(); range.selectNodeContents(node(0, 0).querySelectorAll('span span')[1]);
            check('SetFontLoader: the faces the cells are drawn in are asked for once each', asked.map((face) => face.join(' ')).sort(), ['Calibri false false', 'Calibri false true', 'Calibri true false']);
            check('... and Calibri text is drawn in the Calibri handed', Math.abs(range.getBoundingClientRect().width - context.measureText(' vs. ').width) < 1.5, true);
        }
        check('accounting: the $ at the cell\'s left, the number at its right', [node(0, 1).textContent.replace(/\s+/g, ' '), sign.left - cell.left <= 8, cell.right - number.right <= 8, number.left - sign.right > 20], [' $1,234.50 ', true, true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
checks['cell-look-shared'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx(await excelPackage(LOOK_SHEET, { 'xl/styles.xml': [STYLES_TYPE, STYLES_REL, LOOK_STYLES] }));
        const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        await frame();
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const css = (row, col) => getComputedStyle(node(row, col));
        const ink = (row, col) => { const range = document.createRange(); range.selectNodeContents(node(row, col)); return range.getBoundingClientRect(); };
        check('a negative number of a [Red] section is red, the positive one is not', [css(0, 0).color, css(0, 1).color], ['rgb(255, 0, 0)', 'rgb(0, 0, 0)']);
        const [a2, c2] = [node(1, 0).getBoundingClientRect(), node(1, 2).getBoundingClientRect()], title = ink(1, 0);
        check('centred across the selection: over A2:C2, not in A2', Math.abs((title.left + title.right) / 2 - (a2.left + c2.right) / 2) <= 3, true);
        const wrapped = ink(2, 0), lineHeight = parseFloat(css(2, 0).lineHeight);
        check('a wrapped cell keeps its own line break: two lines', wrapped.height > 1.5 * lineHeight, true);
        const low = ink(2, 1), b3 = node(2, 1).getBoundingClientRect();
        check('in a tall row the text sits at the bottom (Excel\'s default)', b3.bottom - low.bottom <= 6, true);
        check('General numbers to the digits their column has room for (Excel: 0.123457, 1.23E+17)', [node(3, 0).textContent, node(3, 1).textContent], ['0.123457', '1.23E+17']);
        check('a conditional format\'s underline over a struck font: both', ['underline', 'line-through'].map((line) => css(4, 0).textDecorationLine.includes(line)), [true, true]);
        check('an indent on General text pads its left', parseFloat(css(4, 1).paddingLeft) > 15, true);
        // Through the host API.
        grid.SetCell(0, 3, '5'); grid.SetCell(1, 3, '10'); grid.SetCell(2, 3, '1');
        const bar = grid.AddConditionalFormat('D1:D2', { type: 'dataBar', min: { type: 'number', value: 0 }, max: { type: 'number', value: 10 }, color: '#FF0000' });
        await frame();
        const plainBar = node(0, 3).querySelector('.spreadsheet-cf-bar');
        check('a data bar without the x14 part (Excel 2007\'s) fades to white, two pixels in from the cell\'s top and bottom', [!!plainBar && /linear-gradient\(to right, .*rgb\(255, 255, 255\)\)/.test(getComputedStyle(plainBar).backgroundImage), !!plainBar && getComputedStyle(plainBar).top, !!plainBar && getComputedStyle(plainBar).bottom], [true, '2px', '2px']);
        grid.RemoveConditionalFormat(bar);
        grid.AddConditionalFormat('D1:D3', ICON_RULE);
        await frame();
        const icon = node(0, 3).querySelector('[data-cf-icon]');
        check('an icon at the left of its cell, its number at the right', [!!icon, !!icon && icon.getBoundingClientRect().left - node(0, 3).getBoundingClientRect().left <= 6, css(0, 3).textAlign], [true, true, 'right']);
        grid.SetCell(0, 5, 'Name'); grid.SetCell(0, 6, 'Value'); grid.SetCell(1, 5, 'a'); grid.SetCell(1, 6, '1'); grid.SetCell(2, 5, 'b'); grid.SetCell(2, 6, '2');
        grid.DefineTable('F1:G3', 'Plain'); grid.SetTableStyle('Plain', 'none');
        await frame();
        check('a table without a style draws as plain cells (no header fill, no bands)', [css(0, 5).backgroundColor, css(0, 5).fontWeight, css(0, 5).textAlign, css(2, 5).backgroundColor], ['rgb(255, 255, 255)', '400', 'start', 'rgb(255, 255, 255)']);
        // Every cell: the grid's text, fill, colour and alignment are the painters' look.
        const { cellLook } = await import('/spreadsheet/Paint.js');
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        const model = grid.GetModel();
        const rgbOf = (hex) => hex ? `rgb(${parseInt(hex.slice(1, 3), 16)}, ${parseInt(hex.slice(3, 5), 16)}, ${parseInt(hex.slice(5, 7), 16)})` : null;
        const differ = [];
        for (let row = 0; row < 5; row++)
            for (let col = 0; col < 7; col++) {
                const cell = node(row, col);
                if (!cell) continue;
                const look = cellLook(model, row, col, getCell(new CellAddress(row, col), model), undefined, cell.getBoundingClientRect().width);
                const style = getComputedStyle(cell);
                const shown = [...cell.childNodes].filter((child) => child.nodeType === 3 || child.matches('.spreadsheet-rotated, .spreadsheet-across')).map((child) => child.textContent).join('');
                const align = ['left', 'center', 'right'][look.Align.tag];
                const got = [shown, style.backgroundColor, style.color, look.Horizontal === 'centercontinuous' ? 'center' : (style.textAlign === 'start' ? 'left' : style.textAlign)];
                const want = [look.Text, rgbOf(look.Fill) ?? 'rgb(255, 255, 255)', rgbOf(look.TextColor) ?? 'rgb(0, 0, 0)', align];
                if (JSON.stringify(got) !== JSON.stringify(want)) differ.push(`${row},${col}: ${JSON.stringify(got)} vs ${JSON.stringify(want)}`);
            }
        check('every cell: text, fill, colour and alignment as the painters\' look', differ, []);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// What a PDF holds, read back: its pages and the strings its pages show (each content stream's text
// operators, the streams inflated: the glyph codes of the embedded subsets, which are the same for two
// PDFs set in the same font files).
const latin1 = (bytes) => { let text = ''; for (let i = 0; i < bytes.length; i += 0x8000) text += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return text; };
const inflate = async (bytes) => new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate'))).arrayBuffer());
export const pdfSummary = async (bytes) => {
    const text = latin1(bytes);
    const pages = (text.match(/\/Type\s*\/Page(?![a-zA-Z])/g) ?? []).length;
    const shown = [];
    let drawing = 0;
    const streamPattern = /<<((?:(?!>>\s*stream)[\s\S])*)>>\s*stream\r?\n/g;
    let match;
    while ((match = streamPattern.exec(text))) {
        const start = match.index + match[0].length, end = text.indexOf('endstream', start);
        if (end < 0) break;
        let body = bytes.subarray(start, end);
        if (/\/FlateDecode/.test(match[1])) { try { body = await inflate(body); } catch { continue; } }
        const ops = latin1(body);
        if (!/\bBT\b|\bre\b|\bc\b/.test(ops)) continue;
        for (const run of ops.matchAll(/(<[0-9A-Fa-f\s]*>|\((?:\\.|[^\\)])*\))\s*Tj|\[((?:[^\]])*)\]\s*TJ/g)) shown.push(run[0].replace(/\s*T[jJ]$/, ''));
        drawing += (ops.match(/(^|\s)(re|c|l)(\s|$)/g) ?? []).length;
    }
    return { pages, shown, drawing };
};

// The PDF of a sheet in the browser (src/Spreadsheet.Pdf in dist/spreadsheet-pdf, loaded when first
// asked for): the dashboard of dashboard-sales-analysis.xlsx - charts and conditional formatting - by its
// page setup, with the fonts of RB_WORD_FONT_DIR; its pages and text are those of the .NET PDF of the same
// sheet (tools/spreadsheet-pdf-runner, SHEET_PDF_REFERENCE) when that is given.
const PDF_FONTS = [['Calibri', 'calibri'], ['Arial', 'arial']].flatMap(([family, file]) =>
    [[false, false, ''], [true, false, file === 'calibri' ? 'b' : 'bd'], [false, true, 'i'], [true, true, file === 'calibri' ? 'z' : 'bi']]
        .map(([bold, italic, suffix]) => ({ family, bold, italic, file: file + suffix + '.ttf' })));
checks['sheet-pdf'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const pdfLoaded = () => performance.getEntriesByType('resource').some((entry) => /\/spreadsheet-pdf\//.test(entry.name));
        await grid.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
        // (The pivot sheets the dashboard's charts read are converted, as the .NET runner reads them: a chart
        // of a sheet not converted yet draws the cache of its part, which here is not what those cells hold.)
        for (let index = 1; index <= 6; index++) grid.SwitchSheet(index);
        grid.SwitchSheet(0);
        check('the PDF writer is not loaded with the grid', pdfLoaded(), false);
        const fonts = [];
        for (const face of PDF_FONTS) {
            const response = await fetch('/font?name=' + face.file);
            if (response.ok) fonts.push({ family: face.family, bold: face.bold, italic: face.italic, data: new Uint8Array(await response.arrayBuffer()) });
        }
        if (fonts.length < PDF_FONTS.length) note(`fonts: ${fonts.length} of ${PDF_FONTS.length} faces found (RB_WORD_FONT_DIR): text is not compared`);
        grid.SetPdfFonts(fonts);
        const bytes = await grid.ExportPdf({ fileName: 'dashboard-sales-analysis' });
        check('then loaded, once', [pdfLoaded(), performance.getEntriesByType('resource').filter((entry) => /\/spreadsheet-pdf\/Pdf\.js$/.test(entry.name)).length], [true, 1]);
        if (window.keepPdf) { let binary = ''; for (const b of bytes) binary += String.fromCharCode(b); await window.keepPdf('browser-dashboard.pdf', btoa(binary)); }
        const summary = await pdfSummary(bytes);
        const printed = grid.GetPrintPages(0);
        check('a PDF of the pages GetPrintPages lays out', [latin1(bytes.subarray(0, 5)), summary.pages, summary.pages > 0], ['%PDF-', printed.pageCount, true]);
        check('with its charts drawn (paths) and, with fonts, its text', [summary.drawing > 200, fonts.length === 0 || summary.shown.length > 50], [true, true]);
        const reference = await fetch('/sheet-pdf-reference');
        if (reference.ok && fonts.length === PDF_FONTS.length) {
            const dotnet = await pdfSummary(new Uint8Array(await reference.arrayBuffer()));
            check('the .NET PDF of the sheet: the same pages', summary.pages, dotnet.pages);
            const sorted = (list) => [...list].sort();
            const a = sorted(summary.shown), b = sorted(dotnet.shown);
            const missing = b.filter((run, i) => run !== a[i]).length;
            check('... and the same text', [a.length, missing], [b.length, 0]);
            note(`browser ${summary.pages} pages, ${summary.shown.length} text runs, ${summary.drawing} path segments; .NET ${dotnet.pages} pages, ${dotnet.shown.length} text runs, ${dotnet.drawing} path segments`);
        } else note('no .NET PDF to compare with (SHEET_PDF_REFERENCE)');
        // The selection, the whole workbook.
        grid.SelectRange(0, 0, 4, 2);
        const selection = await pdfSummary(await grid.ExportPdf({ what: 'selection' }));
        check('the selection: one page', selection.pages, 1);
        const visible = grid.SheetVisibility();
        const expected = grid.GetSheetNames().reduce((n, _, i) => n + (visible[i] && grid.GetSheetKind(i) === 'worksheet' ? Math.max(1, grid.GetPrintPages(i).pageCount) : 0), 0);
        const whole = await pdfSummary(await grid.ExportPdf({ what: 'workbook' }));
        check('the workbook: every visible worksheet\'s pages', whole.pages, expected);
        const thrown = async (call) => { try { await call(); return null; } catch (e) { return String(e.message); } };
        check('ExportPdf refuses what it is not', [/^ExportPdf: /.test(await thrown(() => grid.ExportPdf({ what: 'page' })) ?? ''), /^ExportPdf: /.test(await thrown(() => grid.ExportPdf({ size: 1 })) ?? ''), /^SetPdfFonts: /.test(await thrown(() => grid.SetPdfFonts([{ family: 'X' }])) ?? '')], [true, true, true]);
        // Printing: the PDF in a hidden frame, the browser's print dialog.
        const printing = grid.PrintPdf();
        await Promise.race([printing, new Promise((resolve) => setTimeout(resolve, 5000))]);
        const frame = document.querySelector('iframe[data-spreadsheet-print]');
        check('PrintPdf: the PDF in a hidden frame of the page', [!!frame, frame && frame.style.visibility, frame && /^blob:/.test(frame.src)], [true, 'hidden', true]);
        if (frame) frame.remove();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// VBA macros in the browser (MacroHost through the grid's API): listed, off until enabled, Workbook_Open
// on enabling, one undo step per run, MsgBox / InputBox / Debug.Print through the host's hooks, a macro
// that stops on an error keeps nothing, the selection it leaves, and the project saved as it was.
checks['macros'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const original = await bytesOf('tests/test-documents-excel/synthetic-macros.xlsm');
        await grid.ImportXlsx(original);
        const engineFetched = () => performance.getEntriesByType('resource').some((entry) => /MacroProject\.js|Rbt\.Vb/.test(entry.name));
        check('it has macros, known without loading the VBA engine', [grid.HasMacros, grid.GetMacros().length, engineFetched()], [true, 0, false]);
        await grid.LoadMacros();
        check('loaded: listed as Excel lists them', [Array.from(grid.GetMacros()), engineFetched()], [['FillTotals', 'Greet', 'Broken', 'Sheet1.Hello'], true]);
        check('its modules, their kinds, the source', grid.GetMacroModules().map((m) => [m.name, m.kind, m.source.length > 0]), [['Module1', 'standard', true], ['ThisWorkbook', 'document', true], ['Sheet1', 'document', true]]);
        check('the project compiles', grid.GetMacroProblems().length, 0);
        let refused = null;
        try { await grid.RunMacro('FillTotals'); } catch (e) { refused = String(e.message); }
        check('macros do not run until enabled', [grid.MacrosEnabled, /not enabled/.test(refused ?? '')], [false, true]);
        const opened = await grid.EnableMacros();
        check('enabling runs Workbook_Open once', [opened.ok, grid.GetRange('D1:D1')[0][0].value], [true, 'opened']);
        await grid.EnableMacros();
        const out = [];
        grid.SetMacroHooks({ output: (text) => out.push(text), messageBox: (prompt, buttons, title) => { out.push(`box:${prompt}|${buttons}|${title}`); return 6; }, inputBox: () => 'Ada' });
        const before = [grid.GetRange('B1:B1')[0][0].value, grid.GetRange('C1:C1')[0][0].value];
        const filled = await grid.RunMacro('FillTotals');
        check('a macro runs on the grid', [filled.ok, grid.GetRange('B1:B5').map((row) => row[0].value), grid.GetRange('C1:C2').map((row) => row[0].value)], [true, [20, 40, 60, 80, 100], [300, 301]]);
        check('the format it set and the selection it leaves', [grid.GetFormatOf('C1').bold, grid.GetSelectionRef()], [true, 'B1']);
        check('Debug.Print to the host', out.some((text) => /^filled\s*5\s*$/.test(text)), true);
        grid.Undo();
        check('one undo step', [grid.GetRange('B1:B1')[0][0].value, grid.GetRange('C1:C1')[0][0].value, grid.GetRange('D1:D1')[0][0].value], [...before, 'opened']);
        grid.Redo();
        check('and redone', grid.GetRange('C1:C1')[0][0].value, 300);
        const greeted = await grid.RunMacro('Greet');
        check('MsgBox and InputBox through the hooks', [greeted.ok ? true : greeted.description, out.includes('box:Write the answer?|4|Greeting'), grid.GetRange('E1:E2').map((row) => row[0].value)], [true, true, [6, 'Ada']]);
        const hello = await grid.RunMacro('Sheet1.Hello');
        check('a macro of a sheet\'s module', [hello.ok, grid.GetRange('D2:D2')[0][0].value], [true, 'hello from ' + grid.GetSheetNames()[0]]);
        grid.EnterText(5, 0, '7');
        check('Worksheet_Change of the sheet\'s module runs on an edit', grid.GetRange('G6:G6')[0][0].value, 700);
        grid.Undo();
        check('and undoes with it', [grid.GetRange('A6:A6')[0][0].value, grid.GetRange('G6:G6')[0][0].value], [null, null]);
        const broken = await grid.RunMacro('Broken');
        check('a macro that stops: its error and where, nothing kept', [broken.ok, broken.number, /Subscript/.test(broken.description ?? ''), broken.module, broken.procedure, grid.GetRange('A1:A1')[0][0].value], [false, 9, true, 'Module1', 'Broken', 10]);
        const zip0 = await window.JSZip.loadAsync(original);
        const zip1 = await window.JSZip.loadAsync(await grid.ExportXlsx());
        const binary = async (zip) => Array.from(await zip.files['xl/vbaProject.bin'].async('uint8array')).join();
        check('saved: the project as it was, the values the macros wrote', [await binary(zip1) === await binary(zip0), /<c r="C1"[^>]*><f>SUM\(B1:B5\)<\/f><v>300<\/v>/.test(await zip1.files['xl/worksheets/sheet1.xml'].async('string'))], [true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// A workbook that runs no macro never fetches the VBA engine: not with the grid, not when a workbook
// with macros is opened, not when cells are edited; it is fetched when macros are loaded or enabled.
checks['vba-loads-on-use'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const fetched = () => performance.getEntriesByType('resource').filter((entry) => /MacroProject\.js|MacroHost\.js|Rbt\.Vb|rbt-vb/.test(entry.name)).map((entry) => entry.name.split('/').slice(-2).join('/'));
        const before = fetched().length;
        // (A module is fetched once per page: after an earlier check loaded the engine there is nothing to see.
        // The build's import-graph check, check-lazy-modules, holds either way.)
        if (before > 0) { note('the VBA engine was loaded by an earlier check of this page'); close(); return lines; }
        grid.SetCell(0, 0, '1');
        await grid.ImportXlsx(await bytesOf('tests/test-documents-excel/synthetic-namedrange.xlsx'));
        grid.EnterText(0, 1, '2');
        await grid.ImportXlsx(await bytesOf('tests/test-documents-excel/synthetic-macros.xlsm'));
        grid.EnterText(5, 0, '3');
        check('no VBA module fetched for a grid, an .xlsx, an .xlsm opened and edited', fetched().length - before, 0);
        await grid.LoadMacros();
        check('fetched when the macros are loaded', fetched().length - before > 0, true);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// The VBA editor from the shell's Data menu: in a workbook without macros a blank Module1 to write in,
// Apply compiles it, its macro runs on the grid (one undo step), the Immediate line evaluates, the output
// shows each, and a save writes the project from the module.
checks['shell-vba-editor'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const shellHost = document.createElement('div');
    document.body.appendChild(shellHost);
    const { createSpreadsheetShell } = await import('/spreadsheet/Toolbar.js');
    const shell = createSpreadsheetShell(grid, () => {});
    const settle = () => new Promise((resolve) => setTimeout(resolve, 30));
    const type = async (testId, value, event = 'input') => {
        const field = shellHost.querySelector(`[data-testid="${testId}"]`);
        field.value = value; field.dispatchEvent(new Event(event, { bubbles: true })); await settle();
    };
    const press = async (testId) => { shellHost.querySelector(`[data-testid="${testId}"]`).click(); await settle(); };
    const output = () => shellHost.querySelector('[data-testid="dlg-field-output"]')?.textContent ?? '';
    try {
        grid.SetCell(0, 0, '5');
        shell.Mount(shellHost);
        shellHost.querySelector('[data-sheet-menu="data"]').click();
        shellHost.querySelector('[data-sheet-menu-item="vba-editor"]').click();
        await settle(); await settle();
        check('the editor opens with a blank Module1', [!!shellHost.querySelector('[data-testid="xl-dlg-vba"]'), shellHost.querySelector('[data-testid="dlg-field-module"]')?.value, shellHost.querySelector('[data-testid="dlg-field-source"]')?.value], [true, 'Module1', '']);
        await type('dlg-field-source', 'Sub Double()\n    Range("B1") = Range("A1") * 2\n    Debug.Print "doubled"; Range("B1")\nEnd Sub');
        await press('dlg-vba-apply');
        check('Apply compiles it', [/Module1: compiled/.test(output()), Array.from(grid.GetMacros())], [true, ['Double']]);
        grid.SetMacroHooks({ output: () => {} });
        check('shown again, its macro listed', shellHost.querySelector('[data-testid="dlg-field-macro"]')?.value, 'Double');
        await press('dlg-vba-run');
        check('Run runs it on the grid, enabling macros', [grid.MacrosEnabled, grid.GetRange('B1:B1')[0][0].value, /doubled\s+10/.test(output())], [true, 10, true]);
        grid.Undo();
        check('one undo step', grid.GetRange('B1:B1')[0][0].value, null);
        grid.Redo();
        await type('dlg-field-immediate', '?Range("B1").Value + 1');
        await press('dlg-vba-immediate');
        check('the Immediate line evaluates, the output shows it', [/> \?Range\("B1"\)\.Value \+ 1/.test(output()), /\b11\b/.test(output())], [true, true]);
        await type('dlg-field-immediate', 'Range("C1") = "set from the Immediate window"');
        await press('dlg-vba-immediate');
        check('an Immediate statement changes the grid', grid.GetRange('C1:C1')[0][0].value, 'set from the Immediate window');
        await type('dlg-field-source', 'Sub Double(\nEnd Sub');
        await press('dlg-vba-apply');
        check('a syntax error is shown, the dialog stays', [/problem/.test(output()), !!shellHost.querySelector('[data-testid="xl-dlg-vba"]'), grid.GetMacroProblems().length > 0], [true, true, true]);
        await type('dlg-field-source', 'Sub Double()\n    Range("B1") = Range("A1") * 3\nEnd Sub');
        await press('dlg-vba-apply');
        const zip = await window.JSZip.loadAsync(await grid.ExportXlsx());
        check('saved: a macro-enabled workbook with the project', [!!zip.files['xl/vbaProject.bin'], /macroEnabled/.test(await zip.files['[Content_Types].xml'].async('string'))], [true, true]);
        await press('dlg-vba-close');
        check('Close closes it', !!shellHost.querySelector('[data-testid="xl-dlg-vba"]'), false);
        const reopened = mountGrid();
        try {
            await reopened.grid.ImportXlsx(await zip.generateAsync({ type: 'uint8array' }));
            await reopened.grid.LoadMacros();
            check('reopened: the module and its code', reopened.grid.GetMacroModules().map((m) => [m.name, /Range\("A1"\) \* 3/.test(m.code)]), [['Module1', true]]);
        } finally { reopened.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    shell.Dispose(); shellHost.remove(); close();
    return lines;
};

// A macro-enabled workbook saved from the browser keeps its VBA project byte for byte, with its signature
// and the code names that tie the document modules to the workbook and its sheets - unedited and edited.
checks['xlsm-vba-project'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const original = await bytesOf('tests/test-documents-excel/synthetic-vba-project.xlsm');
        const zip0 = await window.JSZip.loadAsync(original);
        const binary = async (zip, part) => zip.files[part] ? Array.from(await zip.files[part].async('uint8array')).join() : null;
        await grid.ImportXlsx(original);
        for (const [label, edit] of [['unedited', () => {}], ['a cell edited', () => grid.EnterText(9, 0, 'edited')]]) {
            edit();
            const zip1 = await window.JSZip.loadAsync(await grid.ExportXlsx());
            const part = async (name) => zip1.files[name] ? zip1.files[name].async('string') : '';
            check(label + ': the project and its signature, byte for byte', [await binary(zip1, 'xl/vbaProject.bin') === await binary(zip0, 'xl/vbaProject.bin'), await binary(zip1, 'xl/vbaProjectSignature.bin') === await binary(zip0, 'xl/vbaProjectSignature.bin'), !!zip1.files['xl/_rels/vbaProject.bin.rels']], [true, true, true]);
            const types = await part('[Content_Types].xml');
            check(label + ': a macro-enabled workbook, the project typed', [/macroEnabled\.main\+xml/.test(types), /PartName="\/xl\/vbaProject\.bin" ContentType="application\/vnd\.ms-office\.vbaProject"/.test(types), ((await part('xl/_rels/workbook.xml.rels')).match(/\/vbaProject"/g) ?? []).length], [true, true, 1]);
            check(label + ': the code names', [/codeName="ThisWorkbook"/.test(await part('xl/workbook.xml')), /codeName="Sheet1"/.test(await part('xl/worksheets/sheet1.xml'))], [true, true]);
        }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// No host fonts: ExportPdf sets the text in the package's own faces (dist/fonts, DefaultFonts), fetching
// only the faces the text needs: the dashboard's Calibri as Carlito, its Garamond as Liberation Serif, its
// Segoe UI as Liberation Sans.
checks['sheet-pdf-default-fonts'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    try {
        await grid.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
        for (let index = 1; index <= 6; index++) grid.SwitchSheet(index);
        grid.SwitchSheet(0);
        grid.SetPdfFonts([]);
        const bytes = await grid.ExportPdf({ fileName: 'dashboard-sales-analysis' });
        if (window.keepPdf) { let binary = ''; for (const b of bytes) binary += String.fromCharCode(b); await window.keepPdf('browser-dashboard-default-fonts.pdf', btoa(binary)); }
        const summary = await pdfSummary(bytes);
        const text = latin1(bytes);
        const fetched = performance.getEntriesByType('resource').filter((entry) => /\/fonts\/[^/]+\.ttf$/.test(entry.name)).map((entry) => entry.name.split('/').pop());
        check('the PDF has its text, in the default faces', [summary.pages > 0, summary.shown.length > 50, /\/BaseFont\s*\/[A-Z]{6}\+Carlito/.test(text)], [true, true, true]);
        const embedded = [...text.matchAll(/\/BaseFont\s*\/[A-Z]{6}\+([A-Za-z-]+)/g)].map((m) => m[1] + '.ttf');
        check('the faces of its families, and only those its text needs', [fetched.includes('Carlito-Regular.ttf'), fetched.includes('LiberationSerif-Regular.ttf'), [...new Set(fetched)].sort().join()], [true, true, [...new Set(embedded)].sort().join()]);
        note(`fetched ${[...new Set(fetched)].join(', ')}; ${summary.pages} pages, ${summary.shown.length} text runs`);
        const again = await grid.ExportPdf({ fileName: 'dashboard-sales-analysis' });
        const refetched = performance.getEntriesByType('resource').filter((entry) => /\/fonts\/[^/]+\.ttf$/.test(entry.name)).length;
        check('a second export fetches nothing again', [again.length > 0, refetched], [true, fetched.length]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// File > Print PDF must use the mounted grid's real lazy-loaded PDF writer. Chromium
// headless has no print dialog: intercept only the iframe's platform print() call,
// leaving export and the PDF blob/iframe handoff untouched.
const waitForPrintStatus = async (status, pending) => {
    for (let i = 0; i < 200; i++) {
        if (status.textContent !== pending) return status.textContent;
        await new Promise((resolve) => setTimeout(resolve, 50));
    }
    return status.textContent;
};
checks['shell-print-pdf'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const shellHost = document.createElement('div');
    document.body.appendChild(shellHost);
    const { createSpreadsheetShell } = await import('/spreadsheet/Toolbar.js');
    const contentWindow = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'contentWindow');
    let printCalls = 0, engineCalls = 0;
    const originalPrint = grid.PrintPdf.bind(grid);
    grid.PrintPdf = (...args) => { engineCalls++; return originalPrint(...args); };
    const actions = [];
    const shell = createSpreadsheetShell(grid, (action) => actions.push(action));
    try {
        let response;
        for (const name of ['LiberationSans-Regular.ttf', 'calibri.ttf']) {
            response = await fetch('/font?name=' + name);
            if (response.ok) break;
        }
        if (!response.ok) throw new Error('PDF text requires a TrueType face: set RB_WORD_FONT_DIR to a directory containing LiberationSans-Regular.ttf or calibri.ttf');
        grid.SetPdfFonts([{ family: 'Calibri', bold: false, italic: false, data: new Uint8Array(await response.arrayBuffer()) }]);
        grid.SetCell(0, 0, 'SHELL INITIAL');
        grid.SetPageSetup({ printArea: 'A1:A1' });
        const before = await pdfSummary(await grid.ExportPdf());
        grid.SetCell(0, 0, 'SHELL UPDATED');
        const expected = await pdfSummary(await grid.ExportPdf());
        check('edited workbook changes the PDF text', [before.shown.length > 0, expected.shown.length > 0, JSON.stringify(before.shown) !== JSON.stringify(expected.shown)], [true, true, true]);
        Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', {
            configurable: true,
            get() {
                const real = contentWindow.get.call(this);
                return this.hasAttribute('data-spreadsheet-print')
                    ? { focus: () => real.focus(), print: () => { printCalls++; } }
                    : real;
            }
        });
        shell.Mount(shellHost);
        const status = shellHost.querySelector('[role="status"].sheet-print-status');
        const file = shellHost.querySelector('[data-sheet-menu="file"]');
        const print = shellHost.querySelector('[data-sheet-menu-item="print-pdf"]');
        file.click();
        print.click();
        const pending = status.textContent;
        check('print reports pending and closes File menu', [pending.length > 0, file.getAttribute('aria-expanded')], [true, 'false']);
        print.click();
        check('concurrent clicks invoke the real print engine once', [engineCalls, actions], [1, []]);
        const handoff = await waitForPrintStatus(status, pending);
        check('native print handoff updates shell status', [handoff !== pending, handoff.length > 0, printCalls], [true, true, 1]);
        const frame = document.querySelector('iframe[data-spreadsheet-print]');
        check('PDF blob was handed to one hidden print iframe', [printCalls, !!frame, frame?.style.visibility, /^blob:/.test(frame?.src ?? '')], [1, true, 'hidden', true]);
        if (frame) {
            const printed = new Uint8Array(await (await fetch(frame.src)).arrayBuffer());
            const summary = await pdfSummary(printed);
            check('iframe PDF is the edited workbook with real pages and text',
                [latin1(printed.subarray(0, 5)), summary.pages, summary.shown.length > 0, summary.pages === grid.GetPrintPages(0).pageCount, JSON.stringify(summary.shown) === JSON.stringify(expected.shown)],
                ['%PDF-', expected.pages, true, true, true]);
            frame.remove();
        }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', contentWindow);
        shell.Dispose(); shellHost.remove(); close();
    }
    return lines;
};

// The test runner routes the lazy PDF module request to a genuine network
// failure on a fresh page; no replacement exporter or synthetic PDF is used.
checks['shell-print-pdf-failure'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const shellHost = document.createElement('div');
    document.body.appendChild(shellHost);
    const { createSpreadsheetShell } = await import('/spreadsheet/Toolbar.js');
    const actions = [], rejections = [];
    const onRejection = (event) => rejections.push(String(event.reason));
    window.addEventListener('unhandledrejection', onRejection);
    const shell = createSpreadsheetShell(grid, (action) => actions.push(action));
    try {
        shell.Mount(shellHost);
        const status = shellHost.querySelector('[role="status"].sheet-print-status');
        shellHost.querySelector('[data-sheet-menu="file"]').click();
        const print = shellHost.querySelector('[data-sheet-menu-item="print-pdf"]');
        print.click();
        const pending = status.textContent;
        check('failed module download starts with pending status', pending.length > 0, true);
        const failure = await waitForPrintStatus(status, pending);
        check('failed PDF module request is visible', /^Print PDF failed: .+/.test(failure), true);
        check('failed request does not delegate or print', [actions.length, document.querySelectorAll('iframe[data-spreadsheet-print]').length], [0, 0]);
        print.click();
        const retryPending = status.textContent;
        check('failed request can be retried', [retryPending.length > 0, retryPending !== failure], [true, true]);
        check('failed request releases busy state and catches rejection', [await waitForPrintStatus(status, retryPending), rejections.length], [failure, 0]);
        print.click();
        const detachedPending = status.textContent;
        shellHost.remove();
        await new Promise((resolve) => setTimeout(resolve, 50));
        check('detached shell gets no stale UI update', status.textContent, detachedPending);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        window.removeEventListener('unhandledrejection', onRejection);
        shell.Dispose(); shellHost.remove(); close();
    }
    return lines;
};
