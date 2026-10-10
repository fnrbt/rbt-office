// Browser side of run.mjs, fourth part: features of the Spreadsheet component that are
// persisted in XLSX the way Excel writes them (dynamic arrays, ...). Each check returns a
// list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid } from '/component.js';
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
