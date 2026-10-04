// Browser side of run.mjs for the Spreadsheet web component: change events,
// dispose, sheet access and sizing. Each check returns a list of
// "ok ..." / "FAIL ..." lines.
const main = await import('/office-formats/src/Main.js');
export const OF = typeof main.default === 'function' ? main.default() : main.default;
const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
export { createSpreadsheet };

export const bytesOf = async (name) => new Uint8Array(await (await fetch('/fixture?name=' + encodeURIComponent(name))).arrayBuffer());
export const makeLog = () => {
    const lines = [];
    const check = (label, got, want) => {
        const g = JSON.stringify(got), w = JSON.stringify(want);
        lines.push(g === w ? 'ok   ' + label : `FAIL ${label}: got ${g} want ${w}`);
    };
    return { lines, check, note: (text) => lines.push('ok   (' + text + ')') };
};
export const container = (style) => {
    const div = document.createElement('div');
    div.setAttribute('style', style);
    document.body.appendChild(div);
    return div;
};

// Every listener registered on `document` while `run` executes and still registered afterwards.
const trackDocumentListeners = () => {
    const live = new Map();
    const add = document.addEventListener.bind(document), remove = document.removeEventListener.bind(document);
    document.addEventListener = (type, fn, options) => { live.set(fn, type); add(type, fn, options); };
    document.removeEventListener = (type, fn, options) => { live.delete(fn); remove(type, fn, options); };
    return { live, restore: () => { document.addEventListener = add; document.removeEventListener = remove; } };
};

export const checks = {};

// Item: OnChange / OnSelectionChange / Dispose / GetSheets / container height.
checks['component-host-api'] = async () => {
    const { lines, check } = makeLog();
    const tracker = trackDocumentListeners();
    try {
        const sized = container('position:absolute;left:0;top:0;width:900px;height:420px');
        const grid = createSpreadsheet(sized);
        grid.Mount();
        const gridDiv = sized.querySelector('.spreadsheet-container');
        check('the grid fills a container that has a height', gridDiv.getBoundingClientRect().height, 420);
        sized.style.height = '300px';
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        check('and follows it when it is resized', gridDiv.getBoundingClientRect().height, 300);
        const rowsAt300 = sized.querySelectorAll('.spreadsheet-content > *').length;
        sized.style.height = '700px';
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        check('more rows are rendered in a taller container', sized.querySelectorAll('.spreadsheet-content > *').length > rowsAt300, true);
        const short = container('position:absolute;left:0;top:800px;width:900px;height:90px');
        const shortGrid = createSpreadsheet(short); shortGrid.Mount();
        check('a sensible minimum height', short.querySelector('.spreadsheet-container').getBoundingClientRect().height, 200);
        shortGrid.Dispose(); short.remove();
        const unsized = container('position:absolute;left:0;top:1100px;width:900px');
        const other = createSpreadsheet(unsized); other.Mount();
        check('a container without a height gets the fixed height', unsized.querySelector('.spreadsheet-container').getBoundingClientRect().height, 600);
        other.Dispose(); unsized.remove();

        const changes = [], selections = [];
        const offChange = grid.OnChange((change) => changes.push(change));
        const offSelection = grid.OnSelectionChange((change) => selections.push(change));
        const takeChanges = () => changes.splice(0).map((c) => `${c.Sheets.join(',')}${c.StructureChanged ? ' structure' : ''}`);
        const takeSelections = () => selections.splice(0).map((c) => `${c.SheetIndex}:${c.Selection}`);

        grid.SetCell(0, 0, '12'); grid.SetCell(1, 0, '=A1*2');
        check('cell edits are reported', takeChanges(), ['0', '0']);
        check('they are not selection changes', takeSelections(), []);
        grid.SelectCell(3, 2);
        check('a selection move is reported as one', [takeChanges(), takeSelections()], [[], ['0:C4']]);
        grid.SelectRange(1, 1, 4, 3);
        check('a range selection too', takeSelections(), ['0:B2:D5']);
        grid.Render();
        check('a render with nothing changed reports nothing', [takeChanges(), takeSelections()], [[], []]);
        grid.SelectCell(0, 0); takeSelections();
        grid.ApplyFormat('bold'); check('formatting is a change', takeChanges(), ['0']);
        grid.MergeCells(6, 0, 7, 2); check('a merge is a change', takeChanges(), ['0']);
        grid.InsertRow(0); check('inserting a row is a change', takeChanges(), ['0']);
        grid.Undo(); check('undo is a change', takeChanges(), ['0']);
        grid.Redo(); takeChanges();
        check('GetSheets returns the live models', [grid.GetSheets().length, grid.GetSheets()[0] === grid.GetModel(), grid.GetSheetNames(), grid.GetActiveSheetIndex()], [1, true, ['Sheet1'], 0]);

        // A second sheet: switching is a selection change, edits name the sheet.
        const first = grid.GetSheets()[0];
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([first, createEmpty(100, 10)], ['One', 'Two'], 0, undefined);
        check('installing a workbook is a structure change', takeChanges(), ['1 structure']);
        takeSelections();
        grid.SwitchSheet(1);
        check('switching sheets is a selection change, not a document change', [takeChanges(), takeSelections()], [[], ['1:A1']]);
        grid.SetCell(2, 2, 'on two');
        check('an edit on the second sheet names it', [takeChanges(), grid.GetActiveSheetIndex(), grid.GetSheetNames()], [['1'], 1, ['One', 'Two']]);
        const sheets = grid.GetSheets();
        check('GetSheets has both sheets with their latest edits', [sheets.length, sheets[1] === grid.GetModel(), sheets[0] === first], [2, true, true]);

        offChange();
        grid.SetCell(3, 3, 'unheard');
        check('no change event after unsubscribing', changes.length, 0);
        grid.SelectCell(5, 5);
        check('the other subscription still works', takeSelections(), ['1:F6']);
        offSelection();
        grid.SelectCell(6, 6);
        check('until it is unsubscribed too', selections.length, 0);

        // Dispose with a dialog open: nodes and document listeners all go.
        grid.OnChange((change) => changes.push(change));
        grid.ShowFindReplaceDialog();
        check('listeners are on the document while mounted', tracker.live.size >= 3, true);
        check('the dialog is in the page', !!document.getElementById('find-replace-dialog'), true);
        grid.Dispose();
        check('dispose removes every document listener', [...tracker.live.values()], []);
        check('dispose removes the grid and its dialogs', [sized.children.length, !!document.getElementById('find-replace-dialog'), document.querySelectorAll('.spreadsheet-container').length], [0, false, 0]);
        grid.SetCell(0, 0, 'after dispose');
        check('no events after dispose', changes.length, 0);
        document.dispatchEvent(new MouseEvent('mousemove', { clientX: 5, clientY: 5 }));
        document.dispatchEvent(new MouseEvent('mouseup'));
        sized.remove();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    tracker.restore();
    return lines;
};

window.componentChecks = checks;

// ---- ExportXlsx ----------------------------------------------------------------

const excel = OF.excel;
const sum = (bytes) => { let s = 0; for (const b of bytes) s = (s + b) | 0; return s; };
// What the reader sees of a workbook, sheet by sheet.
const describeWorkbook = (wb) => excel.sheetNames(wb).map((name, index) => {
    const sheet = excel.sheet(name, wb), meta = excel.sheetMeta(wb, sheet), cells = excel.cellsColumnar(wb, sheet);
    return {
        name, state: excel.sheetStates(wb)[index],
        cells: cells.rows.map((row, i) => [row, cells.cols[i], cells.values[i], cells.formulas[i], cells.errors[i] ?? null, JSON.stringify(excel.styleAt(cells.styleIndexes[i], wb))]),
        merged: [...excel.mergedCells(sheet)].sort(),
        widths: Array.from(excel.columnWidths(Math.max(1, excel.dimensions(sheet).colCount), sheet), (w) => Math.round(w * 100) / 100),
        heights: Array.from(excel.rowHeights(Math.max(1, excel.dimensions(sheet).rowCount), sheet), (h) => Math.round(h * 100) / 100),
        frozen: [meta.frozenRows, meta.frozenCols], grid: meta.showGridLines, hiddenRows: meta.hiddenRows, hiddenCols: meta.hiddenCols,
        tables: meta.tables, validations: meta.validations, conditionalFormats: meta.conditionalFormats, hyperlinks: meta.hyperlinks,
        page: [meta.pageSetup, meta.pageMargins, meta.printOptions, meta.headerFooter, meta.printArea, meta.printTitleRows, meta.printTitleCols, meta.rowBreaks, meta.colBreaks],
        comments: Array.from(excel.comments(sheet), (c) => [c.CellRef.Column + c.CellRef.Row, c.Text, c.Author ?? null]),
        images: excel.images(sheet).map((image) => ({ anchor: image.anchor, name: image.name ?? null, type: image.contentType, bytes: image.data.length, sum: sum(image.data) })),
        objects: excel.drawingObjects(sheet).map((item) => ({ anchor: item.anchor, kind: item.kind, name: item.name ?? null })),
        threads: excel.threadedComments(sheet),
    };
});
const names = (wb) => excel.workbookInfo(wb).definedNames.map((n) => [n.name.toUpperCase(), n.refersTo, n.localSheetId ?? null]).sort();
const partsOf = async (bytes, pattern) => {
    const zip = await window.JSZip.loadAsync(bytes), out = {};
    for (const entry of Object.keys(zip.files)) if (pattern.test(entry) && !zip.files[entry].dir) out[entry] = await zip.files[entry].async('base64');
    return out;
};
// Parts the editor keeps as they are, by path; pictures are written under the writer's own names.
const KEPT_PARTS = /^xl\/(charts|theme|embeddings)\//, MEDIA_PARTS = /^xl\/media\//;
const changedParts = (before, after) => Object.keys(before).filter((entry) => before[entry] !== after[entry]).map((entry) => entry + (after[entry] === undefined ? ' (missing)' : ' (changed)'));
// Pictures used by a sheet, as a sorted list of (length, checksum); an unused media part of the file is not carried.
const mediaContents = (parts) => [...new Set(Object.values(parts).map((data) => data.length + ':' + [...data].reduce((n, ch) => (n * 31 + ch.charCodeAt(0)) | 0, 7)))].sort();
// The <col> elements of every worksheet: how many, and how many columns carry a default style.
const colElements = async (bytes) => {
    const zip = await window.JSZip.loadAsync(bytes);
    let count = 0, styledColumns = 0;
    for (const entry of Object.keys(zip.files)) {
        if (!/^xl\/worksheets\/[^/]+\.xml$/.test(entry)) continue;
        const xml = await zip.files[entry].async('string');
        for (const col of xml.match(/<col [^>]*>/g) ?? []) {
            count++;
            const style = /style="(\d+)"/.exec(col), min = Number(/min="(\d+)"/.exec(col)[1]), max = Number(/max="(\d+)"/.exec(col)[1]);
            if (style && style[1] !== '0') styledColumns += max - min + 1;
        }
    }
    return { count, styledColumns };
};
// Pictures and kept objects of every sheet, back to front.
const zOrder = (wb) => excel.sheetNames(wb).map((name) => {
    const sheet = excel.sheet(name, wb);
    return [...Array.from(sheet.Extras.Images, (image) => [image.Anchor.Order, 'picture']), ...Array.from(sheet.Extras.Objects, (item) => [item.Anchor.Order, item.Kind])].sort((x, y) => x[0] - y[0]).map((entry) => entry[1]);
});
// { value, formula, error, bold, fillColor, ... } of a cell, or undefined.
const cellAt = (wb, sheetName, ref) => {
    const sheet = excel.sheet(sheetName, wb), cells = excel.cellsColumnar(wb, sheet);
    const match = /^([A-Z]+)(\d+)$/.exec(ref);
    const col = [...match[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, row = Number(match[2]) - 1;
    for (let i = 0; i < cells.rows.length; i++)
        if (cells.rows[i] === row && cells.cols[i] === col)
            return { ...(excel.styleAt(cells.styleIndexes[i], wb) ?? {}), value: cells.values[i], formula: cells.formulas[i], error: cells.errors[i] };
    return undefined;
};
const firstDifference = (a, b, path = '') => {
    if (JSON.stringify(a) === JSON.stringify(b)) return null;
    if (a && b && typeof a === 'object' && typeof b === 'object')
        for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) { const d = firstDifference(a[key], b[key], path + '/' + key); if (d) return d; }
    return `${path}: ${String(JSON.stringify(a)).slice(0, 160)} -> ${String(JSON.stringify(b)).slice(0, 160)}`;
};
export const mountGrid = () => {
    const host = container('position:absolute;left:0;top:0;width:1100px;height:600px');
    const grid = createSpreadsheet(host);
    grid.Mount();
    return { grid, host, close: () => { grid.Dispose(); host.remove(); } };
};

// An unedited workbook is exported as the file had it: every sheet, cell, format, size,
// merge, table, rule, name, picture and chart reads back the same.
const UNEDITED = ['tests/real-world/milk-street-office.xlsx', 'tests/real-world/chart-zoo.xlsx', 'tests/real-world/saopaulo-condos.xlsx',
    'tests/real-world/conditional-formatting-tutorial.xlsx', 'tests/test-documents-excel/phpsp-27template.xlsx', 'tests/test-documents-excel/poi-47668.xlsx',
    'tests/test-documents-excel/phpsp-32chartreadwrite.xlsx'];
checks['export-unedited'] = async () => {
    const { lines, check, note } = makeLog();
    for (const file of UNEDITED) {
        const short = file.split('/').pop();
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf(file);
            await grid.ImportXlsx(original.slice());
            check(`${short}: the grid knows its source`, grid.HasXlsxSource(), true);
            const saved = await grid.ExportXlsx();
            check(`${short}: bytes come back`, saved instanceof Uint8Array && saved.length > 1000, true);
            const before = await excel.load(original.slice()), after = await excel.load(saved.slice());
            const a = describeWorkbook(before), b = describeWorkbook(after);
            // Cells that hold nothing (no value, default style) are not kept by the editor.
            const defaultStyle = JSON.stringify(excel.styleAt(0, before));
            const meaningful = (list) => list.map((sheet) => ({ ...sheet, cells: sheet.cells.filter((c) => c[2] != null || c[3] != null || c[4] != null || (c[5] !== '{}' && c[5] !== 'null' && c[5] !== defaultStyle)) }));
            check(`${short}: reads back as the file`, firstDifference(meaningful(a), meaningful(b)), null);
            check(`${short}: defined names`, names(after), names(before));
            check(`${short}: chart, theme and embedded parts byte for byte`, changedParts(await partsOf(original, KEPT_PARTS), await partsOf(saved, KEPT_PARTS)), []);
            check(`${short}: every media part is in the saved package`, mediaContents(await partsOf(saved, MEDIA_PARTS)), mediaContents(await partsOf(original, MEDIA_PARTS)));
            // <col> definitions: ranges are not split into one element per column, default cell styles are kept.
            const colsBefore = await colElements(original), colsAfter = await colElements(saved);
            check(`${short}: <col> ranges are not split (${colsBefore.count} -> ${colsAfter.count}), column styles kept`, [colsAfter.count <= colsBefore.count, colsAfter.styledColumns], [true, colsBefore.styledColumns]);
            check(`${short}: z-order of pictures and kept objects`, zOrder(after), zOrder(before));
            note(`${short}: ${a.length} sheets, ${a.reduce((n, s) => n + s.cells.length, 0)} cells, ${a.reduce((n, s) => n + s.images.length, 0)} pictures, ${a.reduce((n, s) => n + s.objects.length, 0)} kept objects, ${saved.length} bytes`);
        } catch (e) { lines.push(`FAIL ${short} threw ` + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    return lines;
};

// Edits made in the grid - cells, a format, an inserted row, a merge, sizes, a second
// sheet - are in the exported file, next to everything the file had.
checks['export-edited'] = async () => {
    const { lines, check, note } = makeLog();
    for (const [file, flow] of [['tests/real-world/milk-street-office.xlsx', 'ImportXlsx'], ['tests/real-world/chart-zoo.xlsx', 'host'], ['tests/real-world/saopaulo-condos.xlsx', 'ImportXlsx'], ['tests/test-documents-excel/phpsp-27template.xlsx', 'host']]) {
        const short = file.split('/').pop();
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf(file);
            if (flow === 'ImportXlsx') await grid.ImportXlsx(original.slice());
            else {
                // The flow hosts have today: the Office.Formats bundle reads, the importer converts.
                const { importWorkbook } = await import('/spreadsheet/XlsxImporter.js');
                grid.InstallImportedWorkbook(importWorkbook(await excel.load(original.slice())), -1);
            }
            const before = await excel.load(original.slice());
            const sheetNames = grid.GetSheetNames(), active = grid.GetActiveSheetIndex(), activeName = sheetNames[active];
            const a = describeWorkbook(before), sheetBefore = a[active];
            // A free column right of the data and a cell with a value below row 6 to follow through the row insert.
            const freeCol = Math.max(0, ...sheetBefore.cells.map((c) => c[1])) + 2;
            const letter = (await import('/spreadsheet/Model.js')).columnToLetter(freeCol);
            const tracked = sheetBefore.cells.find((c) => c[0] >= 6 && c[2] != null && c[3] == null);
            const changes = [];
            grid.OnChange((change) => changes.push(change));
            grid.SetCell(0, freeCol, 'exported text');
            grid.SetCell(1, freeCol, '1234.5');
            grid.SetCell(2, freeCol, `=${letter}2*2`);
            grid.SelectCell(0, freeCol); grid.ApplyFormat('bold'); grid.ApplyBgColor('#FFEE00');
            grid.SetComment(1, freeCol, 'a note', 'Tester');
            grid.InsertRow(4);
            grid.MergeCells(0, freeCol + 1, 1, freeCol + 3);
            const widthBefore = excel.columnWidth(freeCol, excel.sheet(activeName, before));
            // A second sheet, when there is one that shows a tab.
            const visible = grid.SheetVisibility();
            const second = sheetNames.findIndex((_, i) => i !== active && visible[i]);
            if (second >= 0) { grid.SwitchSheet(second); grid.SetCell(0, 40, 'on the second sheet'); grid.SwitchSheet(active); }
            check(`${short}: the edits were reported`, changes.length >= 6, true);

            const saved = await grid.ExportXlsx();
            const after = await excel.load(saved.slice());
            const b = describeWorkbook(after), sheetAfter = b[active];
            check(`${short}: same sheets, same visibility`, b.map((s) => [s.name, s.state]), a.map((s) => [s.name, s.state]));
            const text = cellAt(after, activeName, `${letter}1`);
            check(`${short}: text with its format`, [text.value, text.bold === true, (text.fillColor || '').toUpperCase().replace('#', '').slice(-6)], ['exported text', true, 'FFEE00']);
            check(`${short}: number`, cellAt(after, activeName, `${letter}2`)?.value, 1234.5);
            const formula = cellAt(after, activeName, `${letter}3`);
            check(`${short}: formula with its cached value`, [formula.formula, formula.value], [`${letter}2*2`, 2469]);
            check(`${short}: comment`, sheetAfter.comments.filter((c) => c[0] === `${letter}2`).map((c) => [c[1], c[2]]), [['a note', 'Tester']]);
            check(`${short}: merge`, sheetAfter.merged.includes(`${(await import('/spreadsheet/Model.js')).columnToLetter(freeCol + 1)}1:${(await import('/spreadsheet/Model.js')).columnToLetter(freeCol + 3)}2`), true);
            check(`${short}: the file's merges are all still there`, sheetAfter.merged.length, sheetBefore.merged.length + 1);
            if (tracked) {
                const moved = sheetAfter.cells.find((c) => c[0] === tracked[0] + 1 && c[1] === tracked[1]);
                check(`${short}: the inserted row moved the cells below it`, moved && [moved[2], moved[5]], [tracked[2], tracked[5]]);
            }
            check(`${short}: row 5 is the new, empty row`, sheetAfter.cells.filter((c) => c[0] === 4 && (c[2] != null || c[3] != null)).length, 0);
            check(`${short}: as many cells with content as before, plus the three typed`, sheetAfter.cells.filter((c) => c[2] != null || c[3] != null || c[4] != null).length, sheetBefore.cells.filter((c) => c[2] != null || c[3] != null || c[4] != null).length + 3);
            check(`${short}: formulas kept`, sheetAfter.cells.filter((c) => c[3] != null).length, sheetBefore.cells.filter((c) => c[3] != null).length + 1);
            check(`${short}: freeze panes, gridlines`, [sheetAfter.grid, sheetAfter.frozen[1]], [sheetBefore.grid, sheetBefore.frozen[1]]);
            check(`${short}: tables, validations, conditional formats, hyperlinks (counts)`, [sheetAfter.tables.length, sheetAfter.validations.length, sheetAfter.conditionalFormats.length, sheetAfter.hyperlinks.length],
                [sheetBefore.tables.length, sheetBefore.validations.length, sheetBefore.conditionalFormats.length, sheetBefore.hyperlinks.length]);
            check(`${short}: page setup`, [sheetAfter.page[0], sheetAfter.page[1], sheetAfter.page[2], sheetAfter.page[3]], [sheetBefore.page[0], sheetBefore.page[1], sheetBefore.page[2], sheetBefore.page[3]]);
            check(`${short}: defined names (count)`, names(after).length, names(before).length);
            check(`${short}: hidden rows and columns follow the insert`, [sheetAfter.hiddenRows.length, sheetAfter.hiddenCols], [sheetBefore.hiddenRows.length, sheetBefore.hiddenCols]);
            check(`${short}: an untouched column keeps its width`, excel.columnWidth(0, excel.sheet(activeName, after)), excel.columnWidth(0, excel.sheet(activeName, before)));
            if (second >= 0) {
                check(`${short}: the edit on the second sheet`, cellAt(after, sheetNames[second], 'AO1')?.value, 'on the second sheet');
                // The other sheets read back as the file had them, but for that one cell.
                const strip = (sheet) => ({ ...sheet, widths: sheet.widths.slice(0, 40), cells: sheet.cells.filter((c) => !(c[0] === 0 && c[1] === 40) && (c[2] != null || c[3] != null || c[4] != null)), images: sheet.images.length, objects: sheet.objects.length });
                for (let i = 0; i < a.length; i++) if (i !== active) {
                    // Formulas on other sheets that point into the edited sheet follow the inserted row.
                    // (The row insert recalculates the workbook: a recomputed number may differ from Excel's in its last digits.)
                    const digits = (value) => (typeof value === 'number' ? Number(value.toPrecision(11)) : value);
                    const formulaFree = (sheet) => ({ ...strip(sheet), cells: strip(sheet).cells.map((c) => [c[0], c[1], digits(c[2]), c[3] == null ? null : 'f', c[4], c[5]]) });
                    const difference = firstDifference(formulaFree(a[i]), formulaFree(b[i]));
                    if (difference) { const at = Number(/cells\/(\d+)/.exec(difference)?.[1] ?? -1); lines.push('ok   (detail ' + JSON.stringify(strip(a[i]).cells[at]) + ' -> ' + JSON.stringify(strip(b[i]).cells[at]) + ')'); }
                    check(`${short}: sheet ${a[i].name} is unchanged`, difference, null);
                }
            }
            check(`${short}: pictures`, b.map((s) => s.images.map((i) => [i.type, i.bytes, i.sum])), a.map((s) => s.images.map((i) => [i.type, i.bytes, i.sum])));
            check(`${short}: charts and other kept objects`, b.map((s) => s.objects.map((o) => [o.kind, o.name])), a.map((s) => s.objects.map((o) => [o.kind, o.name])));
            check(`${short}: threaded comments`, b.map((s) => s.threads.length), a.map((s) => s.threads.length));
            check(`${short}: chart, theme and embedded parts byte for byte`, changedParts(await partsOf(original, KEPT_PARTS), await partsOf(saved, KEPT_PARTS)), []);
            check(`${short}: every media part is in the saved package`, mediaContents(await partsOf(saved, MEDIA_PARTS)), mediaContents(await partsOf(original, MEDIA_PARTS)));
            // What was exported opens in the grid again, with the edits.
            await grid.ImportXlsx(saved.slice());
            const { getCellValue, CellAddress, displayValue } = await import('/spreadsheet/Model.js');
            const model = grid.GetSheets()[active];
            check(`${short}: the export imports again`, [displayValue(getCellValue(new CellAddress(0, freeCol), model)), grid.GetSheetNames()], ['exported text', sheetNames]);
            note(`${short} (${flow}): ${saved.length} bytes`);
        } catch (e) { lines.push(`FAIL ${short} threw ` + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    // A workbook made in the grid, with no file behind it.
    const { grid, close } = mountGrid();
    try {
        grid.SetCell(0, 0, 'Name'); grid.SetCell(1, 0, '3'); grid.SetCell(2, 0, '=A2*2');
        check('new workbook: no source', grid.HasXlsxSource(), false);
        const after = await excel.load((await grid.ExportXlsx()).slice());
        check('new workbook: exports and reads back', [excel.sheetNames(after), cellAt(after, 'Sheet1', 'A1').value, cellAt(after, 'Sheet1', 'A3').formula, cellAt(after, 'Sheet1', 'A3').value], [['Sheet1'], 'Name', 'A2*2', 6]);
    } catch (e) { lines.push('FAIL new workbook threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
