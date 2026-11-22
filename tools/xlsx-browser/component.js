// Browser side of run.mjs for the Spreadsheet web component: change events,
// dispose, sheet access and sizing. Each check returns a list of
// "ok ..." / "FAIL ..." lines.
const main = await import('/office-formats/src/Main.js');
export const OF = typeof main.default === 'function' ? main.default() : main.default;
const { default: createRawSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');

// What is wrong with a written package: a relationship whose target is not in it, a part
// without a content type, a content type for a part that is not there, a part nothing leads to.
export const packageProblems = async (bytes) => {
    const zip = await window.JSZip.loadAsync(bytes), problems = [];
    const paths = Object.keys(zip.files).filter((path) => !zip.files[path].dir), has = new Set(paths);
    const attribute = (tag, name) => { const m = new RegExp('\\s' + name + '="([^"]*)"').exec(tag); return m ? m[1].replace(/&amp;/g, '&') : null; };
    const types = has.has('[Content_Types].xml') ? await zip.files['[Content_Types].xml'].async('string') : '';
    const defaults = new Set([...types.matchAll(/<Default\b[^>]*>/g)].map((m) => (attribute(m[0], 'Extension') ?? '').toLowerCase()));
    const overrides = [...types.matchAll(/<Override\b[^>]*>/g)].map((m) => (attribute(m[0], 'PartName') ?? '').replace(/^\//, ''));
    for (const path of overrides) if (!has.has(path)) problems.push('content type for a missing part: ' + path);
    if (new Set(overrides).size !== overrides.length) problems.push('a part has two content types');
    const overridden = new Set(overrides), reached = new Set();
    const resolve = (base, target) => {
        const out = [];
        for (const segment of (target.startsWith('/') || base === '' ? target : base + '/' + target).split('/')) { if (segment === '..') out.pop(); else if (segment !== '' && segment !== '.') out.push(segment); }
        return out.join('/');
    };
    // What real Excel refuses a whole file for (each found by opening saves in Excel).
    const errors = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!', '#BUSY!']);
    for (const path of paths.filter((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p))) {
        const xml = await zip.files[path].async('string');
        for (const m of xml.matchAll(/<col\b[^>]*>/g)) if (Number(attribute(m[0], 'max')) > 16384) problems.push(`${path}: a column past the last: ${m[0]}`);
        for (const m of xml.matchAll(/<c\b[^>]*\bt="e"[^>]*>(?:<f\b[^>]*\/>|<f\b[^>]*>[^<]*<\/f>)?<v>([^<]*)<\/v>/g)) if (!errors.has(m[1].replace(/&amp;/g, '&'))) problems.push(`${path}: the error ${m[1]} is not one of the file format`);
        for (const m of xml.matchAll(/<f\b[^>]*\bt="shared"[^>]*>([^<]+)<\/f>/g)) if (m[1].includes('!')) problems.push(`${path}: a shared formula reads another sheet: ${m[1]}`);
        for (const m of xml.matchAll(/<f\b[^>]*>([^<]+)<\/f>/g)) {
            const bare = /(?:^|[^A-Za-z0-9_.\\])(XLOOKUP|XMATCH|FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|LET|LAMBDA|IFS|MAXIFS|MINIFS|CONCAT|TEXTJOIN|SWITCH|IFNA|STDEV\.S|VAR\.S)\(/i.exec(m[1].replace(/&quot;/g, '"').replace(/"[^"]*"|'[^']*'/g, '""'));
            if (bare) problems.push(`${path}: ${bare[1]} without its _xlfn. prefix: ${m[1]}`);
        }
    }
    for (const path of paths.filter((p) => /^xl\/tables\/[^/]+\.xml$/.test(p))) {
        const xml = await zip.files[path].async('string');
        const columns = [...xml.matchAll(/<tableColumn\b[^>]*>/g)].map((m) => m[0]);
        const names = columns.map((tag) => (attribute(tag, 'name') ?? '').toLowerCase());
        if (new Set(names).size !== names.length) problems.push(`${path}: two columns of one name`);
        const relations = paths.includes(path.replace('tables/', 'tables/_rels/') + '.rels') ? await zip.files[path.replace('tables/', 'tables/_rels/') + '.rels'].async('string') : '';
        const query = /Target="\.\.\/(queryTables\/[^"]+)"/.exec(relations);
        if (query) {
            const fields = [...(await zip.files['xl/' + query[1]].async('string')).matchAll(/<queryTableField\b[^>]*>/g)].map((m) => attribute(m[0], 'tableColumnId'));
            if (fields.join() !== columns.map((tag) => attribute(tag, 'id')).join()) problems.push(`${path}: the fields of its query are not its columns`);
        }
    }
    for (const path of paths) {
        if (path.endsWith('.rels')) {
            const directory = path.slice(0, Math.max(0, path.lastIndexOf('/'))), ownerDirectory = directory.replace(/\/?_rels$/, '');
            const owner = (ownerDirectory ? ownerDirectory + '/' : '') + path.slice(path.lastIndexOf('/') + 1).replace(/\.rels$/, '');
            if (owner !== '' && !has.has(owner)) problems.push('relationships of a missing part: ' + path);
            const ids = new Set();
            for (const m of (await zip.files[path].async('string')).matchAll(/<Relationship\b[^>]*>/g)) {
                const id = attribute(m[0], 'Id');
                if (ids.has(id)) problems.push(`${path}: relationship id ${id} twice`);
                ids.add(id);
                // External, or a place in the workbook (a drawing's link to `#Sheet1!A1`): no part.
                if (attribute(m[0], 'TargetMode') === 'External' || (attribute(m[0], 'Target') ?? '').startsWith('#')) continue;
                const target = resolve(ownerDirectory, attribute(m[0], 'Target') ?? '');
                reached.add(target);
                if (!has.has(target)) problems.push(`${path}: target ${target} is missing`);
            }
        } else if (path !== '[Content_Types].xml') {
            const extension = path.includes('.') ? path.slice(path.lastIndexOf('.') + 1).toLowerCase() : '';
            if (!overridden.has(path) && !defaults.has(extension)) problems.push('no content type: ' + path);
        }
    }
    for (const path of paths) if (!path.endsWith('.rels') && path !== '[Content_Types].xml' && !reached.has(path)) problems.push('orphan: ' + path);
    return problems;
};
// Every workbook a check exports is checked for integrity; run.mjs reports the problems with the check.
window.exportProblems = [];
window.exportsChecked = 0;
// A problem the opened file had already (pivot parts nothing leads to in a synthetic fixture) is not the save's.
export const createSpreadsheet = (element) => {
    const grid = createRawSpreadsheet(element), exportXlsx = grid.ExportXlsx.bind(grid), importXlsx = grid.ImportXlsx.bind(grid);
    let inherited = new Set();
    grid.ImportXlsx = async (bytes, ...rest) => {
        try { inherited = new Set(await packageProblems(bytes)); } catch { inherited = new Set(); }
        return importXlsx(bytes, ...rest);
    };
    grid.ExportXlsx = async () => {
        const bytes = await exportXlsx();
        window.exportsChecked++;
        for (const problem of await packageProblems(bytes)) if (!inherited.has(problem)) window.exportProblems.push(problem);
        return bytes;
    };
    return grid;
};

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
        check('the grid owns its find bar, without a global dialog', [sized.querySelectorAll('.sheet-find-bar').length, !!document.getElementById('find-replace-dialog')], [1, false]);
        grid.Dispose();
        check('dispose removes every document listener', [...tracker.live.values()], []);
        check('dispose removes the grid and its find bar', [sized.children.length, document.querySelectorAll('.sheet-find-bar').length, document.querySelectorAll('.spreadsheet-container').length], [0, 0, 0]);
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

// A copy / cut / paste event carrying a DataTransfer of the check's. (Firefox gives a synthetic
// ClipboardEvent a DataTransfer of its own, without the data of its init; the event's own property is the one handlers read.)
const clipboardEvent = (type, clipboardData) => {
    const event = new ClipboardEvent(type, { clipboardData, bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    return event;
};
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
export const partsOf = async (bytes, pattern) => {
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
            // Structural edits move chart bindings; chart-reference-edits covers
            // their plotted values and save/reopen behavior, not byte identity.
            const keptBefore = await partsOf(original, KEPT_PARTS), keptAfter = await partsOf(saved, KEPT_PARTS);
            for (const path of Object.keys(keptBefore)) {
                if (/^xl\/charts\/chart\d+\.xml$/.test(path)) {
                    delete keptBefore[path]; delete keptAfter[path];
                }
            }
            check(`${short}: other chart, theme and embedded parts byte for byte`, changedParts(keptBefore, keptAfter), []);
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

// What the editor does not model but a save must not lose: pivot tables, pivot caches,
// slicers, slicer caches and external links go back byte for byte, with the relationships
// and content types that make them part of the workbook; a pivot table follows inserted
// rows; sheets keep that through a rename / reorder / added sheet.
const CARRIED_PARTS = /^xl\/(pivotTables|pivotCache|slicers|slicerCaches|externalLinks)\//;
export const textParts = async (bytes, pattern) => {
    const zip = await window.JSZip.loadAsync(bytes), out = {};
    for (const entry of Object.keys(zip.files)) if (pattern.test(entry) && !zip.files[entry].dir) out[entry] = await zip.files[entry].async('string');
    return out;
};
const countOf = (text, pattern) => (text.match(pattern) ?? []).length;
checks['export-kept-parts'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const file = 'tests/real-world/dashboard-sales-analysis.xlsx';
        const original = await bytesOf(file);
        const loadStarted = performance.now();
        await grid.ImportXlsx(original.slice());
        const loadTime = performance.now() - loadStarted;
        const before = await partsOf(original, CARRIED_PARTS);
        check('the file has 8 pivot tables, a pivot cache, slicers and external links', [countOf(Object.keys(before).join('\n'), /pivotTables\/pivotTable\d+\.xml$/gm), countOf(Object.keys(before).join('\n'), /slicerCaches\/slicerCache\d+\.xml$/gm), countOf(Object.keys(before).join('\n'), /externalLinks\/externalLink\d+\.xml$/gm)], [8, 3, 2]);
        // An edit, then a save.
        grid.SetCell(40, 0, 'edited');
        const exportStarted = performance.now();
        const saved = await grid.ExportXlsx();
        const exportTime = performance.now() - exportStarted;
        const after = await partsOf(saved, CARRIED_PARTS);
        check('pivot / slicer / external-link parts and their .rels byte for byte', changedParts(before, after), []);
        const meta = await textParts(saved, /^(\[Content_Types\]\.xml|xl\/workbook\.xml|xl\/_rels\/workbook\.xml\.rels|xl\/worksheets\/_rels\/.*\.rels|xl\/worksheets\/sheet1\.xml)$/);
        const types = meta['[Content_Types].xml'];
        check('every kept part has its content type', Object.keys(before).filter((entry) => entry.endsWith('.xml') && !types.includes(`PartName="/${entry}"`)), []);
        const workbookXml = meta['xl/workbook.xml'], workbookRels = meta['xl/_rels/workbook.xml.rels'];
        check('workbook.xml lists the pivot cache, external references and slicer caches', [countOf(workbookXml, /<pivotCache /g), countOf(workbookXml, /<externalReference /g), countOf(workbookXml, /slicerCache r:id=/g)], [1, 2, 3]);
        check('workbook relationships to them', [countOf(workbookRels, /\/pivotCacheDefinition"/g), countOf(workbookRels, /\/slicerCache"/g), countOf(workbookRels, /\/externalLink"/g)], [1, 3, 2]);
        const ids = [...workbookXml.matchAll(/r:id="([^"]*)"/g)].map((m) => m[1]);
        check('every r:id of workbook.xml is a relationship, none twice', [ids.filter((id) => !workbookRels.includes(`Id="${id}"`)), ids.length - new Set(ids).size], [[], 0]);
        const originalWorkbook = (await textParts(original, /^xl\/workbook\.xml$/))['xl/workbook.xml'];
        const sheetIds = (xml) => [...xml.matchAll(/<sheet name="([^"]*)" sheetId="(\d+)"/g)].map((m) => m[1] + '=' + m[2]);
        check('sheets keep their sheetId (slicer caches name pivot tables by it)', sheetIds(workbookXml), sheetIds(originalWorkbook));
        const sheetRels = Object.entries(meta).filter(([entry]) => entry.startsWith('xl/worksheets/_rels/')).map(([, xml]) => xml).join('');
        check('sheet relationships to pivot tables and the slicer part', [countOf(sheetRels, /\/pivotTable"/g), countOf(sheetRels, /relationships\/slicer"/g)], [8, 1]);
        const slicerId = /x14:slicer r:id="([^"]*)"/.exec(meta['xl/worksheets/sheet1.xml'])?.[1];
        check('the sheet lists its slicers, under a relationship of the sheet', [!!slicerId, (meta['xl/worksheets/_rels/sheet1.xml.rels'] ?? '').includes(`Id="${slicerId}"`)], [true, true]);
        check('the written workbook opens again with its sheets', Array.from(await (await import('/spreadsheet/XlsxImporter.js')).getSheetNamesPromise(saved.slice())), grid.GetSheetNames());
        note(`dashboard: load ${Math.round(loadTime)} ms, export ${Math.round(exportTime)} ms, ${saved.length} bytes`);

        // Two rows inserted above a pivot table: only its <location ref> changes.
        const names = grid.GetSheetNames(), linePivot = names.indexOf('Line Pivot');
        grid.SwitchSheet(linePivot);
        const pivotPath = grid.GetModel().KeptRanges.head.Key;
        const locationOf = (xml) => /<location [^>]*ref="([^"]*)"/.exec(xml)?.[1];
        const pivotBefore = (await textParts(original, CARRIED_PARTS))[pivotPath];
        grid.InsertRow(0); grid.InsertRow(0);
        const moved = await grid.ExportXlsx();
        const pivotAfter = (await textParts(moved, CARRIED_PARTS))[pivotPath];
        check('a pivot table follows rows inserted above it', [locationOf(pivotBefore), locationOf(pivotAfter)], ['A3:D27', 'A5:D29']);
        check('and the rest of its part is untouched', pivotAfter.replace('ref="A5:D29"', 'ref="A3:D27"') === pivotBefore, true);
        check('the other kept parts too', changedParts(before, await partsOf(moved, CARRIED_PARTS)).filter((entry) => !entry.startsWith(pivotPath)), []);
        grid.Undo(); grid.Undo();
        check('undo puts the pivot table back', locationOf((await textParts(await grid.ExportXlsx(), CARRIED_PARTS))[pivotPath]), 'A3:D27');

        // The host renames, reorders and adds sheets: each sheet still has what the file gave it.
        const { createEmpty } = await import('/spreadsheet/Model.js');
        const sheets = grid.GetSheets();
        grid.InstallWorkbook([createEmpty(50, 10), ...[...sheets].reverse()], ['Added', ...[...names].reverse().map((name) => name + ' 2')], 0, undefined);
        check('the grid still knows the file after sheets were renamed, reordered and one added', grid.HasXlsxSource(), true);
        const shuffled = await grid.ExportXlsx();
        check('kept parts byte for byte after rename / reorder / add', changedParts(before, await partsOf(shuffled, CARRIED_PARTS)), []);
        const reread = await excel.load(shuffled.slice());
        const pivotsPerSheet = (wb) => Object.fromEntries(excel.sheetNames(wb).map((name) => [name.replace(/ 2$/, ''), excel.sheet(name, wb).Extras.PivotLocations.length ?? [...excel.sheet(name, wb).Extras.PivotLocations].length]));
        const expectedPivots = { ...pivotsPerSheet(await excel.load(original.slice())), Added: 0 };
        const sorted = (counts) => Object.entries(counts).sort((x, y) => x[0].localeCompare(y[0]));
        check('each sheet has its own pivot tables', sorted(pivotsPerSheet(reread)), sorted(expectedPivots));
        // A different workbook has nothing of this file.
        grid.InstallWorkbook([createEmpty(50, 10)], ['Fresh'], 0, undefined);
        check('a different workbook has no source', grid.HasXlsxSource(), false);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// What-if data tables are written back as data tables (they used to become plain values),
// keep the results the file had, and follow structure edits.
checks['export-data-tables'] = async () => {
    const { lines, check } = makeLog();
    const tablesOf = async (bytes) => {
        const sheets = await textParts(bytes, /^xl\/worksheets\/[^/]+\.xml$/), found = [];
        for (const xml of Object.values(sheets))
            for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*><f ([^>]*t="dataTable"[^>]*)\/><v>([^<]*)<\/v>/g)) {
                const attrs = Object.fromEntries([...m[2].matchAll(/([A-Za-z0-9]+)="([^"]*)"/g)].map((a) => [a[1], a[2]]).filter(([name]) => name !== 'aca'));
                found.push([m[1], attrs.ref, attrs.dt2D ?? '0', attrs.dtr ?? '0', attrs.r1, attrs.r2 ?? '', Number(m[3])]);
            }
        return found.sort((a, b) => a[0].localeCompare(b[0]));
    };
    for (const [file, sheet, expected] of [['tests/real-world/milk-street-office.xlsx', 'Annual', 10], ['tests/real-world/saopaulo-condos.xlsx', 'Assumptions', 8]]) {
        const short = file.split('/').pop();
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf(file);
            await grid.ImportXlsx(original.slice());
            const before = await tablesOf(original);
            check(`${short}: the file has ${expected} data tables`, before.length, expected);
            const index = grid.GetSheetNames().indexOf(sheet);
            grid.SwitchSheet(index);
            check(`${short}: they are in the model`, [...grid.GetModel().DataTables].length, expected);
            // An edit elsewhere (a full recalculation happens on the row insert below).
            grid.SetCell(0, 60, '5');
            check(`${short}: written back as data tables with the file's results`, firstDifference(await tablesOf(await grid.ExportXlsx()), before), null);
            grid.InsertRow(0);
            const shift = (ref) => ref.replace(/(\$?[A-Z]+\$?)(\d+)/g, (_, col, row) => col + (Number(row) + 1));
            const after = await tablesOf(await grid.ExportXlsx());
            check(`${short}: range and input cells follow an inserted row`, firstDifference(after.map((t) => t.slice(0, 6)), before.map((t) => [shift(t[0]), shift(t[1]), t[2], t[3], shift(t[4]), shift(t[5])]).sort((a, b) => a[0].localeCompare(b[0]))), null);
            check(`${short}: the results are still the file's after the recalculation`, after.map((t) => t[6]).sort(), before.map((t) => t[6]).sort());
        } catch (e) { lines.push(`FAIL ${short} threw ` + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    return lines;
};

// Hyperlinks and column default styles follow inserted and deleted rows / columns.
checks['links-and-column-styles-follow-edits'] = async () => {
    const { lines, check } = makeLog();
    {
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf('tests/test-documents-excel/synthetic-formatting.xlsx');
            await grid.ImportXlsx(original.slice());
            const linksOf = async (bytes) => { const wb = await excel.load(bytes.slice()); return excel.sheetNames(wb).flatMap((name) => excel.sheetMeta(wb, excel.sheet(name, wb)).hyperlinks.map((link) => [name, link.range ?? link.Range, link.target ?? link.Target])); };
            const before = await linksOf(original);
            check('the file has a hyperlink', before.length, 1);
            const sheetIndex = grid.GetSheetNames().indexOf(before[0][0]);
            if (sheetIndex !== grid.GetActiveSheetIndex()) grid.SwitchSheet(sheetIndex);
            check('unedited: where it was', await linksOf(await grid.ExportXlsx()), before);
            grid.InsertRow(0); grid.InsertColumn(0);
            const shifted = before[0][1].replace(/([A-Z]+)(\d+)/g, (_, col, row) => String.fromCharCode(col.charCodeAt(0) + 1) + (Number(row) + 1));
            check('it moves with an inserted row and column', await linksOf(await grid.ExportXlsx()), [[before[0][0], shifted, before[0][2]]]);
            grid.Undo(); grid.Undo();
            check('and back with undo', await linksOf(await grid.ExportXlsx()), before);
        } catch (e) { lines.push('FAIL hyperlinks threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    {
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf('tests/real-world/saopaulo-condos.xlsx');
            await grid.ImportXlsx(original.slice());
            const index = grid.GetSheets().findIndex((sheet) => [...sheet.ColumnStyles].length > 0);
            grid.SwitchSheet(index);
            const entry = `xl/worksheets/sheet${index + 1}.xml`;
            // Styled columns of the sheet, one entry per column (up to column 60).
            const styled = async (bytes) => {
                const xml = (await textParts(bytes, /^xl\/worksheets\/[^/]+\.xml$/))[entry], columns = [];
                for (const col of xml.match(/<col [^>]*>/g) ?? []) {
                    const style = /style="(\d+)"/.exec(col), min = Number(/min="(\d+)"/.exec(col)[1]), max = Number(/max="(\d+)"/.exec(col)[1]);
                    if (style && style[1] !== '0') for (let c = min; c <= Math.min(max, 60); c++) columns.push(c);
                }
                return columns;
            };
            const before = await styled(original);
            check('the sheet has columns with a default style', before.length > 0, true);
            grid.InsertColumn(0);
            check('they are still styled, one column to the right, after a column is inserted', await styled(await grid.ExportXlsx()), [...new Set(before.map((c) => Math.min(60, c + 1)).concat(before.includes(60) ? [60] : []))].sort((a, b) => a - b));
            grid.Undo();
            check('and as in the file after undo', await styled(await grid.ExportXlsx()), before);
        } catch (e) { lines.push('FAIL column styles threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    return lines;
};

// Pictures, charts and threaded comments in the grid: shown, moved with inserted rows /
// columns, and exported where they now are.
checks['drawings-in-grid'] = async () => {
    const { lines, check, note } = makeLog();
    const anchors = (wb, sheetName) => {
        const sheet = excel.sheet(sheetName, wb);
        return [...excel.images(sheet).map((image) => ['picture', image.name ?? null, image.anchor.fromRow, image.anchor.fromCol, image.anchor.toRow ?? null, image.data.length, sum(image.data)]),
            ...excel.drawingObjects(sheet).map((item) => [item.kind, item.name ?? null, item.anchor.fromRow, item.anchor.fromCol, item.anchor.toRow ?? null])];
    };
    for (const file of ['tests/test-documents-excel/synthetic-drawing-extras.xlsx', 'tests/test-documents-excel/phpsp-32chartreadwrite.xlsx', 'tests/test-documents-excel/phpsp-27template.xlsx', 'tests/real-world/chart-zoo.xlsx']) {
        const short = file.split('/').pop();
        const { grid, host, close } = mountGrid();
        try {
            const original = await bytesOf(file);
            await grid.ImportXlsx(original.slice());
            const before = await excel.load(original.slice());
            // The sheet with the most drawing objects.
            const sheetNames = grid.GetSheetNames();
            const counts = sheetNames.map((name) => anchors(before, name).length);
            const index = counts.indexOf(Math.max(...counts));
            if (index !== grid.GetActiveSheetIndex()) grid.SwitchSheet(index);
            const name = sheetNames[index], fileAnchors = anchors(before, name);
            const model = grid.GetSheets()[index];
            const drawings = Array.from(model.Drawings);
            check(`${short}: every picture and kept object is in the model`, drawings.length, fileAnchors.length);
            const pictures = fileAnchors.filter((a) => a[0] === 'picture').length;
            await everyObjectSeen(host);
            const nodes = () => [...host.querySelectorAll('.spreadsheet-drawing')];
            const boxes = () => nodes().map((node) => { const r = node.getBoundingClientRect(), c = host.querySelector('.spreadsheet-content').getBoundingClientRect(); return [Math.round(r.left - c.left), Math.round(r.top - c.top), Math.round(r.width), Math.round(r.height)]; });
            check(`${short}: the grid shows them (pictures as images, the rest as labelled frames)`, [nodes().filter((n) => n.tagName === 'IMG').length, nodes().filter((n) => n.tagName === 'DIV').length], [pictures, fileAnchors.length - pictures]);
            if (pictures > 0) {
                const image = nodes().find((n) => n.tagName === 'IMG');
                await image.decode().catch(() => {});
                note(`${short}: first picture ${image.naturalWidth}x${image.naturalHeight}, shown ${boxes()[nodes().indexOf(image)].slice(2).join('x')}`);
                if (!/synthetic/.test(short)) check(`${short}: the picture decodes`, image.naturalWidth > 0, true);
            }
            const frame = nodes().find((n) => n.tagName === 'DIV');
            // (Drawn - a chart, a shape - or a frame with its label; either way its title says what it is.)
            if (frame) check(`${short}: a kept object is drawn or its frame says what it is`, [/^(Chart|Shape|Object)/.test(frame.title), frame.querySelector('svg') !== null || /^(Chart|Shape|Object)/.test(frame.textContent)], [true, true]);
            check(`${short}: every drawing has a size on screen`, boxes().every((b) => b[2] >= 1 && b[3] >= 1), true);
            // The portable painter paints them too.
            const canvas = document.createElement('canvas');
            document.body.appendChild(canvas);
            grid.RenderPortableCanvas(canvas);
            check(`${short}: the portable painter is given the drawings`, grid.PortableDrawingItemCount(), fileAnchors.length);
            if (pictures > 0 && !/synthetic/.test(short)) {
                // The painter decodes a picture on first use; once decoded the canvas shows it, not its frame.
                await new Promise((resolve) => setTimeout(resolve, 300));
                grid.RenderPortableCanvas(canvas);
                const first = grid.GetDrawings().find((d) => d.kind === 'picture');
                const x = Math.round(50 + first.x + 2), y = Math.round(30 + first.y + 2), w = Math.max(1, Math.min(canvas.width - x, Math.round(first.width - 4))), h = Math.max(1, Math.min(canvas.height - y, Math.round(first.height - 4)));
                const data = canvas.getContext('2d').getImageData(x, y, w, h).data;
                const colours = new Set();
                for (let i = 0; i < data.length; i += 4) colours.add((data[i] << 16) | (data[i + 1] << 8) | data[i + 2]);
                if (x + 20 < canvas.width && y + 20 < canvas.height) check(`${short}: the canvas shows the picture's pixels (${colours.size} colours)`, colours.size > 8, true);
                else note(`${short}: the first picture is outside the ${canvas.width}x${canvas.height} canvas (at ${x},${y})`);
            }
            canvas.remove();

            // Two rows inserted at the top: everything on screen moves down two rows, and so does the export.
            const boxesBefore = boxes();
            grid.InsertRow(0); grid.InsertRow(0);
            // (Two rows of the height the sheet gives a row.)
            const rowsHeight = Math.round(grid.GetRowHeight(0) + grid.GetRowHeight(1));
            check(`${short}: the drawings move down with two inserted rows`, boxes().map((b, i) => [b[0] - boxesBefore[i][0], Math.abs(b[1] - boxesBefore[i][1] - rowsHeight) <= 1]), boxesBefore.map(() => [0, true]));
            const saved = await grid.ExportXlsx();
            const after = await excel.load(saved.slice());
            const absolute = new Set(Array.from(model.Drawings, (d, i) => (d.Anchor.Kind === 'absolute' ? i : -1)));
            const expected = fileAnchors.map((a) => { const copy = [...a]; copy[2] += 2; if (copy[4] != null) copy[4] += 2; return copy; });
            const sortKey = (a) => JSON.stringify([a[0], a[1], a[3]]);
            check(`${short}: exported at their new rows, pictures byte for byte`, anchors(after, name).sort((x, y) => sortKey(x).localeCompare(sortKey(y))), absolute.size > 1 ? null : expected.sort((x, y) => sortKey(x).localeCompare(sortKey(y))));
            check(`${short}: z-order kept`, zOrder(after), zOrder(before));
            grid.Undo(); grid.Undo();
            check(`${short}: undo puts them back`, boxes(), boxesBefore);
            // Threaded comments: a marker and the thread as the cell's tooltip; they follow their cell.
            const threads = excel.threadedComments(excel.sheet(name, before));
            if (threads.length > 0) {
                const titles = () => [...host.querySelectorAll('.spreadsheet-content [data-comment]')].map((n) => n.getAttribute('data-comment')).filter((t) => t.includes('Is this right?'));
                check(`${short}: the thread is shown with its author and reply`, titles().map((t) => [t.includes('Ann Author'), t.includes('Bob Reply'), t.includes('Yes & checked')]), [[true, true, true]]);
                grid.InsertRow(0);
                const moved = await excel.load((await grid.ExportXlsx()).slice());
                check(`${short}: the thread is exported at its new cell`, excel.threadedComments(excel.sheet(name, moved)).map((t) => [t.ref ?? t.cellRef ?? null, t.author ?? null, (t.replies ?? []).length]), threads.map((t) => [(t.ref ?? t.cellRef ?? '').replace(/\d+$/, (n) => String(Number(n) + 1)), t.author ?? null, (t.replies ?? []).length]));
            }
        } catch (e) { lines.push(`FAIL ${short} threw ` + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    return lines;
};

// An edit nothing depends on recalculates nothing: every cached value of the file is
// exported as it was (the incremental engine handles it, also in a workbook that has
// an array formula).
// An edit recalculates the formulas that read what it changed (Recalc), not the workbook:
// GetRecalcStats says so, the values are those of a full recalculation (the same edits in a
// grid with incremental recalculation switched off), and the times are a fraction of one.
checks['incremental-recalc'] = async () => {
    const { lines, check, note } = makeLog();
    const valuesOf = (grid) => {
        const out = [];
        for (const [addr, cell] of grid.GetModel().Cells) out.push(addr.Row + ',' + addr.Col + '=' + JSON.stringify(cell.Value));
        return out.sort();
    };
    const firstDifference = (a, b) => { for (let i = 0; i < Math.max(a.length, b.length); i++) if (a[i] !== b[i]) return `${a[i]} vs ${b[i]}`; return null; };
    // A sheet of formulas that read each other: chains, row and column sums, a whole-column sum, a lookup, text.
    const build = (grid) => {
        const rows = [];
        for (let r = 0; r < 60; r++) rows.push([r + 1, (r * 7) % 13, `=A${r + 1}*B${r + 1}`, r === 0 ? '=C1' : `=D${r}+C${r + 1}`, `=SUM(A${r + 1}:C${r + 1})`, `=IF(D${r + 1}>500,"high","low")`].join('\t'));
        grid.SelectCell(0, 0);
        grid.PasteText(rows.join('\n'), '');
        grid.EnterText(0, 8, '=SUM(A:A)'); grid.EnterText(1, 8, '=SUM(C1:C60)+I1'); grid.EnterText(2, 8, '=INDEX(D1:D60,MATCH(MAX(A1:A60),A1:A60,0))');
        grid.EnterText(3, 8, '=COUNTIF(F1:F60,"high")'); grid.EnterText(4, 8, '=I2*2');
    };
    const incremental = mountGrid(), full = mountGrid();
    try {
        full.grid.SetIncrementalRecalculation(false);
        build(incremental.grid); build(full.grid);
        check('the built sheets agree', firstDifference(valuesOf(incremental.grid), valuesOf(full.grid)), null);
        let seed = 12345;
        const random = (n) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
        const key = (grid, host, name, ctrl = false) => host.querySelector('.spreadsheet-content').dispatchEvent(new KeyboardEvent('keydown', { key: name, ctrlKey: ctrl, bubbles: true, cancelable: true }));
        const kinds = { incremental: 0, full: 0 };
        let evaluated = 0, mismatch = null, steps = 0;
        for (let step = 0; step < 120 && !mismatch; step++) {
            const kind = random(10), row = random(60), col = random(3);
            // (What an action types is drawn once: both grids get the same.)
            const typed = String(random(500)), span = random(4), overwrite = [2 + random(3), String(random(100))];
            const pasted = Array.from({ length: 1 + random(6) }, () => `${random(90)}\t${random(40)}`).join('\n');
            const formula = `=A${1 + random(60)}+SUM(B${1 + random(30)}:C${31 + random(30)})`;
            const action =
                kind < 4 ? ['a number typed', (g) => g.EnterText(row, col === 2 ? 0 : col, typed)] :
                kind === 4 ? ['a block pasted', (g) => { g.SelectCell(row, 0); g.PasteText(pasted, ''); }] :
                kind === 5 ? ['a block deleted', (g, h) => { g.SelectRange(row, 0, Math.min(59, row + span), 1); key(g, h, 'Delete'); }] :
                kind === 6 ? ['undo', (g) => g.Undo()] :
                kind === 7 ? ['redo', (g) => g.Redo()] :
                kind === 8 ? ['a formula typed', (g) => g.EnterText(row, 6, formula)] :
                ['a formula overwritten', (g) => g.EnterText(row, overwrite[0], overwrite[1])];
            action[1](incremental.grid, incremental.host); action[1](full.grid, full.host);
            steps++;
            const stats = incremental.grid.GetRecalcStats();
            kinds[stats.Kind] = (kinds[stats.Kind] ?? 0) + 1;
            if (stats.Kind === 'incremental') evaluated += stats.EvaluatedFormulas;
            const difference = firstDifference(valuesOf(incremental.grid), valuesOf(full.grid));
            if (difference) mismatch = `step ${step} (${action[0]}): ${difference}`;
        }
        check('120 random edits (typing, paste, delete, undo, redo, formulas): the values of a full recalculation after each', mismatch, null);
        check('and they were recalculated incrementally', [kinds.incremental > 100, full.grid.GetRecalcStats().Kind], [true, 'full']);
        note(`${kinds.incremental} incremental (${evaluated} formulas evaluated, of ${incremental.grid.GetRecalcStats().Formulas} in the sheet), ${kinds.full ?? 0} full`);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    incremental.close(); full.close();

    // A spilling formula deleted with the cells it spilled into, and the Delete undone: it
    // spills again (the spilled cells came back as ordinary cells, in its way: #SPILL!).
    {
        const { grid, host, close } = mountGrid();
        try {
            grid.SetCell(0, 0, '1'); grid.SetCell(1, 0, '2'); grid.SetCell(2, 0, '3');
            grid.EnterText(0, 2, '=TRANSPOSE(A1:A3)');
            const row = async () => [await valueAt(grid, 0, 2), await valueAt(grid, 0, 3), await valueAt(grid, 0, 4)];
            check('a formula that spills', await row(), ['1', '2', '3']);
            grid.SelectRange(0, 2, 0, 4);
            host.querySelector('.spreadsheet-content').dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true }));
            check('deleted with its spilled cells', await row(), ['', '', '']);
            grid.Undo();
            check('undone: it spills as before', await row(), ['1', '2', '3']);
            grid.SetCell(1, 0, '20');
            check('and follows its source', await row(), ['1', '20', '3']);
            grid.SelectRange(0, 2, 0, 4); grid.Cut(); grid.Undo();
            check('the same after a Cut is undone', await row(), ['1', '20', '3']);
        } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }

    // The formula-heavy workbook: the first edit builds the dependency graph; an edit with few
    // dependents evaluates a few formulas; a block pasted where most formulas read evaluates those.
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/real-world/milk-street-office.xlsx')).slice());
        let t0 = performance.now();
        grid.EnterText(7, 3, '54321');
        const first = performance.now() - t0, firstStats = grid.GetRecalcStats();
        t0 = performance.now();
        grid.EnterText(7, 3, '12345');
        const second = performance.now() - t0, secondStats = grid.GetRecalcStats();
        check('the first edit builds the graph and is incremental', [firstStats.Kind, firstStats.Formulas > 13000, first < 1200], ['incremental', true, true]);
        check('an edit with few dependents evaluates few formulas, in a few milliseconds', [secondStats.Kind, secondStats.EvaluatedFormulas < 50, second < 60], ['incremental', true, true]);
        grid.SwitchSheet(1);
        const block = Array.from({ length: 1000 }, (_, r) => Array.from({ length: 20 }, (_, c) => (c % 3 === 0 ? 'item ' + (r * 20 + c) : String(r * 20 + c + 0.5))).join('\t')).join('\n');
        grid.SelectCell(5, 2);
        t0 = performance.now();
        grid.PasteText(block, '');
        const paste = performance.now() - t0, pasteStats = grid.GetRecalcStats();
        t0 = performance.now();
        grid.Undo();
        const undo = performance.now() - t0, undoStats = grid.GetRecalcStats();
        check('a paste of 1 000 x 20 cells is incremental, and so is its undo', [pasteStats.Kind, pasteStats.ChangedCells > 5000, undoStats.Kind], ['incremental', true, 'incremental']);
        t0 = performance.now();
        grid.EnterText(3, 40, '77');
        const after = performance.now() - t0;
        note(`milk-street-office: first edit ${Math.round(first)} ms (${firstStats.Formulas} formulas in the graph), an edit with few dependents ${second.toFixed(1)} ms, ` +
            `paste of 1 000 x 20 ${Math.round(paste)} ms (${pasteStats.EvaluatedFormulas} formulas evaluated), its undo ${Math.round(undo)} ms, a typed number after it ${after.toFixed(1)} ms`);
        // The undo put every value back.
        const original = describeWorkbook(await excel.load((await bytesOf('tests/real-world/milk-street-office.xlsx')).slice()));
        grid.Undo(); grid.Undo(); grid.Undo();
        const saved = describeWorkbook(await excel.load((await grid.ExportXlsx()).slice()));
        const changed = [];
        original.forEach((sheet, index) => {
            const now = new Map(saved[index].cells.map((c) => [c[0] + ':' + c[1], c]));
            // (A formula that was evaluated again has the engine's value: the file's to the last digits.)
            const same = (a, b) => Object.is(a, b) || (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b)));
            for (const c of sheet.cells) { const n = now.get(c[0] + ':' + c[1]); if (c[3] != null && (!n || !same(n[2], c[2]))) changed.push(`${sheet.name}!${c[0]},${c[1]}: ${c[2]} -> ${n && n[2]}`); }
        });
        check('after undoing every edit the formulas have the values of the file', changed.slice(0, 3), []);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

checks['edit-without-dependents'] = async () => {
    const { lines, check, note } = makeLog();
    for (const file of ['tests/real-world/milk-street-office.xlsx', 'tests/real-world/saopaulo-condos.xlsx']) {
        const short = file.split('/').pop();
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf(file);
            await grid.ImportXlsx(original.slice());
            const before = describeWorkbook(await excel.load(original.slice()));
            const active = grid.GetActiveSheetIndex();
            const freeCol = Math.max(0, ...before[active].cells.map((c) => c[1])) + 3;
            const started = performance.now();
            grid.SetCell(0, freeCol, 'a note'); grid.SetCell(1, freeCol, '42'); grid.SetCell(2, freeCol, '=1+1');
            const elapsed = performance.now() - started;
            const after = describeWorkbook(await excel.load((await grid.ExportXlsx()).slice()));
            let formulas = 0;
            const changed = [];
            // Volatile formulas are recalculated on every pass, whatever was edited.
            const volatile = /\b(TODAY|NOW|RAND|RANDBETWEEN|OFFSET|INDIRECT|CELL|INFO)\(/i;
            before.forEach((sheet, index) => {
                const saved = new Map(after[index].cells.map((c) => [c[0] + ':' + c[1], c]));
                for (const c of sheet.cells) {
                    if (c[3] == null || volatile.test(c[3])) continue;
                    formulas++;
                    const now = saved.get(c[0] + ':' + c[1]);
                    if (!now || !Object.is(now[2], c[2]) || now[4] !== c[4]) changed.push(`${sheet.name}!${c[0]},${c[1]} ${c[3].slice(0, 60)}: ${c[2]} -> ${now && now[2]}`);
                }
            });
            check(`${short}: none of ${formulas} cached formula values changed`, changed.slice(0, 3), []);
            note(`${short}: three edits took ${Math.round(elapsed)} ms`);
        } catch (e) { lines.push(`FAIL ${short} threw ` + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
        close();
    }
    return lines;
};

// ---- the grid itself: scrolling, selection, keyboard, clipboard, formatting, read-only ----

export const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
// Charts and shapes of a file are laid out when they come into view: the sheet in `host` scrolled
// to each object that is not laid out yet, and back to its top.
export const everyObjectSeen = async (host) => {
    const scroller = host.querySelector('.spreadsheet-container');
    await frame();
    for (let pass = 0; pass < 60 && host.querySelectorAll('.spreadsheet-object-pending').length > 0; pass++) {
        const pending = host.querySelector('.spreadsheet-object-pending').getBoundingClientRect(), view = scroller.getBoundingClientRect();
        scroller.scrollTop += pending.top - view.top - 100; scroller.scrollLeft += pending.left - view.left - 100;
        await frame(); await frame();
    }
    scroller.scrollTop = 0; scroller.scrollLeft = 0; await frame();
};
export const parts = (host) => ({ scroller: host.querySelector('.spreadsheet-container'), content: host.querySelector('.spreadsheet-content') });
export const press = (host, key, options = {}) => parts(host).content.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...options }));
// The cell node at a point of the host (CSS pixels from its top-left corner), as [row, col], or what else is there.
const cellAtPoint = (host, x, y) => {
    const box = host.getBoundingClientRect();
    const node = document.elementFromPoint(box.left + x, box.top + y);
    const cell = node && node.closest('.spreadsheet-cell');
    if (cell) return [Number(cell.getAttribute('data-row')), Number(cell.getAttribute('data-col'))];
    for (const name of ['spreadsheet-corner', 'spreadsheet-col-header', 'spreadsheet-row-header', 'spreadsheet-edit-input']) if (node && node.closest('.' + name)) return name;
    return node ? node.className : null;
};
// Where a cell is on screen (centre), from the grid's own geometry.
export const pointOfCell = async (grid, host, row, col) => {
    const { columnLeft, rowTop } = await import('/spreadsheet/Layout.js');
    const model = grid.GetModel(), { scroller } = parts(host);
    // (The header area is the corner's: 50 x 30, plus the outline strips of a sheet with groups.)
    const corner = host.querySelector('.spreadsheet-corner');
    const headerWidth = model.ShowHeaders === false ? 0 : parseFloat(corner.style.width), headerHeight = model.ShowHeaders === false ? 0 : parseFloat(corner.style.height);
    const x = headerWidth + (columnLeft(model, col) + columnLeft(model, col + 1)) / 2 - (col < model.FrozenCols ? 0 : scroller.scrollLeft);
    const y = headerHeight + (rowTop(model, row) + rowTop(model, row + 1)) / 2 - (row < model.FrozenRows ? 0 : scroller.scrollTop);
    return [x, y];
};
export const mouse = async (grid, host, type, row, col, options = {}) => {
    const [x, y] = await pointOfCell(grid, host, row, col), box = host.getBoundingClientRect();
    const target = type === 'mousemove' || type === 'mouseup' ? document : (document.elementFromPoint(box.left + x, box.top + y) ?? parts(host).content);
    target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX: box.left + x, clientY: box.top + y, button: 0, ...options }));
};
const withFields = (record, fields) => Object.assign(Object.create(Object.getPrototypeOf(record)), record, fields);
export const valueAt = async (grid, row, col) => {
    const { getCellValue, CellAddress, displayValue } = await import('/spreadsheet/Model.js');
    return displayValue(getCellValue(new CellAddress(row, col), grid.GetModel()));
};
export const formatAt = async (grid, row, col) => {
    const { getCell, CellAddress, defaultFormat } = await import('/spreadsheet/Model.js');
    return getCell(new CellAddress(row, col), grid.GetModel())?.Format ?? defaultFormat;
};
const active = (grid) => { const cell = grid.GetActiveCell(); return [cell.Row, cell.Col]; };

// The scrolled extent follows the sheet; a sheet switch starts at the top-left with cells on
// screen; the selection is scrolled into view; typing into a cell that is off screen works;
// headers and frozen panes stay above the scrolled cells.
checks['grid-extent-scroll-and-stacking'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        const { scroller, content } = parts(host);
        check('a new grid can be scrolled over its 1000 rows', [content.getBoundingClientRect().height, content.getBoundingClientRect().width], [30 + 1000 * 25, 50 + 26 * 100]);
        grid.InstallWorkbook([createEmpty(3000, 60), createEmpty(40, 8)], ['Big', 'Small'], 0, undefined);
        check('the extent is the installed sheet\'s (3000 x 60)', [content.getBoundingClientRect().height, content.getBoundingClientRect().width], [30 + 3000 * 25, 50 + 60 * 100]);
        grid.SetCell(2900, 3, 'deep');
        scroller.scrollTop = 2890 * 25; await frame();
        check('rows past 1000 can be scrolled to and are drawn', [scroller.scrollTop, cellAtPoint(host, 400, 300)?.[0] > 2890], [2890 * 25, true]);
        grid.InsertRow(0);
        check('an inserted row grows the extent', content.getBoundingClientRect().height, 30 + 3001 * 25);
        grid.Undo();
        check('and undo shrinks it', content.getBoundingClientRect().height, 30 + 3000 * 25);
        scroller.scrollLeft = 3000; await frame();
        check('columns past Z too', cellAtPoint(host, 400, 300)?.[1] > 26, true);

        // Another sheet: at its top-left corner at once, without waiting for a scroll event.
        grid.SwitchSheet(1);
        check('a sheet switch puts the container at the top-left', [scroller.scrollTop, scroller.scrollLeft], [0, 0]);
        check('and the sheet\'s cells are on screen straight away', [cellAtPoint(host, 60, 40), cellAtPoint(host, 400, 300)], [[0, 0], [10, 3]]);
        // (The sheet was made with 40 rows and 8 columns: the grid offers more, beyond what is used and the active cell.)
        check('with the extent of that sheet', [content.getBoundingClientRect().height, content.getBoundingClientRect().width, grid.GetModel().RowCount >= 40, grid.GetModel().ColCount >= 8], [30 + grid.GetModel().RowCount * 25, 50 + grid.GetModel().ColCount * 100, true, true]);
        grid.SwitchSheet(0);
        check('and back', [scroller.scrollTop, cellAtPoint(host, 60, 40), content.getBoundingClientRect().height], [0, [0, 0], 30 + 3000 * 25]);

        // The selection is kept in view.
        grid.SelectCell(0, 0);
        for (let i = 0; i < 100; i++) press(host, 'ArrowDown');
        const visibleRows = () => [Math.ceil(scroller.scrollTop / 25), Math.floor((scroller.scrollTop + scroller.clientHeight - 30) / 25) - 1];
        check('arrow keys scroll the selection into view (down)', [active(grid), visibleRows()[0] <= 100 && 100 <= visibleRows()[1]], [[100, 0], true]);
        for (let i = 0; i < 30; i++) press(host, 'ArrowRight');
        check('(right)', [active(grid), scroller.scrollLeft > 0, scroller.scrollLeft + scroller.clientWidth - 50 >= 31 * 100], [[100, 30], true, true]);
        for (let i = 0; i < 100; i++) press(host, 'ArrowUp');
        for (let i = 0; i < 30; i++) press(host, 'ArrowLeft');
        check('(and back to A1)', [active(grid), scroller.scrollTop, scroller.scrollLeft], [[0, 0], 0, 0]);
        grid.SelectCell(2500, 50);
        check('a cell selected by the host is scrolled to', [cellAtPoint(host, ...(await pointOfCell(grid, host, 2500, 50)))], [[2500, 50]]);

        // Typing into a cell that is not on screen: it is scrolled to and the whole entry is kept.
        grid.SelectCell(5, 2);
        scroller.scrollTop = 40000; await frame();
        press(host, 'h');
        const input = host.querySelector('.spreadsheet-edit-input');
        check('typing scrolls the cell back into view and opens its edit box', [grid.IsEditing(), scroller.scrollTop < 200, document.activeElement === input, input.value], [true, true, true, 'h']);
        const inputBox = input.getBoundingClientRect(), hostBox = host.getBoundingClientRect();
        check('the edit box sits on the cell', [Math.round(inputBox.left - hostBox.left + scroller.scrollLeft), Math.round(inputBox.top - hostBox.top + scroller.scrollTop)], [50 + 200 + 1, 30 + 5 * 25 + 1]);
        input.value = 'hello'; input.dispatchEvent(new Event('input', { bubbles: true }));
        press(host, 'Enter');
        check('the whole entry is committed', [await valueAt(grid, 5, 2), grid.IsEditing(), active(grid)], ['hello', false, [6, 2]]);

        // Headers stay on top of the scrolled cells; the corner on top of both.
        scroller.scrollTop = 5000; scroller.scrollLeft = 730; await frame();
        check('column headers, row headers and the corner cover the scrolled cells', [cellAtPoint(host, 400, 12), cellAtPoint(host, 20, 300), cellAtPoint(host, 20, 12)], ['spreadsheet-col-header', 'spreadsheet-row-header', 'spreadsheet-corner']);
        const header = document.elementFromPoint(hostBox.left + 400, hostBox.top + 12).closest('.spreadsheet-col-header');
        check('the header over a column is that column\'s', Number(header.getAttribute('data-col')), cellAtPoint(host, 400, 300)[1]);
        check('headers do not move with a vertical scroll', Math.round(header.getBoundingClientRect().top - scroller.getBoundingClientRect().top), 1);

        // Frozen panes: two rows and one column stay; the scrolled cells pass under them.
        grid.InstallWorkbook([withFields(createEmpty(3000, 60), { FrozenRows: 2, FrozenCols: 1 })], ['Frozen'], 0, undefined);
        scroller.scrollTop = 5000; scroller.scrollLeft = 730; await frame();
        const body = cellAtPoint(host, 500, 300);
        check('frozen rows stay at the top, above the scrolled cells', [cellAtPoint(host, 500, 30 + 12)[0], cellAtPoint(host, 500, 30 + 37)[0], body[0] > 190], [0, 1, true]);
        check('they scroll sideways with their columns', cellAtPoint(host, 500, 30 + 12)[1], body[1]);
        check('frozen columns stay at the left', [cellAtPoint(host, 50 + 50, 300)[1], cellAtPoint(host, 50 + 50, 300)[0]], [0, body[0]]);
        check('the frozen corner is above both', cellAtPoint(host, 50 + 50, 30 + 12), [0, 0]);
        check('headers above everything', [cellAtPoint(host, 500, 12), cellAtPoint(host, 20, 300), cellAtPoint(host, 20, 12), cellAtPoint(host, 50 + 50, 12), cellAtPoint(host, 20, 30 + 12)], ['spreadsheet-col-header', 'spreadsheet-row-header', 'spreadsheet-corner', 'spreadsheet-col-header', 'spreadsheet-row-header']);
        const frozenHeader = document.elementFromPoint(hostBox.left + 100, hostBox.top + 12).closest('.spreadsheet-col-header');
        check('the header of the frozen column is column A', frozenHeader.getAttribute('data-col'), '0');
        grid.SelectCell(300, 40);
        const at = await pointOfCell(grid, host, 300, 40);
        check('a selected cell is revealed clear of the frozen panes', [cellAtPoint(host, ...at), at[0] > 150, at[1] > 80], [[300, 40], true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
// A sheet may be taller than the browser's layout range. The scrollbar must
// reach its end without changing cell sizes or losing the logical coordinates.
checks['grid-full-height-zoom'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        const last = 1048575;
        grid.InstallWorkbook([createEmpty(last + 1, 26), createEmpty(100, 8), withFields(createEmpty(100, 16384), { DefaultColumnWidth: 250 })], ['Full', 'Small', 'Wide'], 0, undefined);
        grid.SetRowHeight(last, 55);
        grid.SetRowHeight(last - 2, 37);
        grid.HideRows(last - 1);
        grid.FreezeAt(1, 1);
        const { scroller } = parts(host);
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const centre = (element) => {
            const r = element.getBoundingClientRect();
            return [r.left + r.width / 2, r.top + r.height / 2];
        };
        const click = (element) => {
            const [clientX, clientY] = centre(element);
            const target = document.elementFromPoint(clientX, clientY);
            target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, clientX, clientY }));
            document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, button: 0, clientX, clientY }));
        };
        const aligned = (a, b) => ['left', 'top', 'width', 'height'].every((key) => Math.abs(a[key] - b[key]) <= 2);
        for (const zoom of [4, 1, 0.1]) {
            grid.SetZoom(zoom);
            scroller.scrollTop = scroller.scrollHeight;
            await frame(); await frame();
            const cell = node(last, 1);
            const bounds = cell?.getBoundingClientRect(), view = scroller.getBoundingClientRect();
            check(`zoom ${zoom}: scrollbar reaches the last visible row`, !!bounds && bounds.bottom <= view.bottom + 2 && bounds.top >= view.top && Math.abs(bounds.height - 55 * zoom) < 2, true);
            check(`zoom ${zoom}: extent stays within the browser ceiling`, scroller.scrollHeight <= 8000001, true);
            check(`zoom ${zoom}: hidden row is not drawn`, !node(last - 1, 1), true);
            click(cell);
            check(`zoom ${zoom}: clicking the last row hit-tests it`, active(grid), [last, 1]);
            const activeBox = [...host.querySelectorAll('.spreadsheet-active-cell')].find((el) => getComputedStyle(el).display !== 'none');
            check(`zoom ${zoom}: active outline follows the cell`, !!activeBox && aligned(activeBox.getBoundingClientRect(), cell.getBoundingClientRect()), true);
            press(host, 'F2');
            const input = host.querySelector('.spreadsheet-edit-input');
            check(`zoom ${zoom}: edit box follows the cell`, grid.IsEditing() && aligned(input.getBoundingClientRect(), cell.getBoundingClientRect()), true);
            input.value = `last-${zoom}`; input.dispatchEvent(new Event('input', { bubbles: true }));
            press(host, 'Enter');
            check(`zoom ${zoom}: last row is editable`, await valueAt(grid, last, 1), `last-${zoom}`);
            grid.SelectCell(last, 1); await frame();
            check(`zoom ${zoom}: SelectCell reveals the last row`, !!node(last, 1) && active(grid)[0] === last, true);
            press(host, 'ArrowUp');
            check(`zoom ${zoom}: arrow navigation skips the hidden row`, active(grid), [last - 2, 1]);
            press(host, 'PageUp'); const before = active(grid)[0];
            press(host, 'PageDown');
            check(`zoom ${zoom}: page navigation moves through the bottom of the sheet`, before < last - 2 && active(grid)[0] > before, true);
            grid.GoTo('B1048576'); await frame();
            check(`zoom ${zoom}: GoTo reveals the last row`, [active(grid), !!node(last, 1)], [[last, 1], true]);
            grid.SelectRange(last - 2, 1, last, 2); await frame();
            const selection = [...host.querySelectorAll('.spreadsheet-selection')].find((el) => getComputedStyle(el).display !== 'none');
            const selectedCell = node(last, 1)?.getBoundingClientRect(), selectedBox = selection?.getBoundingClientRect();
            check(`zoom ${zoom}: range overlay covers the last cell`, !!selectedBox && !!selectedCell && selectedBox.top <= selectedCell.top + 2 && selectedBox.bottom >= selectedCell.bottom - 2, true);
            check(`zoom ${zoom}: frozen corner remains visible`, !!node(0, 0) && node(0, 0).getBoundingClientRect().top < view.top + 31 * zoom + 2, true);
        }
        grid.SetZoom(4); grid.SelectCell(500000, 1); await frame();
        grid.SetZoom(1); await frame();
        check('zooming preserves the deep viewport', !!node(500000, 1), true);
        grid.SetZoom(4); await frame();
        const view = scroller.getBoundingClientRect();
        const target = document.elementFromPoint(view.left + 700, view.top + 320)?.closest('.spreadsheet-cell');
        const before = target?.getBoundingClientRect().top;
        scroller.dispatchEvent(new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: 40 }));
        await frame();
        check('wheel travel remains in screen pixels on the bounded scrollbar', !!target && Math.abs(target.getBoundingClientRect().top - before + 40) <= 4, true);
        const canvas = document.createElement('canvas'); canvas.width = 30; canvas.height = 15;
        const picture = grid.InsertPicture(canvas.toDataURL('image/png'), 'B500001');
        grid.SelectCell(500000, 1); await frame();
        const image = host.querySelector(`.spreadsheet-picture[data-drawing-id="${picture}"]`);
        const pictureCell = node(500000, 1)?.getBoundingClientRect(), pictureBox = image?.getBoundingClientRect();
        check('a drawing at a rebased row follows its cell', !!pictureBox && !!pictureCell && Math.abs(pictureBox.left - pictureCell.left) <= 2 && Math.abs(pictureBox.top - pictureCell.top) <= 2, true);
        grid.DeletePicture(picture);
        grid.SwitchSheet(1); await frame();
        check('switching sheets resets the bounded scroll origin', [scroller.scrollTop, active(grid)[0], !!node(0, 0)], [0, 0, true]);
        grid.SwitchSheet(0); await frame();
        check('returning to the full sheet starts at the top', [scroller.scrollTop, !!node(0, 0)], [0, true]);
        grid.SwitchSheet(2); grid.SetZoom(4); grid.SelectCell(1, 16383); await frame();
        const lastColumn = node(1, 16383);
        check('a bounded full-width sheet reveals its last column', [!!lastColumn && lastColumn.getBoundingClientRect().left < scroller.getBoundingClientRect().right, active(grid)], [true, [1, 16383]]);
        scroller.scrollLeft = scroller.scrollWidth; await frame();
        check('the horizontal scrollbar reaches the last column', !!node(1, 16383) && scroller.scrollWidth <= 8000001, true);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};


// Selection with the keyboard and the mouse, the active cell, and the editing API.
checks['grid-selection-and-keyboard'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([createEmpty(400, 30)], ['Keys'], 0, undefined);
        // A block of data B3:D8, a gap, and more in column B further down.
        for (let row = 2; row <= 7; row++) for (let col = 1; col <= 3; col++) grid.SetCell(row, col, String(row * 10 + col));
        grid.SetCell(20, 1, 'x'); grid.SetCell(21, 1, 'y');
        const selections = [];
        const off = grid.OnSelectionChange((change) => selections.push(change));
        grid.SelectCell(2, 1);
        check('the selection event names the active cell', [selections.at(-1).Selection, selections.at(-1).ActiveCell, selections.at(-1).ActiveRow, selections.at(-1).ActiveCol], ['B3', 'B3', 2, 1]);
        press(host, 'ArrowDown', { shiftKey: true }); press(host, 'ArrowRight', { shiftKey: true }); press(host, 'ArrowDown', { shiftKey: true });
        check('Shift+arrow extends from the active cell', [grid.GetSelectionRef(), active(grid)], ['B3:C5', [2, 1]]);
        press(host, 'ArrowUp', { shiftKey: true });
        check('and shrinks again', grid.GetSelectionRef(), 'B3:C4');
        grid.SelectCell(5, 3);
        press(host, 'ArrowUp', { shiftKey: true }); press(host, 'ArrowLeft', { shiftKey: true });
        check('extending up and left keeps the active cell at the bottom-right', [grid.GetSelectionRef(), active(grid)], ['C5:D6', [5, 3]]);
        press(host, 'ArrowLeft');
        check('a plain arrow collapses to the active cell and moves', [grid.GetSelectionRef(), active(grid)], ['C6', [5, 2]]);

        // Mouse: click, Shift+click, drag (the cell the drag started at stays the active cell).
        await mouse(grid, host, 'mousedown', 3, 1); await mouse(grid, host, 'mouseup', 3, 1);
        await mouse(grid, host, 'mousedown', 6, 3, { shiftKey: true }); await mouse(grid, host, 'mouseup', 6, 3);
        check('Shift+click extends from the active cell', [grid.GetSelectionRef(), active(grid)], ['B4:D7', [3, 1]]);
        await mouse(grid, host, 'mousedown', 7, 3); await mouse(grid, host, 'mousemove', 5, 2); await mouse(grid, host, 'mousemove', 2, 1); await mouse(grid, host, 'mouseup', 2, 1);
        check('a drag up and to the left keeps its starting cell active', [grid.GetSelectionRef(), active(grid)], ['B3:D8', [7, 3]]);
        await mouse(grid, host, 'mousemove', 12, 5);
        check('the drag ended with the button', grid.GetSelectionRef(), 'B3:D8');

        // Ctrl+arrow: the edges of the data.
        grid.SelectCell(2, 1);
        press(host, 'ArrowDown', { ctrlKey: true });
        check('Ctrl+Down: the last cell of the block', active(grid), [7, 1]);
        press(host, 'ArrowDown', { ctrlKey: true });
        check('again: the next cell with data', active(grid), [20, 1]);
        press(host, 'ArrowDown', { ctrlKey: true }); press(host, 'ArrowDown', { ctrlKey: true });
        check('then the end of that run, then the last row a sheet can have', active(grid), [1048575, 1]);
        press(host, 'ArrowUp', { ctrlKey: true });
        check('Ctrl+Up comes back to the data', active(grid), [21, 1]);
        grid.SelectCell(2, 1);
        press(host, 'ArrowRight', { ctrlKey: true, shiftKey: true }); press(host, 'ArrowDown', { ctrlKey: true, shiftKey: true });
        check('Ctrl+Shift+arrows select the block', [grid.GetSelectionRef(), active(grid)], ['B3:D8', [2, 1]]);
        press(host, 'ArrowRight', { ctrlKey: true });
        check('Ctrl+Right from the active cell', active(grid), [2, 3]);
        press(host, 'ArrowRight', { ctrlKey: true });
        check('with nothing further: the last column a sheet can have', active(grid), [2, 16383]);

        // Home / End / Ctrl+Home / Ctrl+End / PageDown / PageUp.
        grid.SelectCell(4, 2);
        press(host, 'Home'); check('Home: the first column of the row', active(grid), [4, 0]);
        press(host, 'End'); check('End: the last cell with data in the row', active(grid), [4, 3]);
        press(host, 'End', { ctrlKey: true }); check('Ctrl+End: the last cell of the used range', active(grid), [21, 3]);
        press(host, 'Home', { ctrlKey: true }); check('Ctrl+Home: the first cell', active(grid), [0, 0]);
        const pageRows = Math.floor((parts(host).scroller.clientHeight - 30) / 25);
        press(host, 'PageDown');
        check('PageDown moves a screenful', [Math.abs(active(grid)[0] - pageRows) <= 1, active(grid)[1]], [true, 0]);
        const afterOne = active(grid)[0];
        press(host, 'PageDown'); press(host, 'PageUp');
        check('PageUp comes back', active(grid), [afterOne, 0]);
        press(host, 'PageUp'); press(host, 'PageUp');
        check('and stops at the top', active(grid), [0, 0]);
        press(host, 'PageDown', { shiftKey: true });
        check('Shift+PageDown selects', [grid.GetSelectionRef(), active(grid)], [`A1:A${afterOne + 1}`, [0, 0]]);
        press(host, 'a', { ctrlKey: true });
        // (Every row and column the grid offers: the 400 x 30 it was made with, and the room beyond what is used.)
        check('Ctrl+A selects the sheet, the active cell stays', [Number(/^A1:[A-Z]+(\d+)$/.exec(grid.GetSelectionRef())?.[1]), grid.GetModel().RowCount >= 400, grid.GetModel().ColCount >= 30, active(grid)], [grid.GetModel().RowCount, true, true, [0, 0]]);

        // Enter / Tab walk through a selection; with one cell selected they move it (Enter edits).
        grid.SelectRange(2, 1, 3, 2);
        const walk = (key, options, n) => { const seen = []; for (let i = 0; i < n; i++) { press(host, key, options); seen.push(active(grid).join(',')); } return seen; };
        check('Tab walks the selection row by row', [walk('Tab', {}, 4), grid.GetSelectionRef()], [['2,2', '3,1', '3,2', '2,1'], 'B3:C4']);
        check('Shift+Tab walks back', walk('Tab', { shiftKey: true }, 2), ['3,2', '3,1']);
        grid.SelectRange(2, 1, 3, 2);
        check('Enter walks it column by column', walk('Enter', {}, 4), ['3,1', '2,2', '3,2', '2,1']);
        check('Shift+Enter walks back', walk('Enter', { shiftKey: true }, 2), ['3,2', '2,2']);
        // Typing into the active cell of a selection keeps the selection; Enter goes on to the next cell.
        grid.SelectRange(10, 5, 11, 5);
        press(host, '7');
        const input = host.querySelector('.spreadsheet-edit-input');
        input.value = '77'; input.dispatchEvent(new Event('input', { bubbles: true }));
        press(host, 'Enter');
        press(host, '8'); press(host, 'Enter');
        check('entries typed down a selected column', [await valueAt(grid, 10, 5), await valueAt(grid, 11, 5), grid.GetSelectionRef(), active(grid)], ['77', '8', 'F11:F12', [10, 5]]);
        grid.SelectCell(12, 5);
        press(host, 'Tab'); check('Tab with one cell selected moves right', active(grid), [12, 6]);
        press(host, 'Tab', { shiftKey: true }); check('Shift+Tab moves left', active(grid), [12, 5]);
        press(host, 'Enter');
        check('Enter on a single cell starts editing it', grid.IsEditing(), true);

        // The editing API.
        const edits = [];
        const offEdit = grid.OnEditStateChange((state) => edits.push([state.Editing, state.Row, state.Col, state.Text]));
        grid.CancelEdit();
        check('CancelEdit drops the entry', [grid.IsEditing(), edits], [false, [[false, -1, -1, '']]]);
        edits.length = 0;
        grid.SelectCell(14, 0);
        press(host, 'a');
        input.value = 'abc'; input.dispatchEvent(new Event('input', { bubbles: true }));
        check('OnEditStateChange: start and text', edits, [[true, 14, 0, 'a'], [true, 14, 0, 'abc']]);
        const undoBefore = grid.CanUndo();
        check('CommitEdit commits where Enter would, without moving', [grid.CommitEdit(), grid.IsEditing(), await valueAt(grid, 14, 0), active(grid), edits.at(-1)], [true, false, 'abc', [14, 0], [false, -1, -1, '']]);
        check('CanUndo / CanRedo', [undoBefore, grid.CanUndo(), grid.CanRedo()], [true, true, false]);
        grid.Undo();
        check('after an undo', [await valueAt(grid, 14, 0), grid.CanRedo()], ['', true]);
        offEdit(); edits.length = 0;
        grid.StartEdit('zz'); grid.SetEditText('typed in a formula bar');
        check('StartEdit / SetEditText / CommitEdit from a formula bar', [input.value, grid.CommitEdit(), await valueAt(grid, 14, 0), edits.length], ['typed in a formula bar', true, 'typed in a formula bar', 0]);
        // A fresh grid has nothing to undo.
        const other = mountGrid();
        check('a new grid has nothing to undo or redo', [other.grid.CanUndo(), other.grid.CanRedo()], [false, false]);
        other.close();
        // OnValidationNotice returns its unsubscribe.
        const notices = [];
        const offNotice = grid.OnValidationNotice((style) => notices.push(style));
        check('OnValidationNotice returns an unsubscribe function', typeof offNotice, 'function');
        offNotice(); off();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Copy, cut and paste through the system clipboard: tab-separated text and an HTML table.
checks['grid-clipboard'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { content } = parts(host);
        const clipboard = (type, data = {}) => {
            const transfer = new DataTransfer();
            for (const [flavour, text] of Object.entries(data)) transfer.setData(flavour, text);
            const event = clipboardEvent(type, transfer);
            content.dispatchEvent(event);
            return { text: transfer.getData('text/plain'), html: transfer.getData('text/html'), handled: event.defaultPrevented };
        };
        grid.SetCell(0, 0, 'Name'); grid.SetCell(0, 1, 'Qty'); grid.SetCell(0, 2, 'Note');
        grid.SetCell(1, 0, 'Apples'); grid.SetCell(1, 1, '12'); grid.SetCell(1, 2, 'tab\there');
        grid.SetCell(2, 0, 'Pears'); grid.SetCell(2, 1, '=B2*2'); grid.SetCell(2, 2, 'say "hi"');
        grid.SelectCell(1, 1); grid.ApplyFormat('bold');
        grid.SelectRange(0, 0, 2, 2);
        const copied = clipboard('copy');
        check('copy puts tab-separated text on the clipboard', [copied.handled, copied.text], [true, 'Name\tQty\tNote\nApples\t12\t"tab\there"\nPears\t24\t"say ""hi"""']);
        const table = new DOMParser().parseFromString(copied.html, 'text/html').querySelector('table');
        check('and an HTML table', [table.rows.length, table.rows[0].cells.length, table.rows[2].cells[1].textContent, /bold/.test(table.rows[1].cells[1].getAttribute('style') ?? '')], [3, 3, '24', true]);
        check('GetSelectionText / GetSelectionHtml give the same', [grid.GetSelectionText() === copied.text, grid.GetSelectionHtml() === copied.html], [true, true]);

        // The grid's own cells pasted back: formulas move, formats come along.
        grid.SelectCell(5, 0);
        const own = clipboard('paste', { 'text/plain': copied.text, 'text/html': copied.html });
        const { editText } = await import('/spreadsheet/Input.js');
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        check('pasting the grid\'s own copy keeps formulas and formats', [own.handled, await valueAt(grid, 7, 1), editText(getCell(new CellAddress(7, 1), grid.GetModel())), (await formatAt(grid, 6, 1)).Bold, grid.GetSelectionRef()], [true, '24', '=B7*2', true, 'A6:C8']);
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        grid.Undo();
        check('the paste is one undo step', [await valueAt(grid, 5, 0), await valueAt(grid, 7, 2), changes.length], ['', '', 1]);

        // Cells from another application: tab-separated text (as Excel writes it) ...
        grid.SelectCell(10, 1);
        clipboard('paste', { 'text/plain': 'Region\tSales\r\nNorth\t1,250.50\r\n"multi\nline"\t=1+2\r\n' });
        check('pasting tab-separated text: values are read like typed entries', [await valueAt(grid, 10, 1), await valueAt(grid, 11, 2), await valueAt(grid, 12, 1), await valueAt(grid, 12, 2), grid.GetSelectionRef()], ['Region', '1250.5', 'multi\nline', '3', 'B11:C13']);
        grid.Undo();
        check('one undo step', [await valueAt(grid, 10, 1), await valueAt(grid, 12, 2)], ['', '']);
        // ... or an HTML table (as Excel and Sheets also write).
        grid.SelectCell(10, 1);
        clipboard('paste', { 'text/plain': 'ignored when there is a table', 'text/html': '<html><body><!--StartFragment--><table><tr><td>Q1</td><td style="font-weight:bold">42</td></tr><tr><td>Q2</td><td>58%</td></tr></table><!--EndFragment--></body></html>' });
        check('pasting an HTML table', [await valueAt(grid, 10, 1), await valueAt(grid, 10, 2), await valueAt(grid, 11, 1), (await import('/spreadsheet/Model.js')).getCellValue(new CellAddress(11, 2), grid.GetModel()).fields[0]], ['Q1', '42', 'Q2', 0.58]);
        // One copied value fills the whole selection.
        grid.SelectRange(15, 0, 16, 2);
        clipboard('paste', { 'text/plain': '9' });
        check('a single value fills the selection', [await valueAt(grid, 15, 0), await valueAt(grid, 16, 2), await valueAt(grid, 17, 0)], ['9', '9', '']);
        check('a newline-only external cell clears the selected values',
            [grid.PasteText('\n', ''), await valueAt(grid, 15, 0), await valueAt(grid, 16, 2)],
            [true, '', '']);
        check('PasteText does the same for a host', [(grid.SelectCell(20, 0), grid.PasteText('a\tb', '')), await valueAt(grid, 20, 1)], [true, 'b']);

        // Cut: on the clipboard, and gone from the sheet in one step.
        grid.SelectRange(0, 0, 0, 2);
        const cut = clipboard('cut');
        check('cut copies and clears', [cut.text, await valueAt(grid, 0, 0), await valueAt(grid, 0, 2)], ['Name\tQty\tNote', '', '']);
        grid.Undo();
        check('and is undone in one step', await valueAt(grid, 0, 1), 'Qty');
        // While a cell is being edited the clipboard belongs to the edit box.
        grid.SelectCell(1, 0); press(host, 'F2');
        check('copy while editing is left to the edit box', clipboard('copy').handled, false);
        grid.CancelEdit();
        // The programmatic calls resolve even when the browser refuses the system clipboard.
        grid.SelectRange(1, 0, 1, 1);
        const copyResult = await grid.Copy();
        check('Copy() resolves to whether the system clipboard took it', typeof copyResult, 'boolean');
        grid.SelectCell(25, 0);
        const pasted = await grid.Paste();
        check('Paste() pastes (the grid\'s own cells when the system clipboard cannot be read)', [pasted, await valueAt(grid, 25, 0), await valueAt(grid, 25, 1)], [true, 'Apples', '12']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// The internal clipboard belongs to the workbook, not the sheet currently displayed.
checks['grid-clipboard-cross-sheet-copy'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    const ownClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([createEmpty(40, 12), createEmpty(40, 12)], ['Source', 'Target'], 0, undefined);
        grid.SetCell(1, 0, '13');
        grid.SetCell(1, 1, '=A2*2');
        grid.SetCell(1, 2, '=B2+1');
        grid.SelectCell(1, 1); grid.ApplyFormat('bold');
        grid.SelectCell(1, 2); grid.ApplyBgColor('#DDEBF7');
        grid.SelectRange(1, 1, 1, 2);
        const transfer = new DataTransfer();
        const copyEvent = clipboardEvent('copy', transfer);
        parts(host).content.dispatchEvent(copyEvent);
        const copiedText = transfer.getData('text/plain'), copiedHtml = transfer.getData('text/html');
        check('copy event exposes the matching system text', [copyEvent.defaultPrevented, copiedText], [true, '26\t27']);
        const held = { text: copiedText, html: copiedHtml, reads: 0 };
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
            read: async () => { held.reads++; return [{ types: ['text/plain', 'text/html'], getType: async (type) => new Blob([type === 'text/plain' ? held.text : held.html], { type }) }]; },
            readText: async () => held.text,
        } });
        // A matching system read uses the rich internal cells, not the HTML table of displayed values.
        grid.SwitchSheet(1);
        grid.SetCell(3, 2, '9');
        grid.SetCell(3, 3, '901');
        grid.SelectCell(3, 3);
        check('matching system clipboard pastes rich cells onto another sheet', [await grid.Paste(), held.reads, grid.GetRange('Target!D4:E4')[0].map((cell) => [cell.formula, cell.value]), grid.GetSelectionRef()],
            [true, 1, [['=C4*2', 18], ['=D4+1', 19]], 'D4:E4']);
        check('cross-sheet copy preserves both cell formats', [grid.GetCell(3, 3).format.bold, grid.GetCell(3, 4).format.fill], [true, '#DDEBF7']);
        check('copy leaves its source formulas and values untouched', grid.GetRange('Source!B2:C2')[0].map((cell) => [cell.formula, cell.value]), [['=A2*2', 26], ['=B2+1', 27]]);
        grid.Undo();
        check('copy paste undo restores only the target', [grid.GetCell(3, 3).value, grid.GetCell(3, 4).value, grid.GetRange('Source!B2')[0][0].value], [901, null, 26]);
        grid.Redo();
        check('copy paste redo restores shifted formulas', grid.GetRange('Target!D4:E4')[0].map((cell) => [cell.formula, cell.value]), [['=C4*2', 18], ['=D4+1', 19]]);

        // An unrelated system clipboard always wins over the grid's prior copy.
        held.text = 'other application'; held.html = '';
        grid.SelectCell(8, 3);
        check('unrelated system text takes precedence over internal copy', [await grid.Paste(), grid.GetCell(8, 3).value, grid.GetCell(8, 3).formula], [true, 'other application', null]);
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
            read: async () => { throw new Error('clipboard read denied'); },
            readText: async () => { throw new Error('clipboard read denied'); },
        } });
        grid.SelectCell(10, 3);
        check('denied system read falls back to cross-sheet rich copy', [await grid.Paste(), grid.GetCell(10, 3).formula, grid.GetCell(10, 4).formula, grid.GetCell(10, 3).format.bold],
            [true, '=C11*2', '=D11+1', true]);
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
        grid.SelectCell(12, 2); grid.SetCell(12, 2, '4');
        grid.SelectCell(12, 3);
        check('unavailable system clipboard falls back to cross-sheet rich copy', [await grid.Paste(), grid.GetCell(12, 3).formula, grid.GetCell(12, 4).value],
            [true, '=C13*2', 9]);

        // Import replaces the workbook, even if its sheet has the same name as the copied source.
        const fresh = mountGrid();
        let bytes;
        try {
            fresh.grid.InstallWorkbook([fresh.grid.GetModel()], ['Source'], 0, undefined);
            fresh.grid.SetCell(0, 0, '314');
            bytes = await fresh.grid.ExportXlsx();
        }
        finally { fresh.close(); }
        await grid.ImportXlsx(bytes.slice());
        grid.SelectCell(4, 3);
        check('import invalidates prior workbook internal clipboard', [grid.GetSheetNames(), await grid.Paste(), grid.GetCell(4, 3).value, grid.GetCell(0, 0).value],
            [['Source'], false, null, 314]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        if (ownClipboard) Object.defineProperty(navigator, 'clipboard', ownClipboard); else delete navigator.clipboard;
        close();
    }
    return lines;
};

checks['grid-clipboard-cross-sheet-cut'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const ownClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([createEmpty(40, 12), createEmpty(40, 12), createEmpty(40, 12)], ['Source', 'Target', 'Witness'], 0, undefined);
        grid.SetCell(0, 0, '11');
        grid.SetCell(0, 1, '=A1*3');
        grid.SetCell(0, 2, '=B1+Source!$D$1');
        grid.SetCell(0, 3, '7');
        grid.SetCell(0, 4, '=B1+1');
        grid.SetDefinedName('MovedNumber', '=Source!$B$1');
        grid.SwitchSheet(1);
        grid.SetCell(0, 0, '70');
        grid.SetCell(0, 1, '=Source!A1+5');
        grid.SetCell(4, 5, '901'); grid.SetCell(4, 6, '902'); grid.SetCell(4, 7, '903');
        grid.SwitchSheet(2);
        grid.SetCell(0, 0, '=Source!B1+2');
        grid.SetCell(0, 1, '=MovedNumber+3');
        grid.SwitchSheet(0);
        grid.SelectRange(0, 0, 0, 2);
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
        // Cut clears immediately, independently of the later paste.
        await grid.Cut();
        check('cut immediately clears source cells, not their neighbour', [grid.GetRange('Source!A1:C1')[0].map((cell) => cell.value), grid.GetCell(0, 3).value], [[null, null, null], 7]);
        grid.SwitchSheet(1);
        grid.SelectCell(4, 5);
        check('cross-sheet cut pastes stored source, not active target cells', [await grid.Paste(), grid.GetRange('Target!F5:H5')[0].map((cell) => [cell.formula, cell.value])],
            [true, [[null, 11], ['=F5*3', 33], ['=G5+Source!$D$1', 40]]]);
        const dependents = () => [grid.GetRange('Source!E1')[0][0], grid.GetRange('Target!B1')[0][0], grid.GetRange('Witness!A1:B1')[0][0], grid.GetRange('Witness!B1')[0][0]];
        check('carried formulas keep stationary source dependencies and self refs travel', [grid.GetCell(4, 7).formula, grid.GetCell(4, 7).value], ['=G5+Source!$D$1', 40]);
        check('foreign dependents and defined name follow the moved block', [dependents().map((cell) => [cell.formula, cell.value]), grid.GetDefinedNames().find((name) => name.name === 'MovedNumber').refersTo],
            [[['=Target!G5+1', 34], ['=F5+5', 16], ['=Target!G5+2', 35], ['=MovedNumber+3', 36]], 'Target!$G$5']);

        grid.Undo();
        check('paste undo restores target and workbook references, but not the earlier cut', [grid.GetRange('Target!F5:H5')[0].map((cell) => cell.value),
            grid.GetRange('Source!A1:C1')[0].map((cell) => cell.value), dependents().map((cell) => cell.formula),
            grid.GetDefinedNames().find((name) => name.name === 'MovedNumber').refersTo],
            [[901, 902, 903], [null, null, null], ['=B1+1', '=Source!A1+5', '=Source!B1+2', '=MovedNumber+3'], 'Source!$B$1']);
        grid.SwitchSheet(0);
        grid.Undo();
        check('source cut undo after switching back restores source, leaves target intact', [grid.GetRange('Source!A1:C1')[0].map((cell) => [cell.formula, cell.value]),
            grid.GetRange('Target!F5:H5')[0].map((cell) => cell.value), dependents().map((cell) => cell.value)],
            [[[null, 11], ['=A1*3', 33], ['=B1+Source!$D$1', 40]], [901, 902, 903], [34, 16, 35, 36]]);
        grid.Redo();
        grid.SwitchSheet(1);
        grid.Redo();
        check('cut and paste redo restores source clear, target and foreign references',
            [grid.GetRange('Source!A1:C1')[0].map((cell) => cell.value), grid.GetRange('Target!F5:H5')[0].map((cell) => cell.value),
                dependents().map((cell) => cell.value), grid.GetDefinedNames().find((name) => name.name === 'MovedNumber').refersTo],
            [[null, null, null], [11, 33, 40], [34, 16, 35, 36], 'Target!$G$5']);

        const reopened = mountGrid();
        try {
            await reopened.grid.ImportXlsx((await grid.ExportXlsx()).slice());
            check('save and reopen retains exact moved formulas and calculated results',
                [reopened.grid.GetRange('Target!F5:H5')[0].map((cell) => [cell.formula, cell.value]),
                    reopened.grid.GetRange('Source!A1:E1')[0].map((cell) => [cell.formula, cell.value]),
                    reopened.grid.GetRange('Witness!A1:B1')[0].map((cell) => [cell.formula, cell.value]),
                    reopened.grid.GetDefinedNames().find((name) => name.name === 'MovedNumber').refersTo],
                [[[null, 11], ['=F5*3', 33], ['=G5+Source!$D$1', 40]],
                    [[null, null], [null, null], [null, null], [null, 7], ['=Target!G5+1', 34]],
                    [['=Target!G5+2', 35], ['=MovedNumber+3', 36]], 'Target!$G$5']);
        } finally { reopened.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        if (ownClipboard) Object.defineProperty(navigator, 'clipboard', ownClipboard); else delete navigator.clipboard;
        close();
    }
    return lines;
};

checks['grid-clipboard-deleted-source'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const ownClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    try {
        grid.AddSheet('Target', undefined);
        grid.SwitchSheet(0); grid.SetCell(0, 0, '51'); grid.SetCell(0, 1, '=A1+1');
        grid.SelectRange(0, 0, 0, 1);
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
        await grid.Cut();
        grid.DeleteSheet(0);
        grid.SelectCell(3, 3);
        check('deleted cut source cannot supply a stale internal paste', [grid.GetSheetNames(), await grid.Paste(), grid.GetCell(3, 3).value],
            [['Target'], false, null]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        if (ownClipboard) Object.defineProperty(navigator, 'clipboard', ownClipboard); else delete navigator.clipboard;
        close();
    }
    return lines;
};

checks['grid-clipboard-session-lifecycle'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
        grid.InstallWorkbook([createEmpty(20, 10), createEmpty(20, 10)], ['Source', 'Target'], 0, undefined);
        grid.SetCell(0, 0, '6'); grid.SetCell(0, 1, '=Source!$A$1*2');
        grid.MarkSaved();
        grid.SetCell(0, 3, '9'); grid.Undo();
        const saved = grid.GetRevision();
        grid.SelectCell(0, 1); await grid.Copy();
        check('copy retains saved revision and redo', [grid.GetRevision(), grid.IsModified(), grid.CanRedo()], [saved, false, true]);
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '' } });
        grid.SwitchSheet(1); grid.SelectCell(0, 1);
        check('an empty system read retains the internal rich fallback',
            [await grid.Paste(), grid.GetCell(0, 1).formula, grid.GetCell(0, 1).value],
            [true, '=Source!$A$1*2', 12]);
        grid.Undo(); grid.SwitchSheet(0);
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
        grid.RenameSheet(0, 'Renamed Source');
        grid.SwitchSheet(1); grid.SelectCell(1, 1);
        check('copied formula follows source rename', [await grid.Paste(), grid.GetCell(1, 1).formula, grid.GetCell(1, 1).value],
            [true, "='Renamed Source'!$A$1*2", 12]);
        grid.Undo(); grid.SwitchSheet(0); grid.Undo();
        grid.SwitchSheet(1); grid.SelectCell(1, 1);
        check('copied formula follows rename undo', [await grid.Paste(), grid.GetCell(1, 1).formula, grid.GetCell(1, 1).value],
            [true, '=Source!$A$1*2', 12]);
        grid.SwitchSheet(0); grid.Redo(); grid.SwitchSheet(1); grid.SelectCell(3, 3);
        check('copied formula follows rename redo', [await grid.Paste(), grid.GetCell(3, 3).formula, grid.GetCell(3, 3).value],
            [true, "='Renamed Source'!$A$1*2", 12]);
        let completeRead;
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
            readText: () => new Promise((resolve) => { completeRead = resolve; }),
        } });
        const pending = grid.Paste();
        grid.InstallWorkbook([createEmpty(20, 10)], ['Replacement'], 0, undefined);
        grid.MarkSaved();
        completeRead('999');
        check('late clipboard read cannot edit a replacement workbook',
            [await pending, grid.GetCell(0, 0).value, grid.IsModified(), grid.CanUndo()],
            [false, null, false, false]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        if (original) Object.defineProperty(navigator, 'clipboard', original); else delete navigator.clipboard;
        close();
    }
    return lines;
};

// Formatting and structure commands act on the whole selection: one change, one undo step.
checks['grid-selection-formatting'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        for (let row = 1; row <= 3; row++) for (let col = 1; col <= 3; col++) grid.SetCell(row, col, String(row + col));
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        const all = async (pick) => { const seen = []; for (let row = 1; row <= 3; row++) for (let col = 1; col <= 3; col++) seen.push(pick(await formatAt(grid, row, col))); return [...new Set(seen.map((v) => JSON.stringify(v ?? null)))].map((v) => JSON.parse(v)); };
        const outside = async (pick) => [pick(await formatAt(grid, 0, 1)), pick(await formatAt(grid, 4, 3)), pick(await formatAt(grid, 2, 4))];
        const step = async (label, apply, pick, wanted, plain) => {
            changes.length = 0;
            apply();
            check(`${label}: every cell of the selection`, [await all(pick), await outside(pick), changes.length], [[wanted], [plain, plain, plain], 1]);
            grid.Undo();
            check(`${label}: one undo step`, [await all(pick), grid.CanRedo()], [[plain], true]);
            grid.Redo();
            check(`${label}: and one redo`, await all(pick), [wanted]);
        };
        grid.SelectRange(1, 1, 3, 3);
        await step('bold', () => grid.ApplyFormat('bold'), (f) => f.Bold, true, false);
        await step('italic', () => grid.ApplyFormat('italic'), (f) => f.Italic, true, false);
        await step('text colour', () => grid.ApplyTextColor('#C00000'), (f) => f.TextColor ?? null, '#C00000', null);
        await step('fill colour', () => grid.ApplyBgColor('#FFF2CC'), (f) => f.BgColor ?? null, '#FFF2CC', null);
        await step('number format', () => grid.ApplyNumberFormat('0.00'), (f) => f.NumberFormat ?? null, '0.00', null);
        await step('alignment', () => grid.ApplyFormat('alignCenter'), (f) => f.HorizontalAlign ?? null, 'center', null);
        await step('border', () => grid.ApplyBorder('all', 'thin', '#000000'), (f) => [f.BorderTop?.Style ?? null, f.BorderRight?.Style ?? null], ['thin', 'thin'], [null, null]);
        // A toggle goes by the active cell: bold is on everywhere, so it goes off everywhere.
        changes.length = 0;
        grid.ApplyFormat('bold');
        check('bold again turns it off for the whole selection', [await all((f) => f.Bold), changes.length], [[false], 1]);
        // Mixed: the active cell decides. One cell is made bold; with a plain active cell all become bold ...
        grid.SelectCell(2, 2); grid.ApplyFormat('bold');
        grid.SelectRange(1, 1, 3, 3);
        check('a mixed selection', (await all((f) => f.Bold)).sort(), [false, true]);
        grid.ApplyFormat('bold');
        check('with a plain active cell a mixed selection becomes bold', await all((f) => f.Bold), [true]);
        // ... and with a bold active cell all become plain.
        grid.Undo();
        grid.SelectRange(2, 2, 1, 1);
        grid.ApplyFormat('bold');
        check('with a bold active cell it becomes plain', [await all((f) => f.Bold), active(grid)], [[false], [2, 2]]);
        check('the values are untouched', [await valueAt(grid, 1, 1), await valueAt(grid, 3, 3)], ['2', '6']);
        // Empty cells of the selection take the format too.
        grid.SelectRange(6, 0, 7, 1); grid.ApplyBgColor('#DDEBF7');
        check('empty cells are formatted too', [(await formatAt(grid, 6, 0)).BgColor, (await formatAt(grid, 7, 1)).BgColor], ['#DDEBF7', '#DDEBF7']);
        // A whole-sheet selection formats the cells that exist, at once.
        grid.SelectAll();
        const started = performance.now();
        grid.ApplyFormat('italic');
        check('a whole-sheet selection is formatted in one step, quickly', [(await formatAt(grid, 1, 1)).Italic, (await formatAt(grid, 6, 0)).Italic, performance.now() - started < 500], [true, true, true]);
        grid.Undo();
        check('and undone in one', [(await formatAt(grid, 1, 1)).Italic, (await formatAt(grid, 6, 0)).Italic], [true, false]);

        // Merge and unmerge are undoable.
        const { findMergedCell, CellAddress } = await import('/spreadsheet/Model.js');
        const merged = () => findMergedCell(new CellAddress(10, 1), grid.GetModel()) != null;
        const undoable = grid.CanUndo();
        changes.length = 0;
        grid.MergeCells(10, 0, 11, 2);
        check('merge', [merged(), changes.length], [true, 1]);
        grid.Undo(); check('undo of a merge', merged(), false);
        grid.Redo(); check('redo of a merge', merged(), true);
        grid.UnmergeCells(10, 1); check('unmerge', merged(), false);
        grid.Undo(); check('undo of an unmerge', merged(), true);
        void undoable;
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Read-only mode is a viewer: nothing can be changed and nothing invites it, while
// selecting, copying, switching sheets and moving with the keyboard work.
checks['grid-read-only'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { scroller, content } = parts(host);
        await grid.ImportXlsx((await bytesOf('tests/real-world/milk-street-office.xlsx')).slice());
        // (Not the parts of a header node that waits in the pool: those are in a hidden node.)
        const shown = (selector) => [...host.querySelectorAll(selector)].filter((node) => getComputedStyle(node).display !== 'none' && node.getClientRects().length > 0).length;
        check('an editor shows a fill handle, filter buttons and resize handles', [shown('.spreadsheet-fill-handle') > 0, shown('.col-filter-btn') > 0, shown('.col-resize-handle') > 0, shown('.row-resize-handle') > 0], [true, true, true, true]);
        grid.SetReadOnly(true);
        check('a viewer shows none of them', [shown('.spreadsheet-fill-handle'), shown('.col-filter-btn'), shown('.col-resize-handle'), shown('.row-resize-handle'), grid.IsReadOnly()], [0, 0, 0, 0, true]);
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        const before = grid.GetModel();
        // Selecting with the mouse and the keyboard.
        await mouse(grid, host, 'mousedown', 4, 2); await mouse(grid, host, 'mouseup', 4, 2);
        check('a click selects', [grid.GetSelectionRef(), active(grid)], ['C5', [4, 2]]);
        press(host, 'ArrowDown'); press(host, 'ArrowRight', { shiftKey: true });
        check('the keyboard moves and extends', grid.GetSelectionRef(), 'C6:D6');
        for (let i = 0; i < 60; i++) press(host, 'ArrowDown');
        check('and scrolls', [active(grid)[0], scroller.scrollTop > 0], [65, true]);
        press(host, 'Home', { ctrlKey: true });
        // Nothing edits.
        press(host, 'x'); press(host, 'F2'); press(host, 'Delete'); press(host, 'Enter');
        await mouse(grid, host, 'dblclick', 6, 1);
        check('typing, F2, Delete, Enter and a double click do not edit', [grid.IsEditing(), shown('.spreadsheet-edit-input')], [false, 0]);
        const header = host.querySelector('.col-header-text');
        header.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        const transfer = new DataTransfer(); transfer.setData('text/plain', 'pasted');
        content.dispatchEvent(clipboardEvent('paste', transfer));
        press(host, 'z', { ctrlKey: true }); press(host, 'd', { ctrlKey: true });
        check('a header click does not sort, a paste does not paste', [grid.GetModel().Cells === before.Cells, grid.GetModel().SortState == null, changes.length], [true, true, 0]);
        // Copy works; cut copies without clearing.
        grid.SelectRange(1, 1, 3, 4);
        const out = new DataTransfer();
        content.dispatchEvent(clipboardEvent('copy', out));
        check('copy works', out.getData('text/plain') === grid.GetSelectionText() && out.getData('text/html').startsWith('<table>'), true);
        const cutOut = new DataTransfer();
        content.dispatchEvent(clipboardEvent('cut', cutOut));
        check('a cut copies and leaves the cells', [cutOut.getData('text/plain') === out.getData('text/plain'), grid.GetModel().Cells === before.Cells], [true, true]);
        // Sheets.
        const visible = grid.SheetVisibility(), other = grid.GetSheetNames().findIndex((_, index) => index !== grid.GetActiveSheetIndex() && visible[index]);
        grid.SwitchSheet(other);
        // (That sheet has grouped rows: its cells begin after the outline strip.)
        check('sheets can be switched', [grid.GetActiveSheetIndex(), scroller.scrollTop, cellAtPoint(host, ...(await pointOfCell(grid, host, 0, 0)))], [other, 0, [0, 0]]);
        check('still no change was reported', changes.length, 0);
        // Back to an editor.
        grid.SetReadOnly(false);
        check('an editor again', [shown('.spreadsheet-fill-handle') > 0, shown('.col-resize-handle') > 0], [true, true]);
        grid.SelectCell(2, 30); press(host, 'q');
        check('typing edits again', grid.IsEditing(), true);
        grid.SetReadOnly(true);
        check('going read-only drops an entry in progress', [grid.IsEditing(), await valueAt(grid, 2, 30)], [false, '']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Scrolling a large sheet: frames stay short, nodes are reused, and a selection move draws no cell.
checks['grid-scroll-frames'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { scroller, content } = parts(host);
        await grid.ImportXlsx((await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx')).slice());
        grid.SetReadOnly(true);
        grid.SwitchSheet(grid.GetSheetNames().indexOf('Data'));
        const rows = grid.GetModel().RowCount;
        check('the sheet can be scrolled to its end', [rows > 70000, scroller.scrollHeight >= content.getBoundingClientRect().height - 1], [true, true]);
        const next = () => new Promise((resolve) => requestAnimationFrame(resolve));
        await next(); await next();
        // (The sheet was converted when it was shown: let the collection of what that left
        // behind pass before frames are timed.)
        if (window.collectGarbage) await window.collectGarbage();
        await next(); await next();
        // A mouse wheel's worth per frame, then a screenful per frame. What is measured is the work
        // of a frame - the time inside the grid's animation-frame callbacks plus the layout they leave
        // to do - not the time between frames: that also counts waiting for the display and for the
        // processor, and under load made this check fail more often than not.
        const run = async (step, frames) => {
            scroller.scrollTop = 0; await next(); await next();
            const raf = window.requestAnimationFrame;
            let spent = 0;
            window.requestAnimationFrame = (callback) => raf.call(window, (time) => { const start = performance.now(); try { callback(time); } finally { spent += performance.now() - start; } });
            const times = [];
            try {
                for (let i = 0; i < frames; i++) {
                    spent = 0;
                    scroller.scrollTop = (i + 1) * step;
                    // (The scroll event comes with the next frame; the grid draws in a callback of that frame.)
                    await new Promise((resolve) => raf.call(window, () => raf.call(window, resolve)));
                    const start = performance.now(); content.getBoundingClientRect(); scroller.scrollHeight;
                    times.push(spent + performance.now() - start);
                }
            } finally { window.requestAnimationFrame = raf; }
            return times;
        };
        const wheel = await run(120, 120), pages = await run(scroller.clientHeight - 30, 90);
        const median = (times) => [...times].sort((a, b) => a - b)[Math.floor(times.length / 2)];
        const stats = (times) => `median ${median(times).toFixed(1)} ms, mean ${(times.reduce((a, b) => a + b, 0) / times.length).toFixed(1)} ms, max ${Math.max(...times).toFixed(1)} ms, ${times.filter((t) => t > 33).length} over 33 ms`;
        note(`work per frame at 120 px per frame: ${stats(wheel)}; a screenful per frame: ${stats(pages)}`);
        // Half of the frames within half a refresh interval, and no more than one in ten over two intervals
        // (a collection of a heap holding this sheet's 600 000 cells takes some 200 ms whenever the browser
        // decides on one, and the machine may be busy): a grid that draws slower fails the first, one that
        // stalls every few frames the second.
        check('120 px per frame: the median frame takes under 8 ms of work', median(wheel) < 8, true);
        check('120 px per frame: no more than one frame in ten over 33 ms', wheel.filter((t) => t > 33).length <= wheel.length / 10, true);
        // (Every cell on screen is new in each of these frames.)
        check('a screenful per frame: the median frame takes under 16 ms of work', median(pages) < 16, true);
        check('a screenful per frame: no more than one frame in ten over 33 ms', pages.filter((t) => t > 33).length <= pages.length / 10, true);
        scroller.scrollTop = scroller.scrollHeight; await next(); await next();
        check('the last row is reached', cellAtPoint(host, 300, scroller.clientHeight - 30)?.[0] >= rows - 3, true);
        // Nodes are reused rather than made for every cell that scrolls by.
        const nodes = host.querySelectorAll('.spreadsheet-cell').length, visible = [...host.querySelectorAll('.spreadsheet-cell')].filter((node) => node.style.display !== 'none').length;
        check(`cell nodes are pooled (${visible} on display, ${nodes} in all)`, [nodes < 3 * visible + 700, visible > 100], [true, true]);
        // A selection move draws no cell.
        // (The cell is selected first: near the end of the grid that makes the grid longer, which draws the new rows.)
        grid.SelectCell(rows - 5, 2); await next(); await next();
        let drawn = 0;
        const observer = new MutationObserver((records) => { for (const record of records) if (record.target.closest && record.target.closest('.spreadsheet-cell')) drawn++; });
        observer.observe(content, { subtree: true, childList: true, attributes: true, characterData: true });
        grid.SelectCell(rows - 5, 2); press(host, 'ArrowUp'); press(host, 'ArrowRight', { shiftKey: true });
        await next();
        observer.disconnect();
        check('moving the selection touches no cell node', [drawn, grid.GetSelectionRef()], [0, `C${rows - 5}:D${rows - 5}`]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// A browser that raises no copy / cut / paste event on the grid (it is not an editable
// element): only the keys arrive. The grid then goes through the asynchronous clipboard.
// navigator.clipboard is replaced by a recorder, and no ClipboardEvent is dispatched.
checks['grid-clipboard-without-events'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    const original = Object.getOwnPropertyDescriptor(navigator, 'clipboard') ?? Object.getOwnPropertyDescriptor(Navigator.prototype, 'clipboard');
    try {
        const held = { text: '', html: '', writes: 0, reads: 0 };
        const stub = {
            write: async (items) => { held.writes++; held.text = ''; held.html = ''; for (const item of items) { if (item.types.includes('text/plain')) held.text = await (await item.getType('text/plain')).text(); if (item.types.includes('text/html')) held.html = await (await item.getType('text/html')).text(); } },
            writeText: async (text) => { held.writes++; held.text = text; held.html = ''; },
            read: async () => { held.reads++; const flavours = { 'text/plain': new Blob([held.text], { type: 'text/plain' }) }; if (held.html) flavours['text/html'] = new Blob([held.html], { type: 'text/html' }); return [new ClipboardItem(flavours)]; },
            readText: async () => { held.reads++; return held.text; },
        };
        Object.defineProperty(navigator, 'clipboard', { value: stub, configurable: true });
        const settle = async () => { await new Promise((resolve) => setTimeout(resolve, 30)); await frame(); };
        grid.SetCell(0, 0, 'Name'); grid.SetCell(0, 1, 'Qty'); grid.SetCell(1, 0, 'Apples'); grid.SetCell(1, 1, '=6*2');
        grid.SelectRange(0, 0, 1, 1);
        press(host, 'c', { ctrlKey: true }); await settle();
        check('Ctrl+C with no copy event: the cells are on the system clipboard', [held.writes, held.text, /<table>.*Apples.*<\/table>/.test(held.html)], [1, 'Name\tQty\nApples\t12', true]);
        grid.SelectCell(4, 0);
        press(host, 'v', { ctrlKey: true }); await settle();
        const { editText } = await import('/spreadsheet/Input.js');
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        check('Ctrl+V with no paste event: pasted from the system clipboard (the grid\'s own copy, with its formula)', [held.reads, await valueAt(grid, 4, 0), await valueAt(grid, 5, 1), editText(getCell(new CellAddress(5, 1), grid.GetModel()))], [1, 'Name', '12', '=6*2']);
        grid.Undo();
        check('as one undo step', await valueAt(grid, 4, 0), '');
        // Something another application put there.
        held.text = 'x\ty\r\n1\t2'; held.html = '';
        grid.SelectCell(8, 2);
        press(host, 'v', { ctrlKey: true }); await settle();
        check('Ctrl+V of another application\'s cells', [await valueAt(grid, 8, 2), await valueAt(grid, 9, 3)], ['x', '2']);
        // Cut.
        grid.SelectRange(0, 0, 0, 1);
        const writesBefore = held.writes;
        press(host, 'x', { ctrlKey: true }); await settle();
        check('Ctrl+X with no cut event: on the clipboard, and cleared', [held.writes - writesBefore, held.text, await valueAt(grid, 0, 0), await valueAt(grid, 0, 1)], [1, 'Name\tQty', '', '']);
        grid.Undo();
        check('the cut is one undo step', await valueAt(grid, 0, 1), 'Qty');
        // When the browser does raise the event, the event is the only path: nothing is written or read twice.
        const { content } = parts(host);
        grid.SelectRange(1, 0, 1, 1);
        const before = { writes: held.writes, reads: held.reads };
        press(host, 'c', { ctrlKey: true });
        const transfer = new DataTransfer();
        content.dispatchEvent(clipboardEvent('copy', transfer));
        await settle();
        check('with a copy event the asynchronous clipboard is not used', [held.writes - before.writes, transfer.getData('text/plain')], [0, 'Apples\t12']);
        grid.SelectCell(12, 0);
        press(host, 'v', { ctrlKey: true });
        const incoming = new DataTransfer(); incoming.setData('text/plain', 'from the event');
        content.dispatchEvent(clipboardEvent('paste', incoming));
        await settle();
        check('with a paste event it is pasted once, from the event', [held.reads - before.reads, await valueAt(grid, 12, 0), grid.CanUndo() && (grid.Undo(), await valueAt(grid, 12, 0))], [0, 'from the event', '']);
        // Read-only: copy works, cut copies, paste does nothing.
        grid.SetReadOnly(true);
        grid.SelectRange(1, 0, 1, 1);
        const cells = grid.GetModel().Cells;
        press(host, 'x', { ctrlKey: true }); await settle();
        grid.SelectCell(20, 0); held.text = 'nope';
        press(host, 'v', { ctrlKey: true }); await settle();
        check('read-only: a cut copies, a paste is ignored', [grid.GetModel().Cells === cells, await valueAt(grid, 20, 0)], [true, '']);
        // While a cell is edited the keys belong to the edit box.
        grid.SetReadOnly(false);
        grid.SelectCell(1, 0); press(host, 'F2');
        const writes = held.writes;
        press(host, 'c', { ctrlKey: true }); await settle();
        check('while editing the grid does not touch the clipboard', held.writes - writes, 0);
        grid.CancelEdit();
        // A browser without any clipboard access: the grid's own clipboard still works.
        Object.defineProperty(navigator, 'clipboard', { value: undefined, configurable: true });
        grid.SelectRange(1, 0, 1, 1);
        press(host, 'c', { ctrlKey: true }); await settle();
        grid.SelectCell(22, 0);
        press(host, 'v', { ctrlKey: true }); await settle();
        check('without navigator.clipboard: copy and paste within the grid', [await valueAt(grid, 22, 0), await valueAt(grid, 22, 1)], ['Apples', '12']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    if (original) Object.defineProperty(navigator, 'clipboard', original); else delete navigator.clipboard;
    close();
    return lines;
};

// The grid draws a sheet as Excel 2019 does (tests/test-documents-excel/synthetic-grid-parity.xlsx, made by
// tools/excel-parity/make-grid-parity-fixture.py; each rule found on a real-world model and checked in Excel).
checks['grid-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-grid-parity.xlsx')).slice());
        grid.SwitchSheet(0); await frame(); await everyObjectSeen(host);
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const style = (element) => getComputedStyle(element);
        // A run with properties of its own is in its own font: not bold in a bold cell.
        const runs = [...node(0, 0).querySelectorAll('span')].map((span) => [span.textContent, style(span).fontWeight]);
        check('a rich run without <b/> is not bold in a bold cell', runs.find(([text]) => text === 'italic only')?.[1], '400');
        // `ht` without `customHeight` is not the row's height: Excel fits the row to its cells on opening.
        check('a row of ht 15 (no customHeight) of Arial 10 text is fitted: 17 px', Math.round(node(1, 0).getBoundingClientRect().height), 17);
        check('... the rows round it the sheet\'s 12.75 pt (17 px)', Math.round(node(2, 0).getBoundingClientRect().height), 17);
        // Conditional formats with lines: a side named without a style takes the edge's lines away.
        const edges = (row, col) => [...node(row, col).querySelectorAll('.spreadsheet-edge')].map((edge) => edge.getAttribute('style'));
        check('a rule of <left/><right/><top/><bottom/>: no line above or left of C4', edges(3, 2).filter((s) => /border-(top|left):\s*\d/.test(s)).length, 0);
        check('... nor left of D4 (C4\'s right) nor above C5 (its bottom)', [edges(3, 3).some((s) => /border-left/.test(s)), edges(4, 2).some((s) => /border-top/.test(s))], [false, false]);
        check('... the lines of the cells round the block stay', [edges(3, 1).some((s) => /border-left/.test(s)), edges(2, 2).some((s) => /border-top/.test(s))], [true, true]);
        check('a rule\'s dashed left and thick red top are E4\'s, over its neighbours\' thin lines',
            [edges(3, 4).some((s) => /border-left:\s*1px dashed rgb\(0, 0, 255\)/.test(s)), edges(3, 4).some((s) => /border-top:\s*3px solid rgb\(255, 0, 0\)/.test(s))], [true, true]);
        // A dxf's automatic font colour is black.
        check('a rule\'s auto font colour on a black fill: black on black', [style(node(6, 0)).color, style(node(6, 0)).backgroundColor], ['rgb(0, 0, 0)', 'rgb(0, 0, 0)']);
        // Wrapped text is set in 3 px from the cell's sides.
        check('wrapped text: 3 px from either side', [style(node(8, 1)).paddingLeft, style(node(8, 1)).paddingRight], ['3px', '3px']);
        // Indented text that runs over its neighbours has room for its indent.
        const clip = /inset\(0px (-?\d+)px/.exec(node(12, 0).getAttribute('style') ?? '');
        check('an indented text runs on over the cells beside it, as far as it needs', clip !== null && Number(clip[1]) <= -node(12, 1).getBoundingClientRect().width, true);
        // A column's default style is every empty cell's of it.
        check('an empty cell of a column with a default style has its fill', [style(node(29, 9)).backgroundColor, style(node(29, 4)).backgroundColor], ['rgb(217, 217, 217)', 'rgb(255, 255, 255)']);
        // Drawings: hidden ones and an empty, invisible text box are not drawn (nor a placeholder); a linked picture is.
        const objects = [...host.querySelectorAll('.spreadsheet-drawing')];
        check('no placeholder frame, no picture of the hidden ones', [host.querySelectorAll('.spreadsheet-object-label').length, host.querySelectorAll('img.spreadsheet-picture').length], [0, 0]);
        check('hidden and empty objects draw nothing', objects.filter((item) => item.querySelector('svg') && !item.querySelector('svg text') && !item.querySelector('svg rect')).length, 3);
        const camera = objects.find((item) => item.querySelector('svg text'));
        check('a linked picture shows the cells of its range', camera ? [...camera.querySelectorAll('svg text')].map((text) => text.textContent) : null, ['Linked', '7.27', 'Picture', '12']);
        check('... in their fonts', camera ? [...camera.querySelectorAll('svg text')].map((text) => text.getAttribute('font-weight') ?? '') : null, ['', '', 'bold', '']);
        grid.SetCell(9, 1, '8.5'); await frame(); await everyObjectSeen(host);
        const again = [...host.querySelectorAll('.spreadsheet-drawing svg text')].map((text) => text.textContent);
        check('... as they are now', again.includes('8.5'), true);
        // At 70 %: a medium line is 2 screen pixels; two filled cells meet.
        grid.SwitchSheet(1); await frame();
        const line = [...node(0, 0).querySelectorAll('.spreadsheet-edge'), ...(node(1, 0)?.querySelectorAll('.spreadsheet-edge') ?? [])].map((edge) => edge.getBoundingClientRect()).find((box) => box.height > 0.5 && box.width > 20);
        check('a medium line at 70 %: 2 px', line ? Math.round(line.height) : null, 2);
        const a = node(2, 0).getBoundingClientRect(), b = node(2, 1).getBoundingClientRect();
        check('filled cells at 70 % meet at a whole screen pixel', [a.right === b.left, Number.isInteger(b.left - host.getBoundingClientRect().left)], [true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// A cell's text as Excel 2019 sizes and places it (synthetic-text-zoom, checked in Excel; ExcelText): at a whole
// pixel size on screen - 15 px for 11 pt at 100 %, 12 px at both 80 % and 85 % -, in GDI's advance widths (twenty
// zeros of Calibri 11 are 140 px at 100 %, 120 at 80 % and 85 %), in Small Fonts under 8 points on screen (67 %;
// Arial 7 at 100 %); 2 px in from the cell's left side, 3 px from its right; a number fitted at the zoom.
checks['grid-text-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-text-zoom.xlsx')).slice());
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        // (The text alone: a cell of a number stored as text has the green triangle of Excel's error check too.)
        const textBox = (element) => {
            const range = document.createRange(), walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
            let left = Infinity, right = -Infinity, top = Infinity, bottom = -Infinity;
            for (let node = walker.nextNode(); node; node = walker.nextNode()) {
                range.selectNodeContents(node);
                const r = range.getBoundingClientRect();
                if (r.width > 0) { left = Math.min(left, r.left); right = Math.max(right, r.right); top = Math.min(top, r.top); bottom = Math.max(bottom, r.bottom); }
            }
            return { left, right, top, bottom, width: right - left, height: bottom - top };
        };
        const near = (value, expected, slack = 0.75) => Math.abs(value - expected) <= slack ? expected : Math.round(value * 100) / 100;
        const sheets = grid.GetSheetNames();
        const show = async (name) => { grid.SwitchSheet(sheets.indexOf(name)); await frame(); await frame(); };
        // (Where the browser has no Calibri the grid draws the next family as tall as Calibri's letters: smaller.)
        const hasCalibri = (() => { const c = document.createElement('canvas').getContext('2d'), t = 'mmmmmmmmmmlli WQ 0123456789'; c.font = '72px monospace'; const a = c.measureText(t).width; c.font = "72px 'Calibri', monospace"; return c.measureText(t).width !== a; })();
        const calibriScale = hasCalibri ? 1 : 0.4644 / 0.5186;
        for (const [name, zoom, px, zeros] of [['Z100', 1, 15, 140], ['Z80', 0.8, 12, 120], ['Z85', 0.85, 12, 120], ['Z67', 0.67, 10, 120]]) {
            await show(name);
            const size = parseFloat(getComputedStyle(node(0, 1)).fontSize);
            // (Small Fonts' text - 67 % - in a stand-in of its own, at its size.)
            check(`${name}: the text is drawn at ${px} px on screen`, near(size * zoom / (name === 'Z67' ? 1 : calibriScale), px, 0.05), px);
            check(`${name}: twenty zeros are ${zeros} px wide, as Excel draws them`, near(textBox(node(2, 1)).width, zeros), zeros);
            const cell = node(0, 1).getBoundingClientRect(), right = node(1, 1).getBoundingClientRect();
            check(`${name}: a left-aligned line starts 2 px x the zoom into the cell`, near(textBox(node(0, 1)).left - cell.left, 2 * zoom), 2 * zoom);
            check(`${name}: a right-aligned line ends 3 px x the zoom before the cell's end`, near(right.right - textBox(node(1, 1)).right, 3 * zoom), 3 * zoom);
        }
        await show('Z80');
        const pangram80 = textBox(node(0, 1)).width;
        await show('Z85');
        check('80 % and 85 % draw a line of Calibri 11 alike', near(textBox(node(0, 1)).width, pangram80, 0.05), pangram80);
        await show('Arial');
        check('Arial 10 at 100 %: 13 px; twenty zeros 140 px', [near(parseFloat(getComputedStyle(node(0, 1)).fontSize), 13, 0.05), near(textBox(node(0, 1)).width, 140)], [13, 140]);
        check('Arial 7 at 100 %: 9 px, as wide as Small Fonts draws it (twenty zeros 100 px)', near(textBox(node(1, 1)).width, 100), 100);
        await show('Numbers');
        check('numbers at 75 % as Excel shows them there', [0, 1].map((row) => [1, 2, 3].map((col) => node(row, col)?.textContent ?? '')),
            [['#####', '1E+06', '1234567'], ['#####', '#######', '#########']]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Segoe UI Black (acre-property-acquisition's font) measured as Excel 2019 draws it (synthetic-font-faces,
// make-font-faces-fixture.py; Range.Text read from Excel): $1,592,811 in columns of 40 .. 85 px shows as the number
// or as many #s as Excel shows; a number's room in its cell for 10 fonts at 9 zooms; Small Fonts' threshold.
checks['font-faces-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-font-faces.xlsx')).slice());
        grid.SwitchSheet(0); await frame(); await frame();
        const text = (row, col) => [...host.querySelectorAll(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`)].find((node) => node.style.display !== 'none')?.textContent;
        const shown = (row) => Array.from({ length: 10 }, (_, i) => text(row, i + 1));
        const n = '$1,592,811', h = (count) => '#'.repeat(count);
        check('Segoe UI Black 9', shown(0), [h(5), h(5), h(6), h(7), h(7), h(8), n, n, n, n]);
        check('Segoe UI Black 10', shown(1), [h(4), h(5), h(5), h(6), h(6), h(7), h(8), n, n, n]);
        // (Its digit 9 px: the room is the cell less 7 px, `CellText.fittedRoom`.)
        check('Segoe UI Black 11', shown(2), [h(3), h(4), h(4), h(5), h(5), h(6), h(7), h(7), h(8), n]);
        check('Segoe UI 10', shown(3), [h(4), h(5), h(5), h(6), h(6), h(7), n, n, n, n]);
        // A number's room: the cell less 2 x ceil(d / 4) + 1 px, d the font's widest digit as drawn (synthetic-font-faces
        // "Room*": 10 fonts at 9 zooms, every cell's Range.Text from Excel in synthetic-font-faces.excel.json).
        const excel = JSON.parse(new TextDecoder().decode(await bytesOf('tests/test-documents-excel/synthetic-font-faces.excel.json')));
        const scroller = host.querySelector('.spreadsheet-container');
        for (const [sheet, rows] of Object.entries(excel)) {
            grid.SwitchSheet(grid.GetSheetNames().indexOf(sheet)); await frame(); await frame();
            const seen = new Map();
            for (const top of [0, 300, 600]) for (let left = 0; left < 8000; left += 600) {
                scroller.scrollTop = top; scroller.scrollLeft = left; await frame(); await frame();
                for (const node of host.querySelectorAll('.spreadsheet-cell')) {
                    if (node.style.display === 'none') continue;
                    seen.set(node.dataset.row + ',' + node.dataset.col, node.textContent);
                }
            }
            scroller.scrollLeft = 0; scroller.scrollTop = 0;
            let same = 0, total = 0; const wrong = [];
            for (const [row, texts] of rows) texts.forEach((want, i) => {
                const got = seen.get(row + ',' + (i + 1)) ?? '?';
                const norm = typeof want === 'number' ? '#'.repeat(want) : want;
                total++;
                if (got.trim() === norm.trim()) same++; else if (wrong.length < 12) wrong.push(`${row},${i + 1}: ${got} want ${norm}`);
            });
            check(`${sheet}: every number's text as Excel's (${wrong.join('; ')})`, same, total);
        }
        // Small Fonts under 8 points on screen, taken to a twentieth of a point (synthetic-small-fonts, from Excel's
        // pictures: Calibri / Arial 7 .. 16 at 23 zooms): 14 pt at 57 % (7.98) and 10.5 at 76 % are TrueType, 15 at 53 %
        // (7.95) and 11 at 72 % (7.92) Small Fonts.
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-small-fonts.xlsx')).slice());
        const names = grid.GetSheetNames();
        const small = async (sheet, row) => {
            grid.SwitchSheet(names.indexOf(sheet)); await frame(); await frame();
            const cell = [...host.querySelectorAll(`.spreadsheet-cell[data-row="${row}"][data-col="1"]`)].find((node) => node.style.display !== 'none');
            return /Verdana/.test(cell?.style.fontFamily ?? '');
        };
        // (Rows: Calibri 7, 7.5, 8, 8.5, 9, 10, 10.5, 11, 12, 13, 14, 15, 16.)
        check('Small Fonts: Calibri 14 / 13 at 57 %, 15 / 16 at 53 %, 11 / 12 at 72 %, 10.5 / 10 at 76 %',
            [await small('Z57', 10), await small('Z57', 9), await small('Z53', 11), await small('Z53', 12), await small('Z72', 7), await small('Z72', 8), await small('Z76', 6), await small('Z76', 5)],
            [false, true, true, false, true, false, false, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Wrapped text broken into lines where Excel 2019 breaks it on screen (synthetic-wrap-budget, make-wrap-fixture.py,
// each column's lines read from Excel's pictures): a line fits while it is no wider than the cell on screen less 5 px,
// in Excel's widths at the zoom (bold Calibri 11 "Releasing Cost", 90 px, one line from a 95-px column).
checks['wrap-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-wrap-budget.xlsx')).slice());
        const sheets = grid.GetSheetNames();
        const linesOf = async (col) => {
            grid.SelectCell(3, col); await frame();
            // (The grid keeps the nodes it frees, hidden: the one shown.)
            const cell = [...host.querySelectorAll(`.spreadsheet-cell[data-row="0"][data-col="${col}"]`)].find((node) => node.style.display !== 'none');
            return [...(cell?.querySelectorAll(':scope > span') ?? [])].map((span) => span.textContent);
        };
        const columns = async (from, to) => { const all = []; for (let col = from; col <= to; col++) all.push(await linesOf(col)); return all; };
        const repeat = (n, value) => Array.from({ length: n }, () => value);
        for (const [name, last, ruled] of [
            ['B100', 18, [...repeat(2, ['Releasing', 'Cost per', 'Unit']), ...repeat(9, ['Releasing', 'Cost per Unit']), ...repeat(7, ['Releasing Cost', 'per Unit'])]],
            ['R85', 20, [...repeat(12, ['The quick', 'brown fox', 'jumps']), ...repeat(8, ['The quick brown', 'fox jumps'])]],
            ['R70', 20, [...repeat(17, ['The quick brown', 'fox jumps']), ...repeat(3, ['The quick brown fox', 'jumps'])]],
        ]) {
            grid.SwitchSheet(sheets.indexOf(name)); await frame(); await frame();
            check(`${name}: the lines of each column (a pixel wider each) as Excel draws them`, await columns(1, last), ruled);
        }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Rows and columns as Excel 2019 opens them (tools/excel-parity/make-row-height-fixtures.py; every number read from
// Excel 2019 on the fixtures): a height's pixels, fractional font sizes, the rows of no height of their own a sheet
// hides (zeroHeight), columns hidden to XFD, a row's fonts meeting at their baseline, merged cells.
checks['rows-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const heights = (n) => Array.from({ length: n }, (_, row) => grid.GetRowHeight(row));
    const sheet = async (index) => { grid.SwitchSheet(index); await frame(); };
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-row-heights.xlsx')).slice());
        await sheet(0);
        const custom = heights(59);
        check('custom heights 15 .. 15.65 pt are 20 px, 15.7 .. 16 21 (a twentieth short of a pixel is the pixel)', [custom.slice(0, 14).every((h) => h === 20), custom.slice(14, 21).every((h) => h === 21)], [true, true]);
        check('... 19.5 .. 20.15 pt 26 px, 20.2 .. 20.5 27', [custom.slice(21, 35).every((h) => h === 26), custom.slice(35, 42).every((h) => h === 27)], [true, true]);
        check('... 12, 12.3, 12.375, 12.5, 12.7, 14.25, 14.4, 15.6, 15.95, 0.75, 1, 2, 30.1, 30.4, 30.6, 30.9, 409 pt', custom.slice(42), [16, 16, 16, 16, 17, 19, 19, 20, 21, 1, 1, 2, 40, 40, 40, 41, 545]);
        const defaults = [];
        for (const index of [1, 2, 3, 4, 5]) { await sheet(index); defaults.push(grid.GetRowHeight(1)); }
        check('a sheet\'s own default of 15.7, 15.95, 20, 16, 14.4 pt', defaults, [21, 21, 26, 21, 19]);
        await sheet(6);
        check('fitted rows of a sheet of 30 pt of its own: the fonts\' lines, not the default; the rows past them 40 px', heights(7), [53, 20, 20, 35, 20, 20, 40]);
        await sheet(7);
        check('... every column in Calibri 8: 15 px rows, empty or not', heights(7), [53, 15, 20, 35, 15, 15, 21]);
        await sheet(8);
        check('fractional sizes: Calibri 10.5 / 7.5 / 13.5 / 9.5 / 11.5, Arial 10.5 / 7.5 / 13.5 / 11.5, Times 10.5 / 13.5, Calibri 15.5, Arial 17.5',
            heights(13), [20, 20, 24, 20, 20, 20, 20, 23, 20, 20, 23, 28, 29]);
        check('... a size of 10.5 pt is kept', grid.GetCell(0, 0).format.fontSize, 10.5);
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-hidden-axes.xlsx')).slice());
        await sheet(0);
        check('<col min="4" max="16384" hidden>: D .. XFD hidden', [0, 2, 3, 1023, 1045, 4999, 16383].map((col) => grid.GetColumnWidth(col)), [64, 64, 0, 0, 0, 0, 0]);
        await sheet(1);
        check('defaultColWidth="0": the columns of no <col> hidden', [0, 2, 3, 100].map((col) => grid.GetColumnWidth(col)), [70, 70, 0, 0]);
        await sheet(2);
        check('zeroHeight: the rows not listed hidden, between and past the listed ones', [...heights(12), grid.GetRowHeight(99)], [20, 20, 0, 20, 0, 40, 20, 0, 20, 35, 0, 0, 0]);
        grid.UnhideRows('3:3'); await frame();
        check('... a row of them shown again at the default height', grid.GetRowHeight(2), 20);
        await sheet(3);
        check('zeroHeight with a default of its own: listed rows of 15.95 pt 21 px, a row of no ht fitted', heights(6), [21, 21, 0, 20, 0, 0]);
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-row-merged.xlsx')).slice());
        await sheet(0);
        check('merged cells: across columns wrapped a line (its descent only), a 20 pt line counts, over three rows nothing',
            heights(11), [20, 20, 35, 20, 20, 20, 20, 20, 35, 200, 23]);
        await sheet(1);
        check('a column of Calibri 20 under wrapped Calibri 11: its descent below the text\'s (7, 3 lines, plain, 2 lines)', heights(4), [143, 63, 35, 43]);
        await sheet(2);
        check('... of Calibri 8: no more', heights(4), [140, 60, 20, 40]);
        await sheet(3);
        check('wrapped at spaces, not at no-break spaces (5 lines; 4, broken in the words; 4; 3)', heights(4), [100, 80, 80, 60]);
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-row-baseline.xlsx')).slice());
        await sheet(0);
        check('cells at the bottom meet at their baseline, centred or at the top only their height', heights(11), [21, 21, 20, 20, 20, 20, 20, 63, 60, 60, 63]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Grid chrome Excel 2019 draws (tools/excel-parity/make-error-check-fixtures.py, checked in Excel with error-probe.vbs and
// shoot-sheets.ps1): the green triangle of a background error check at a cell's top left - not where <ignoredErrors>
// says so - and the line along the edges of frozen panes.
checks['grid-chrome-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    const marked = () => [...host.querySelectorAll('[data-error-check]')].map((mark) => {
        const cell = mark.closest('.spreadsheet-cell');
        return String.fromCharCode(65 + Number(cell.getAttribute('data-col'))) + (Number(cell.getAttribute('data-row')) + 1) + ' ' + mark.getAttribute('data-error-check');
    }).sort();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-error-checks.xlsx')).slice());
        grid.SwitchSheet(0); await frame();
        check('the cells Excel 2019 marks, and why', marked(), ['A1 numberStoredAsText', 'A10 numberStoredAsText', 'A11 numberStoredAsText', 'A2 numberStoredAsText',
            'A4 numberStoredAsText', 'A5 numberStoredAsText', 'A6 numberStoredAsText', 'A7 numberStoredAsText', 'A8 numberStoredAsText', 'A9 numberStoredAsText',
            'B5 formulaRange', 'F10 formula', 'G20 formulaRange', 'J3 formula', 'K1 evalError', 'L1 twoDigitTextYear', 'L2 twoDigitTextYear']);
        const mark = host.querySelector('[data-error-check]');
        check('... a green triangle of 6 px inside the grid lines', [getComputedStyle(mark).borderTopColor, getComputedStyle(mark).borderTopWidth, mark.style.top, mark.style.left], ['rgb(0, 128, 0)', '6px', '1px', '1px']);
        grid.SwitchSheet(1); await frame();
        check('... none where <ignoredErrors> turns the check off', marked(), []);
        const shownLines = () => [...host.querySelectorAll('.spreadsheet-freeze-line')].filter((line) => line.style.display !== 'none').map((line) => getComputedStyle(line).backgroundColor);
        check('no frozen panes, no line', shownLines(), []);
        grid.FreezeAt(2, 1); await frame();
        check('frozen panes: a line along each edge, #9B9B9B', shownLines(), ['rgb(155, 155, 155)', 'rgb(155, 155, 155)']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Charts as Excel 2019 draws them (tools/excel-parity/make-chart-fixtures.py, checked in Excel with chart-probe.vbs):
// the value axes' scales, a date axis, a scatter chart's x values of two columns, a pie's labels, typefaces.
checks['charts-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-chart-parity.xlsx')).slice());
        await everyObjectSeen(host);
        const texts = [...host.querySelectorAll('svg text')].map((t) => t.textContent);
        check('the scales Excel chooses (0 .. 10 by 5 on a short axis, 44 .. 62, 300,000, x to 150)', ['10', '44', '62', '300,000', '150'].map((t) => texts.includes(t)), [true, true, true, true, true]);
        check('a date axis labelled every two years', ['2024', '2026', '2032'].map((t) => texts.includes(t)), [true, true, true]);
        check('pie labels a line a part, broken to a fifth of the chart', [['Two', 'BR/One', 'BA', '14%'].every((t) => texts.includes(t)), texts.some((t) => t.includes('; '))], [true, false]);
        const faces = [...host.querySelectorAll('svg g[font-family]')].map((g) => g.getAttribute('font-family').split(',')[0]);
        check('chart text in its typeface (Arial, Segoe UI)', [faces.includes("'Arial'"), faces.includes("'Segoe UI'")], [true, true]);
        // synthetic-chart-lines.xlsx: lines and legend keys in whole pixels of their width, category labels
        // broken, turned and left out as Excel 2019 sets them.
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-chart-lines.xlsx')).slice());
        await everyObjectSeen(host);
        // (Each chart's SVG, as the grid shows it when its cells are in view.)
        const seen = new Map();
        for (const [row, col] of [[0, 12], [19, 12], [57, 12], [76, 12], [0, 30], [0, 34], [19, 30], [19, 34]]) {
            grid.SelectCell(row, col); await frame();
            for (let i = 0; i < 30 && host.querySelector('.spreadsheet-object-pending'); i++) await frame();
            for (const svg of host.querySelectorAll('svg')) seen.set(svg.outerHTML, svg);
        }
        const svgs = [...seen.values()];
        const all = (selector) => svgs.flatMap((svg) => [...svg.querySelectorAll(selector)]);
        // (Its lines say nothing of smoothing: Excel 2019 smooths them, so they are paths.)
        const widths = all('polyline, path').map((l) => l.getAttribute('stroke-width'));
        check('lines of 0.25 .. 4.5 pt: 1, 1, 1, 1, 2, 3, 4, 6 px', ['1', '2', '3', '4', '6'].map((w) => widths.includes(w)), [true, true, true, true, true]);
        const keys = all('line').filter((l) => Math.abs(Number(l.getAttribute('x2')) - Number(l.getAttribute('x1')) - 26) < 0.01).map((l) => l.getAttribute('stroke-width'));
        check('legend keys 26 px long, as wide as their lines', ['1', '4', '6'].map((w) => keys.includes(w)), [true, true, true]);
        const labels = all('text');
        const turned = (angle) => labels.filter((t) => (t.getAttribute('transform') ?? '').startsWith(`rotate(${angle} `)).map((t) => t.textContent);
        check('labels turned -45 (every second), -90 (all twelve)', [turned(-45).includes('Category 3'), turned(-45).includes('Category 2'), turned(-90).includes('Category 12')], [true, false, true]);
        const texts2 = labels.map((t) => t.textContent);
        check('"Cat 1" broken at its space; 60 numbers every second', [texts2.includes('Cat'), texts2.includes('59'), texts2.includes('60')], [true, true, false]);
        // synthetic-chart-looks.xlsx: shapes over a chart (c:userShapes), a title of runs, a legend's own box, the
        // labels' typefaces.
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-chart-looks.xlsx')).slice());
        await everyObjectSeen(host);
        const looks = new Map();
        for (const [row, col] of [[0, 6], [28, 6], [0, 14], [0, 22], [19, 22], [0, 30]]) {
            grid.SelectCell(row, col); await frame();
            for (let i = 0; i < 30 && host.querySelector('.spreadsheet-object-pending'); i++) await frame();
            for (const svg of host.querySelectorAll('svg')) looks.set(svg.outerHTML, svg);
        }
        const lsvgs = [...looks.values()];
        const ltexts = lsvgs.flatMap((svg) => [...svg.querySelectorAll('text, tspan')].map((t) => t.textContent));
        check('a chart\'s shapes drawn over it, one linked to its cell', ['Centre', 'Next Look', 'Fashions Direct 71%'].map((t) => ltexts.includes(t)), [true, true, true]);
        const runs = lsvgs.flatMap((svg) => [...svg.querySelectorAll('text')]).filter((t) => ['Sales ', 'by ', 'Region'].includes(t.textContent) && !t.getAttribute('transform')).map((t) => [t.textContent, t.getAttribute('font-size'), t.getAttribute('fill')]);
        check('a title of runs, each its size and colour', runs.slice(0, 3), [['Sales ', '24', '#C00000'], ['by ', '16', '#595959'], ['Region', '24', '#0070C0']]);
        const legendBox = lsvgs.flatMap((svg) => [...svg.querySelectorAll('rect')]).some((r) => r.getAttribute('fill') === '#FFF2CC' && r.getAttribute('stroke') === '#C00000');
        check('a legend\'s own fill and outline', legendBox, true);
        const faces2 = lsvgs.flatMap((svg) => [...svg.querySelectorAll('g[font-family]')]).map((g) => g.getAttribute('font-family').split(',')[0]);
        check('Tw Cen MT and Garamond labels', [faces2.includes("'Tw Cen MT'"), faces2.includes("'Garamond'")], [true, true]);
        // synthetic-chartex-looks.xlsx (Excel's own chartEx charts): a histogram's bins, a treemap's labels.
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-chartex-looks.xlsx')).slice());
        const exTexts = new Set();
        for (const sheet of ['Hist', 'Tree']) {
            grid.SwitchSheet(grid.GetSheetNames().indexOf(sheet)); await frame();
            await everyObjectSeen(host);
            for (const [row, col] of [[0, 13], [15, 13], [30, 13], [45, 13], [60, 13], [75, 13]]) {
                grid.SelectCell(row, col); await frame();
                for (let i = 0; i < 30 && host.querySelector('.spreadsheet-object-pending'); i++) await frame();
                for (const t of host.querySelectorAll('svg text')) exTexts.add(t.textContent);
            }
        }
        check('chartEx: histogram bins of Scott\'s width ([11.87, 24.87]), under- and overflow bins, a treemap\'s labels',
            ['[11.87, 24.87]', '≤ 100', '> 800', 'Milk', 'Dairy', 'A, 50'].map((t) => exTexts.has(t)), [true, true, true, true, true, true]);
        // synthetic-chartsheets.xlsx: a chart sheet as Excel shows one - no cells, its page fitted to the window, the
        // chart in its middle.
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-chartsheets.xlsx')).slice());
        grid.SwitchSheet(grid.GetSheetNames().indexOf('Portrait')); await frame(); await frame();
        const layer = host.querySelector('.spreadsheet-chartsheet');
        const chartBox = layer?.firstElementChild?.getBoundingClientRect();
        const layerBox = layer?.getBoundingClientRect();
        check('a chart sheet: a grey window over the grid, the chart in its middle, the portrait page fitted to it',
            [!!layer && getComputedStyle(layer).display !== 'none', !!chartBox && Math.abs((chartBox.left + chartBox.right) / 2 - (layerBox.left + layerBox.right) / 2) < 2,
             !!chartBox && chartBox.height > layerBox.height * 0.9 && chartBox.height <= layerBox.height, !!layer?.querySelector('svg')], [true, true, true, true]);
        grid.SwitchSheet(grid.GetSheetNames().indexOf('Data')); await frame();
        check('... gone on a worksheet', !layer || getComputedStyle(layer).display === 'none', true);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Sizes and theme fills as Excel 2019 shows them (tools/excel-parity/make-sizing-fixtures.py; every number
// read from Excel 2019 on the fixtures): column widths in the Normal font's maximum digit width, row heights
// of its font, the rows Excel fits on opening, a row's default style to the last column, a shape's theme fill.
checks['sizing-excel-parity'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    const sheetXml = async (bytes, n) => (await window.JSZip.loadAsync(bytes)).files[`xl/worksheets/sheet${n}.xml`].async('string');
    const cols = (xml) => [...xml.matchAll(/<col\b[^>]*>/g)].map((m) => m[0]);
    try {
        const widthsFile = await bytesOf('tests/test-documents-excel/synthetic-font-widths.xlsx');
        await grid.ImportXlsx(widthsFile.slice());
        grid.SwitchSheet(0); await frame();
        const widths = (n) => Array.from({ length: n }, (_, col) => grid.GetColumnWidth(col));
        const heights = (n) => Array.from({ length: n }, (_, row) => grid.GetRowHeight(row));
        // Arial MT 12 (Excel's Arial): a maximum digit width of 9 px.
        check('Normal font Arial MT 12: columns of 1, 2.1, 8.9, 10, 20, 38.8, 0.5, 2 and none are Excel\'s pixels', widths(9), [9, 19, 80, 90, 180, 349, 4, 18, 80]);
        check('... rows a line of Arial 12 (15 pt), whatever defaultRowHeight says', heights(4), [20, 20, 20, 20]);
        grid.SwitchSheet(1); await frame();
        check('defaultColWidth 20: 180 px', grid.GetColumnWidth(5), 180);
        grid.SwitchSheet(2); await frame();
        check('baseColWidth 12: 12 x 9 + 7 up to a multiple of 8, 120 px', grid.GetColumnWidth(5), 120);
        grid.SwitchSheet(3); await frame();
        check('defaultRowHeight 30 without customHeight: the font\'s 20 px', heights(5), [20, 20, 20, 20, 20]);
        grid.SwitchSheet(4); await frame();
        check('defaultRowHeight 30 customHeight: 40 px', heights(5), [40, 40, 40, 40, 40]);
        // Unedited, the widths go back as they were; a width set goes back in the font's digits.
        grid.SwitchSheet(0); await frame();
        const unedited = await grid.ExportXlsx();
        check('unedited: the <col> widths are the file\'s', cols(await sheetXml(unedited, 1)), cols(await sheetXml(widthsFile, 1)));
        grid.SetColumnWidth(1, 45); await frame();
        const edited = await grid.ExportXlsx();
        const b = cols(await sheetXml(edited, 1)).find((tag) => /min="2"/.test(tag)) ?? '';
        check('a column set to 45 px is written 5 digits of 9 px', /width="5"/.test(b), true);
        check('... the others keep theirs', cols(await sheetXml(edited, 1)).filter((tag) => !/min="2"/.test(tag)), cols(await sheetXml(widthsFile, 1)).filter((tag) => !/min="2"/.test(tag)));
        await grid.ImportXlsx(edited.slice()); grid.SwitchSheet(0); await frame();
        check('... and opens again at 45 px, the others as they were', widths(9), [9, 45, 80, 90, 180, 349, 4, 18, 80]);
        // Calibri 11: 7 px digits, truncated (9.140625 is 64 px).
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-row-fit.xlsx')).slice());
        grid.SwitchSheet(13); await frame();
        check('Calibri 11: columns of 9.14, 0.5, 2.1, 13, 1, 8.9 are 64, 3, 15, 91, 7, 62 px', widths(6), [64, 3, 15, 91, 7, 62]);
        // The rows Excel fits on opening: no customHeight, whatever ht says.
        grid.SwitchSheet(0); await frame();
        const fit = heights(29);
        check('rows fitted to a 20 pt font, of a custom height, of ht 30 and short text, of an empty 20 pt cell, a number, ht 12, a font smaller than Normal, customHeight without ht',
            [fit[0], fit[3], fit[4], fit[6], fit[7], fit[10], fit[11], fit[13], fit[14]], [35, 20, 20, 35, 35, 35, 20, 35, 20]);
        check('a line break without wrapping is one line; three lines wrapped are three', [fit[9], fit[12]], [20, 60]);
        check('a row style\'s font counts, a cell in a hidden column counts', [fit[16], fit[17], fit[20]], [35, 35, 35]);
        check('rows of Arial 20, Times New Roman 14, Calibri 20 bold', [fit[24], fit[25], fit[26]], [34, 25, 35]);
        check('a pixel more for thickBot, two for thickTop and thickBot', [fit[27], fit[28]], [21, 37]);
        // (Excel: 7 lines. How many ours are follows the widths the text is measured at - Calibri where the
        // browser has it, at Excel's 15 px: 7 - so lines of 20 px, 6 or 7 of them.)
        const wrapped = (height) => height % 20 === 0 && height >= 120 && height <= 140;
        check('wrapped text in a column of 84 px: 7 lines of 20 px, whatever ht says', [wrapped(fit[1]), fit[2] === fit[1]], [true, true]);
        check('... in a column of 64 px: 10 lines', fit[5], 200);
        // A sheet whose default height is its own: fitted only when one of its rows has a height of its own.
        const sheetHeights = async (index) => { grid.SwitchSheet(index); await frame(); return heights(3); };
        const s1 = await sheetHeights(1);
        check('S1: fitted', [s1[0], wrapped(s1[1]), s1[2]], [35, true, 20]);
        check('S2 (customHeight default, no row of a height): not fitted', await sheetHeights(2), [20, 20, 20]);
        check('S3 (... a row of the default ht): not fitted', await sheetHeights(3), [20, 20, 20]);
        const s4 = await sheetHeights(4);
        check('S4 (... a row of customHeight): fitted', [s4[0], wrapped(s4[1])], [35, true]);
        const s6 = await sheetHeights(6);
        check('S6 (... a row of ht 30, not custom): fitted, that row too', [s6[0], wrapped(s6[1]), grid.GetRowHeight(3)], [35, true, 20]);
        check('S7: a default of 20 pt customHeight is 26 px (cut down)', await sheetHeights(7), [26, 26, 26]);
        check('a column style of 20 pt makes every row 35 px', [(await sheetHeights(11))[0], grid.GetRowHeight(2), grid.GetRowHeight(40)], [35, 35, 35]);
        check('every column in a style of Arial 10: rows of 17 px, a Calibri 11 cell\'s 20; some columns: 20', [...(await sheetHeights(9)), grid.GetRowHeight(40), (await sheetHeights(10))[1]], [20, 17, 17, 17, 20]);
        // A row's default style (customFormat) to the last column, over a column's style; a cell's own style is its.
        grid.SwitchSheet(12); await frame();
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const fill = (row, col) => node(row, col) ? getComputedStyle(node(row, col)).backgroundColor : null;
        const yellow = 'rgb(255, 255, 0)', grey = 'rgb(217, 217, 217)', white = 'rgb(255, 255, 255)';
        check('a row style is drawn past its spans, over a column style', [fill(2, 3), fill(2, 7), fill(2, 12), fill(4, 12)], [yellow, yellow, yellow, yellow]);
        check('... a column style where the row has none', fill(3, 7), grey);
        check('... not without customFormat; a cell of its own style keeps it', [fill(6, 1), fill(8, 1), fill(8, 2)], [white, white, yellow]);
        // Shapes of the theme's fills: phClr in the reference's colour.
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-theme-shapes.xlsx')).slice());
        grid.SwitchSheet(0); await frame(); await everyObjectSeen(host);
        const svgs = [...host.querySelectorAll('.spreadsheet-drawing svg')];
        const shapeFill = (n) => {
            const svg = svgs.find((item) => [...item.querySelectorAll('text')].some((text) => text.textContent === n));
            const shape = svg?.querySelector('rect, path, polygon');
            const paint = shape?.getAttribute('fill') ?? null;
            const id = /url\(#([^)]+)\)/.exec(paint ?? '');
            return id ? [...svg.querySelectorAll(`#${id[1]} stop`)].map((stop) => stop.getAttribute('stop-color')) : paint;
        };
        check('fillRef 1 accent1: the solid accent', shapeFill('1'), '#4F81BD');
        const button = shapeFill('2');
        check('fillRef 2 dk1 (the Office 2007 theme\'s gradient in black): light greys, not black', Array.isArray(button) && button.length === 3 && button.every((color) => color !== '#000000' && parseInt(color.slice(1, 3), 16) > 100), true);
        check('fillRef 0: no fill', shapeFill('5'), 'none');
        check('gradFill of accent1 lumMod 75 % to lumMod 40 % lumOff 60 % (Excel\'s accent 1 darker 25 %, lighter 60 %)', shapeFill('6'), ['#376092', '#B9CDE5']);
        check('... the 2007 theme\'s gradient: tints of black mixed in linear light (Excel: C0C0C0 near the bottom, E8E8E8 near the top)', button, ['#BCBCBC', '#D0D0D0', '#EDEDED']);
        check('fillRef 1001 / 1002: the background fills (Excel: D5E0C4)', [shapeFill('9'), shapeFill('10')], ['#4F81BD', '#D5E0C4']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
