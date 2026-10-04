// Browser side of run.mjs, second part: what host applications rely on.
// Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, bytesOf, frame, parts, press, valueAt, formatAt, OF, createSpreadsheet, container } from '/component.js';

const cellNode = (host, row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
const borderOf = (host, row, col, side) => {
    const style = getComputedStyle(cellNode(host, row, col));
    return `${style.getPropertyValue(`border-${side}-width`)} ${style.getPropertyValue(`border-${side}-style`)} ${style.getPropertyValue(`border-${side}-color`)}`;
};
const GRIDLINE = '1px solid rgb(224, 224, 224)';

// Cell borders are drawn: each Excel border style gives a complete CSS border
// (width, line style, colour); ApplyBorder's "outside" / "inside" / "none";
// an imported workbook's borders; the HTML flavour of a copy carries them.
checks['borders'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const sides = (row, col) => ['top', 'right', 'bottom', 'left'].map((side) => borderOf(host, row, col, side));
        grid.SelectRange(1, 1, 3, 3);
        grid.ApplyBorder('all', 'thin', '#000000');
        const black = 'rgb(0, 0, 0)';
        check('all / thin: a 1px solid border on every side of every cell', [sides(1, 1), sides(2, 2), sides(3, 3)].flat().every((v) => v === `1px solid ${black}`), true);
        check('a cell outside the selection keeps its grid lines', sides(5, 5), [GRIDLINE, GRIDLINE, GRIDLINE, GRIDLINE]);
        const kinds = { thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', hair: '1px dotted',
                        mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed' };
        grid.SelectCell(6, 1);
        for (const [style, css] of Object.entries(kinds)) {
            grid.ApplyBorder('top', style, '#ff0000');
            check(`${style}: computed border-top`, borderOf(host, 6, 1, 'top'), `${css} rgb(255, 0, 0)`);
        }
        check('the other sides of that cell are untouched', sides(6, 1).slice(1), [GRIDLINE, GRIDLINE, GRIDLINE]);
        grid.ApplyBorder('top', 'none', '#000000');
        check('a style of none clears the side', borderOf(host, 6, 1, 'top'), GRIDLINE);

        // none: one step, the whole selection.
        grid.SelectRange(1, 1, 3, 3);
        grid.ApplyBorder('none', 'thin', '#000000');
        check('none: every border of the selection cleared', [sides(1, 1), sides(2, 2), sides(3, 3)].flat().every((v) => v === GRIDLINE), true);
        grid.Undo();
        check('none: one undo step brings them all back', [sides(1, 1), sides(3, 3)].flat().every((v) => v === `1px solid ${black}`), true);
        grid.Redo();
        // outside: the outline only.
        grid.ApplyBorder('outside', 'medium', '#0000ff');
        const blue = '2px solid rgb(0, 0, 255)';
        check('outside: the corner cell has its two outer sides', sides(1, 1), [blue, GRIDLINE, GRIDLINE, blue]);
        check('outside: an edge cell has one', sides(1, 2), [blue, GRIDLINE, GRIDLINE, GRIDLINE]);
        check('outside: the middle cell has none', sides(2, 2), [GRIDLINE, GRIDLINE, GRIDLINE, GRIDLINE]);
        check('outside: the far corner', sides(3, 3), [GRIDLINE, blue, blue, GRIDLINE]);
        grid.Undo();
        check('outside: one undo step', [sides(1, 1), sides(3, 3)].flat().every((v) => v === GRIDLINE), true);
        grid.ApplyBorder('inside', 'thin', '#000000');
        const line = `1px solid ${black}`;
        check('inside: the corner cell has its two inner sides', sides(1, 1), [GRIDLINE, line, line, GRIDLINE]);
        check('inside: the middle cell has all four', sides(2, 2), [line, line, line, line]);
        grid.Undo();

        // The HTML flavour of a copy.
        grid.SelectRange(1, 1, 1, 2);
        grid.ApplyBorder('bottom', 'double', '#00aa00');
        grid.SetCell(1, 1, 'x');
        const html = grid.GetSelectionHtml();
        check('GetSelectionHtml carries the border', /border-bottom:3px double #00aa00/.test(html), true);
        const probe = document.createElement('div'); probe.innerHTML = html; document.body.appendChild(probe);
        check('and a browser reads it as a border', getComputedStyle(probe.querySelector('td')).borderBottomStyle, 'double');
        probe.remove();

        // The portable painter draws them too (a thick red box on a canvas).
        grid.SelectRange(9, 1, 10, 2);
        grid.ApplyBorder('outside', 'thick', '#ff0000');
        grid.SelectCell(0, 0);   // the selection outline is drawn over the cells
        const canvas = document.createElement('canvas');
        grid.RenderPortableCanvas(canvas);
        const { columnLeft, rowTop } = await import('/spreadsheet/Layout.js');
        const model = grid.GetModel();
        const pixel = canvas.getContext('2d').getImageData(Math.round(50 + (columnLeft(model, 1) + columnLeft(model, 2)) / 2), Math.round(30 + rowTop(model, 9)), 1, 1).data;
        check('the portable painter draws the border', [pixel[0] > 200, pixel[1] < 80, pixel[2] < 80], [true, true, true]);

        // An imported workbook: every bordered cell of the sheet on display has a drawn border.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        await frame();
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        let bordered = 0, drawn = 0;
        for (const node of host.querySelectorAll('.spreadsheet-cell')) {
            if (node.style.display === 'none') continue;
            const row = Number(node.getAttribute('data-row')), col = Number(node.getAttribute('data-col'));
            const format = getCell(new CellAddress(row, col), grid.GetModel())?.Format;
            for (const [side, border] of [['top', format?.BorderTop], ['left', format?.BorderLeft]]) {
                if (!border) continue;
                bordered++;
                const style = getComputedStyle(node);
                if (style.getPropertyValue(`border-${side}-style`) !== 'none' && style.getPropertyValue(`border-${side}-color`) !== 'rgb(224, 224, 224)' && parseFloat(style.getPropertyValue(`border-${side}-width`)) >= 1) drawn++;
            }
        }
        check('imported borders are drawn (bordered sides in view > 0, all drawn)', [bordered > 0, drawn], [true, bordered]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// InsertRow / DeleteRow / InsertColumn / DeleteColumn take a count; InsertRows / DeleteRows /
// InsertColumns / DeleteColumns act on the selection when called without arguments. Each is
// one change, one undo step, and formulas are rewritten once.
checks['row-column-counts'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        const formula = (row, col) => { const v = getCell(new CellAddress(row, col), grid.GetModel())?.Value; return v && v.fields ? String(v.fields[0]) : null; };
        const fill = () => { for (let row = 0; row < 10; row++) { grid.SetCell(row, 0, String(row + 1)); grid.SetCell(row, 1, String((row + 1) * 10)); } grid.SetCell(12, 0, '=SUM(A1:A10)'); grid.SetCell(0, 5, '=B10+A9'); };
        fill();
        const column = async (col, n) => { const out = []; for (let row = 0; row < n; row++) out.push(await valueAt(grid, row, col)); return out; };
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        const rows0 = grid.GetModel().RowCount, cols0 = grid.GetModel().ColCount;
        grid.InsertRow(2, 3);
        check('InsertRow(2, 3): three rows, one change', [await column(0, 7), changes.length, grid.GetModel().RowCount - rows0], [['1', '2', '', '', '', '3', '4'], 1, 3]);
        check('formulas follow, rewritten once', [formula(15, 0), await valueAt(grid, 15, 0), formula(0, 5)], ['=SUM(A1:A13)', '55', '=B13+A12']);
        grid.Undo();
        check('one undo step', [await column(0, 4), formula(12, 0), grid.GetModel().RowCount - rows0, grid.CanUndo()], [['1', '2', '3', '4'], '=SUM(A1:A10)', 0, false]);
        grid.Redo(); grid.Undo();
        changes.length = 0;
        grid.DeleteRow(1, 4);
        check('DeleteRow(1, 4)', [await column(0, 6), changes.length, formula(8, 0), await valueAt(grid, 8, 0)], [['1', '6', '7', '8', '9', '10'], 1, '=SUM(A1:A6)', '41']);
        grid.Undo();
        check('undone in one step', [await column(0, 3), await valueAt(grid, 12, 0)], [['1', '2', '3'], '55']);
        grid.DeleteRow(0);
        check('DeleteRow(0): one row as before', [await column(0, 2), await valueAt(grid, 11, 0)], [['2', '3'], '54']);
        grid.Undo();
        grid.InsertColumn(1, 2);
        check('InsertColumn(1, 2)', [await valueAt(grid, 0, 1), await valueAt(grid, 0, 3), formula(0, 7), grid.GetModel().ColCount - cols0], ['', '10', '=D10+A9', 2]);
        grid.Undo();
        grid.DeleteColumn(1, 1);
        check('DeleteColumn(1, 1)', [await valueAt(grid, 0, 1), grid.GetModel().ColCount - cols0], ['', -1]);
        grid.Undo();
        // On the selection.
        grid.SelectRange(3, 0, 5, 1);
        changes.length = 0;
        grid.InsertRows();
        check('InsertRows(): as many rows as selected, above them', [await column(0, 8), changes.length], [['1', '2', '3', '', '', '', '4', '5'], 1]);
        grid.Undo();
        grid.SelectRange(3, 0, 5, 1);
        grid.DeleteRows();
        check('DeleteRows(): the selected rows', [await column(0, 5), await valueAt(grid, 9, 0)], [['1', '2', '3', '7', '8'], '40']);
        grid.Undo();
        check('and back in one step', await valueAt(grid, 12, 0), '55');
        grid.SelectRange(0, 1, 2, 2);
        grid.InsertColumns();
        check('InsertColumns(): two columns before the selection', [await valueAt(grid, 0, 1), await valueAt(grid, 0, 2), await valueAt(grid, 0, 3)], ['', '', '10']);
        grid.Undo();
        grid.SelectRange(0, 1, 2, 1);
        grid.DeleteColumns();
        check('DeleteColumns(): the selected column', [await valueAt(grid, 0, 1), grid.GetModel().ColCount - cols0], ['', -1]);
        grid.Undo();
        grid.InsertRows(4, 2); grid.DeleteRows(4, 2);
        check('InsertRows(at, count) / DeleteRows(at, count)', [await column(0, 6), await valueAt(grid, 12, 0)], [['1', '2', '3', '4', '5', '6'], '55']);
        grid.InsertColumns(0, 2); grid.DeleteColumns(0, 2);
        check('InsertColumns(at, count) / DeleteColumns(at, count)', [await valueAt(grid, 0, 0), await valueAt(grid, 0, 1)], ['1', '10']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// GetRevision / MarkSaved / IsModified: undo back to the saved state is clean again; a save
// that finishes late marks the revision it wrote; OnChange carries Revision and Modified.
checks['revision'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        check('a new grid is unmodified', [grid.IsModified(), grid.GetRevision()], [false, 0]);
        const changes = [];
        grid.OnChange((change) => changes.push([change.Revision, change.Modified]));
        grid.EnterText(0, 0, 'one');
        const r1 = grid.GetRevision();
        check('an edit is a new revision, reported with the change', [grid.IsModified(), r1 !== 0, changes], [true, true, [[r1, true]]]);
        grid.SelectCell(4, 4); grid.Render();
        check('a selection move is not', grid.GetRevision(), r1);
        grid.Undo();
        check('undo restores the revision', [grid.GetRevision(), grid.IsModified(), changes[changes.length - 1]], [0, false, [0, false]]);
        grid.Redo();
        check('redo too', [grid.GetRevision(), grid.IsModified()], [r1, true]);
        grid.MarkSaved();
        check('MarkSaved: unmodified', grid.IsModified(), false);
        grid.EnterText(1, 0, 'two'); grid.SelectRange(0, 0, 1, 0); grid.ApplyFormat('bold'); grid.InsertRow(0, 2);
        const r4 = grid.GetRevision();
        check('three more steps: modified', [grid.IsModified(), new Set([r1, r4]).size], [true, 2]);
        grid.Undo(); grid.Undo();
        check('part of the way back: still modified', grid.IsModified(), true);
        grid.Undo();
        check('undone back to the saved state: unmodified', [grid.IsModified(), grid.GetRevision()], [false, r1]);
        grid.Undo();
        check('undone past it: modified', grid.IsModified(), true);
        grid.Redo();
        check('and redone to it: unmodified', grid.IsModified(), false);
        grid.Redo(); grid.Redo(); grid.Redo();
        check('redone to the end: the revision it had', grid.GetRevision(), r4);
        // A new edit after an undo is a revision of its own (not the one that was undone).
        grid.Undo();
        const undone = grid.GetRevision();
        grid.EnterText(5, 5, 'x');
        check('a new edit after an undo is a new revision', [grid.GetRevision() !== r4, grid.GetRevision() !== undone], [true, true]);
        // The save protocol: the revision that was written is the one marked.
        const written = grid.GetRevision();
        const bytes = await grid.ExportXlsx();
        grid.EnterText(6, 6, 'typed while saving');
        grid.MarkSaved(written);
        check('a save that finishes after more typing leaves the workbook modified', [grid.IsModified(), bytes.length > 0], [true, true]);
        grid.Undo();
        check('and undoing that typing gives the saved state', grid.IsModified(), false);
        // A change without an undo step cannot be walked back to a known revision.
        grid.SetCell(8, 8, 'host write');
        const hostRevision = grid.GetRevision();
        grid.Undo();
        check('undo after a host write is a new revision, not a saved one', [grid.IsModified(), grid.GetRevision() !== hostRevision, grid.GetRevision() !== written], [true, true, true]);
        // Opening a workbook: saved.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        check('an opened workbook is unmodified', grid.IsModified(), false);
        const opened = grid.GetRevision();
        grid.SwitchSheet(1); grid.SwitchSheet(0);
        check('switching sheets changes nothing', [grid.IsModified(), grid.GetRevision()], [false, opened]);
        grid.EnterText(0, 0, 'edit'); grid.Undo();
        check('edit and undo in an opened workbook: unmodified', [grid.IsModified(), grid.GetRevision()], [false, opened]);
        // Unmounted: the revision still follows.
        const bare = createSpreadsheet(document.createElement('div'));
        bare.EnterText(0, 0, 'a'); const b1 = bare.GetRevision(); bare.EnterText(0, 1, 'b');
        bare.Undo();
        check('not mounted: edits and undo move the revision', [bare.IsModified(), bare.GetRevision()], [true, b1]);
        bare.Undo();
        check('not mounted: all undone is unmodified', bare.IsModified(), false);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// ExportXlsx keeps the document properties, printer settings (with the pageSetup r:id), the
// calculation chain (only while the formula cells are where they were) and the x14 extensions
// of sheets that were not restructured; what belonged to a removed sheet is not left behind.
// (Every export here also goes through the package-integrity check of component.js.)
checks['export-package-parts'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const partsIn = async (bytes) => { const zip = await window.JSZip.loadAsync(bytes); const out = {}; for (const [path, file] of Object.entries(zip.files)) if (!file.dir && !/^xl\/(media|pivotCache\/pivotCacheRecords|worksheets\/sheet8)/.test(path)) out[path] = await file.async('string'); return out; };
    try {
        const { packageProblems } = await import('/component.js');
        const original = await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx');
        const before = await partsIn(original);
        check('the fixture itself is whole', await packageProblems(original), []);
        await grid.ImportXlsx(original);
        grid.EnterText(40, 0, 'edited');
        let after = await partsIn(await grid.ExportXlsx());
        for (const part of ['docProps/core.xml', 'docProps/app.xml', 'xl/printerSettings/printerSettings1.bin', 'xl/calcChain.xml'])
            check(`a value edit: ${part} is carried over`, after[part] === before[part] && after[part] !== undefined, true);
        check('the package relationships name the properties', ['docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml'].every((target) => after['_rels/.rels'].includes(target)), true);
        const dataIndex = grid.GetSheetNames().indexOf('Data') + 1;
        const relId = (/<pageSetup[^>]*\sr:id="([^"]*)"/.exec(await (await window.JSZip.loadAsync(await grid.ExportXlsx())).files[`xl/worksheets/sheet${dataIndex}.xml`].async('string')) ?? [])[1];
        check('pageSetup names the printer settings by a relationship of the sheet', new RegExp(`Id="${relId}"[^>]*printerSettings1\\.bin`).test(after[`xl/worksheets/_rels/sheet${dataIndex}.xml.rels`]), true);
        check('x14 conditional formats of an untouched sheet are kept', [(after['xl/worksheets/sheet1.xml'].match(/<x14:cfRule\b/g) ?? []).length > 0, (after['xl/worksheets/sheet1.xml'].match(/<x14:cfRule\b/g) ?? []).length], [true, (before['xl/worksheets/sheet1.xml'].match(/<x14:cfRule\b/g) ?? []).length]);
        // A formula added: Excel rebuilds the chain; its relationship and content type go with it.
        grid.EnterText(41, 0, '=1+1');
        after = await partsIn(await grid.ExportXlsx());
        check('a formula added: no calcChain part, relationship or content type', [after['xl/calcChain.xml'] === undefined, /calcChain/.test(after['xl/_rels/workbook.xml.rels']), /calcChain/.test(after['[Content_Types].xml'])], [true, false, false]);
        check('the properties and printer settings stay', [after['docProps/core.xml'] === before['docProps/core.xml'], after['xl/printerSettings/printerSettings1.bin'] !== undefined], [true, true]);
        grid.Undo();
        after = await partsIn(await grid.ExportXlsx());
        check('undone: the chain is kept again', after['xl/calcChain.xml'] === before['xl/calcChain.xml'], true);
        // Rows inserted in the first sheet: its x14 formats would be stale.
        grid.InsertRows(2, 2);
        after = await partsIn(await grid.ExportXlsx());
        check('rows inserted: the x14 formats of that sheet are not written stale', /x14:conditionalFormattings/.test(after['xl/worksheets/sheet1.xml']), false);
        grid.Undo();
        // The Data sheet removed by the host.
        const sheets = grid.GetSheets(), names = grid.GetSheetNames();
        const keep = names.map((_, i) => i).filter((i) => names[i] !== 'Data');
        grid.InstallWorkbook(keep.map((i) => sheets[i]), keep.map((i) => names[i]), 0, undefined);
        after = await partsIn(await grid.ExportXlsx());
        check('a removed sheet: its printer settings are not left in the package', [after['xl/printerSettings/printerSettings1.bin'] === undefined, /printerSettings/.test(after['[Content_Types].xml'].replace(/<Default[^>]*>/g, ''))], [true, false]);
        check('the document properties stay', after['docProps/app.xml'] === before['docProps/app.xml'], true);
        // Custom properties.
        const custom = await bytesOf('tests/real-world/milk-street-office.xlsx');
        await grid.ImportXlsx(custom);
        const customBefore = await partsIn(custom);
        after = await partsIn(await grid.ExportXlsx());
        check('custom properties and several printer settings', ['docProps/custom.xml', 'docProps/core.xml', 'docProps/app.xml', 'xl/printerSettings/printerSettings1.bin', 'xl/printerSettings/printerSettings2.bin', 'xl/calcChain.xml'].map((part) => after[part] === customBefore[part] && after[part] !== undefined), [true, true, true, true, true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Text wider than its cell runs over the empty cells beside it, up to the first filled one
// (measured from the DOM); numbers never do - one that does not fit shows ####.
checks['text-overflow'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        // How far a cell's text is visible past its own box: [left, right] in px, from its clip.
        const let_out = (row, col) => {
            const node = cellNode(host, row, col), style = getComputedStyle(node);
            // (A browser gives the shortest form: "inset(0px -200px)" when left and right are the same.)
            const inset = (/inset\(([^)]*)\)/.exec(style.clipPath)?.[1] ?? '').split(' ').filter(Boolean).map((v) => -parseFloat(v));
            return style.overflow === 'visible' && inset.length >= 2 ? [inset[3] ?? inset[1], inset[1]] : [0, 0];
        };
        // The ink of a cell's text, relative to the cell's box: [left overhang, right overhang].
        const ink = (row, col) => {
            const node = cellNode(host, row, col), range = document.createRange();
            range.selectNodeContents(node);
            const text = [...range.getClientRects()].reduce((box, r) => ({ left: Math.min(box.left, r.left), right: Math.max(box.right, r.right) }), { left: Infinity, right: -Infinity });
            const box = node.getBoundingClientRect();
            return [Math.max(0, Math.round(box.left - text.left)), Math.max(0, Math.round(text.right - box.right))];
        };
        // What is drawn at a point of the grid (the topmost cell there).
        const topCellAt = (row, col) => {
            const box = cellNode(host, row, col).getBoundingClientRect();
            const node = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2).closest('.spreadsheet-cell');
            return [Number(node.getAttribute('data-row')), Number(node.getAttribute('data-col'))];
        };
        const title = 'A title that is much longer than one column of the sheet';
        grid.SetCell(1, 1, title);
        let over = ink(1, 1), allowed = let_out(1, 1);
        check('long text: the cell shows all of it', cellNode(host, 1, 1).textContent, title);
        check('its ink runs past the cell to the right, not to the left', [over[0], over[1] > 100], [0, true]);
        check('and the cell lets it out that far (whole columns)', [allowed[0], allowed[1] >= over[1], allowed[1] % 100], [0, true, 0]);
        check('the text is drawn over the empty neighbour', topCellAt(1, 2), [1, 1]);
        check('no ellipsis', getComputedStyle(cellNode(host, 1, 1)).textOverflow, 'clip');
        check('the grid lines it crosses are covered', cellNode(host, 1, 1).querySelectorAll('.spreadsheet-overflow-gap').length, allowed[1] / 100);
        // A filled cell stops it.
        grid.SetCell(1, 3, 'stop');
        check('a filled cell two columns on: let out over one column only', let_out(1, 1), [0, 100]);
        check('that cell is drawn itself', topCellAt(1, 3), [1, 3]);
        grid.SetCell(1, 2, 'x');
        check('a filled neighbour: the text stays in its cell', [let_out(1, 1), getComputedStyle(cellNode(host, 1, 1)).overflow], [[0, 0], 'hidden']);
        grid.SetCell(1, 2, ''); grid.SetCell(1, 3, '');
        check('cleared again: let out again', let_out(1, 1)[1] >= over[1], true);
        // Right-aligned and centred text.
        grid.SetCell(3, 6, title); grid.SelectCell(3, 6); grid.ApplyFormat('alignRight');
        over = ink(3, 6); allowed = let_out(3, 6);
        check('right-aligned: runs to the left', [over[0] > 100, over[1], allowed[0] >= over[0], allowed[1]], [true, 0, true, 0]);
        grid.SetCell(5, 5, title); grid.SelectCell(5, 5); grid.ApplyFormat('alignCenter');
        over = ink(5, 5); allowed = let_out(5, 5);
        check('centred: runs both ways, evenly', [over[0] > 50, over[1] > 50, Math.abs(over[0] - over[1]) <= 2, allowed[0] >= over[0], allowed[1] >= over[1]], [true, true, true, true, true]);
        // Numbers.
        grid.SetCell(7, 1, '123456789012345678');
        grid.SetCell(8, 1, '12.5');
        const { setColumnWidth } = await import('/spreadsheet/Model.js');
        grid.SetWorkbook([setColumnWidth(1, 40, grid.GetModel())], ['Sheet1'], 0);
        await frame();
        check('a number never overflows', [let_out(7, 1), getComputedStyle(cellNode(host, 7, 1)).overflow], [[0, 0], 'hidden']);
        check('a number too wide for its cell shows ####', [/^#+$/.test(cellNode(host, 7, 1).textContent), cellNode(host, 8, 1).textContent], [true, '12.5']);
        check('and the hashes fit the cell', ink(7, 1), [0, 0]);
        check('text in the narrow column still runs over', let_out(1, 1)[1] > 100, true);
        // Wrapped text and merged cells stay where they are.
        const { getCell, CellAddress, setCell } = await import('/spreadsheet/Model.js');
        grid.SetCell(10, 1, title);
        grid.MergeCells(10, 1, 10, 2);
        check('a merged cell keeps its text inside', let_out(10, 1), [0, 0]);
        // The workbook the hosts reported: B2 of the first sheet shows its long title.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        await frame();
        const b2 = cellNode(host, 1, 1);
        over = ink(1, 1); allowed = let_out(1, 1);
        check('milk-street-office B2: the whole title is shown', [b2.textContent.length > 10, over[1] > 0, allowed[1] >= over[1], getComputedStyle(b2).textOverflow], [true, true, true, 'clip']);
        const right = b2.getBoundingClientRect().right + Math.min(over[1], 60) - 5, middle = b2.getBoundingClientRect().top + b2.getBoundingClientRect().height / 2;
        check('and drawn over the cells beside it', document.elementFromPoint(right, middle).closest('.spreadsheet-cell') === b2, true);
        // Scrolled so that column B is out of view: its text still shows in the columns it runs over.
        const { scroller } = parts(host);
        const { columnLeft } = await import('/spreadsheet/Layout.js');
        scroller.scrollLeft = columnLeft(grid.GetModel(), 2) + 30; scroller.dispatchEvent(new Event('scroll'));
        await frame();
        check('scrolled past its column: the cell is still drawn', [cellNode(host, 1, 1) !== null && cellNode(host, 1, 1).style.display !== 'none', let_out(1, 1)[1] >= over[1]], [true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Dispose removes every listener the grid registered - also those on its own (now detached)
// nodes and on the nodes of its dialogs - and a disposed grid is collected.
checks['dispose-releases'] = async () => {
    const { lines, check } = makeLog();
    // Every listener registered anywhere while this runs and not removed since.
    const live = new Map();   // target -> Map(type -> Set(handler))
    const add = EventTarget.prototype.addEventListener, remove = EventTarget.prototype.removeEventListener;
    EventTarget.prototype.addEventListener = function (type, handler, options) {
        if (!live.has(this)) live.set(this, new Map());
        const types = live.get(this);
        if (!types.has(type)) types.set(type, new Set());
        types.get(type).add(handler);
        return add.call(this, type, handler, options);
    };
    EventTarget.prototype.removeEventListener = function (type, handler, options) {
        live.get(this)?.get(type)?.delete(handler);
        return remove.call(this, type, handler, options);
    };
    const count = (where) => [...live].reduce((n, [target, types]) => n + (where(target) ? [...types.values()].reduce((m, set) => m + set.size, 0) : 0), 0);
    const describe = () => [...live].flatMap(([target, types]) => [...types].filter(([, set]) => set.size > 0).map(([type, set]) => `${target === document ? 'document' : target === window ? 'window' : (target.className || target.nodeName)}:${type}x${set.size}`));
    try {
        const run = async () => {
            const host = container('position:absolute;left:0;top:0;width:1100px;height:600px');
            const grid = createSpreadsheet(host);
            grid.Mount();
            await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
            grid.OnChange(() => {}); grid.OnSelectionChange(() => {}); grid.OnEditStateChange(() => {}); grid.OnValidationNotice(() => {});
            grid.SelectRange(2, 1, 6, 4); grid.EnterText(3, 1, 'x'); grid.StartEdit('abc'); grid.CancelEdit();
            // Dialogs and a dropdown: one closed again, the others left open.
            grid.ShowFindReplaceDialog();
            [...document.querySelectorAll('button')].find((button) => button.textContent === 'Close')?.click();
            grid.ShowFindReplaceDialog();
            grid.ShowConditionalFormattingDialog();
            host.querySelector('.col-filter-btn')?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            await new Promise((resolve) => setTimeout(resolve, 150));
            const registered = count(() => true);
            const onGrid = count((target) => target instanceof Node && host.contains(target));
            const nodes = [...host.querySelectorAll('*')];
            const bodyChildren = document.body.children.length;
            grid.Dispose();
            return { registered, onGrid, weak: new WeakRef(grid), model: new WeakRef(grid.GetModel()), hostNode: host, nodes: nodes.length };
        };
        const result = await run();
        check('the grid had listeners on the document and on its own nodes', [result.registered > 10, result.onGrid > 5], [true, true]);
        check('after Dispose none is left anywhere', describe(), []);
        check('its nodes are out of the page', [result.hostNode.querySelector('.spreadsheet-container'), document.querySelectorAll('.spreadsheet-container').length], [null, 0]);
        result.hostNode.remove();
        delete result.hostNode;
        // A disposed grid nobody holds is collected, with its workbook.
        let collected = false;
        for (let attempt = 0; attempt < 5 && !collected; attempt++) {
            await new Promise((resolve) => setTimeout(resolve, 50));
            await window.collectGarbage();
            collected = result.weak.deref() === undefined && result.model.deref() === undefined;
        }
        check('a disposed grid and its workbook are collected', [result.weak.deref() === undefined, result.model.deref() === undefined], [true, true]);
        // A disposed grid can be mounted again and works.
        const host = container('position:absolute;left:0;top:0;width:900px;height:400px');
        const grid = createSpreadsheet(host);
        grid.Mount(); grid.SetCell(0, 0, 'kept'); grid.Dispose();
        check('disposed: no listeners', describe(), []);
        grid.Mount();
        check('mounted again: the workbook is still there and drawn', [cellNode(host, 0, 0)?.textContent, count(() => true) > 5], ['kept', true]);
        grid.Dispose(); host.remove();
        check('and disposed again: no listeners', describe(), []);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    EventTarget.prototype.addEventListener = add; EventTarget.prototype.removeEventListener = remove;
    return lines;
};
