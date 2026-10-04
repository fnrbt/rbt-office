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
