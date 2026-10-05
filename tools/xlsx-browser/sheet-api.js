// Browser side of run.mjs, third part: the host API of the Spreadsheet component - what a host
// application builds its toolbars, dialogs and panels on. One check per group of members
// ("api-..."). Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, bytesOf, frame, OF, createSpreadsheet, container } from '/component.js';

// The message of what a call throws, or null when it returns.
export const thrown = (run) => { try { run(); return null; } catch (e) { return String(e && e.message || e); } };
// Export the grid's workbook and open it in a second grid: { grid, close }.
export const reopened = async (grid) => {
    const bytes = await grid.ExportXlsx();
    const again = mountGrid();
    await again.grid.ImportXlsx(bytes);
    return again;
};
const pick = (object, keys) => Object.fromEntries(keys.map((key) => [key, object[key]]));

// Reading cells: GetCell, GetRange, GetSelectionFormat.
checks['api-read'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        grid.SetCell(0, 0, '12.5'); grid.SetCell(1, 0, '=A1*2'); grid.SetCell(2, 0, 'hello'); grid.SetCell(3, 0, '=1/0');
        grid.SetCell(4, 0, 'TRUE'); grid.SetCell(5, 0, '2024-03-05'); grid.SetCell(6, 0, '45%'); grid.SetCell(7, 0, '1234.5');
        grid.SelectCell(7, 0); grid.ApplyNumberFormat('#,##0.00');

        // ---- GetCell ----
        const a1 = grid.GetCell(0, 0);
        check('a number', pick(a1, ['row', 'col', 'address', 'value', 'valueType', 'text', 'formula']), { row: 0, col: 0, address: 'A1', value: 12.5, valueType: 'number', text: '12.5', formula: null });
        check('a formula gives its result and its text', pick(grid.GetCell(1, 0), ['value', 'valueType', 'text', 'formula']), { value: 25, valueType: 'number', text: '25', formula: '=A1*2' });
        check('text', pick(grid.GetCell(2, 0), ['value', 'valueType', 'text', 'formula']), { value: 'hello', valueType: 'string', text: 'hello', formula: null });
        check('an error', pick(grid.GetCell(3, 0), ['value', 'valueType', 'text', 'formula']), { value: '#DIV/0!', valueType: 'error', text: '#DIV/0!', formula: '=1/0' });
        check('a boolean', pick(grid.GetCell(4, 0), ['value', 'valueType', 'text']), { value: true, valueType: 'boolean', text: 'TRUE' });
        const date = grid.GetCell(5, 0);
        check('a date is its serial number, shown through its format', [date.value, date.valueType, date.text, typeof date.format.numberFormat], [45356, 'number', host.querySelector('.spreadsheet-cell[data-row="5"][data-col="0"]').textContent, 'string']);
        check('a percentage', pick(grid.GetCell(6, 0), ['value', 'text']), { value: 0.45, text: '45%' });
        check('a number format shapes the text', [grid.GetCell(7, 0).text, grid.GetCell(7, 0).format.numberFormat], ['1,234.50', '#,##0.00']);
        check('the text is what the grid shows', [0, 1, 2, 3, 4, 6, 7].map((row) => grid.GetCell(row, 0).text), [0, 1, 2, 3, 4, 6, 7].map((row) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="0"]`).textContent));
        const empty = grid.GetCell(20, 5);
        check('an empty cell', pick(empty, ['address', 'value', 'valueType', 'text', 'formula', 'isMerged', 'mergeRange', 'comment', 'hyperlink', 'validation']),
            { address: 'F21', value: null, valueType: 'empty', text: '', formula: null, isMerged: false, mergeRange: null, comment: null, hyperlink: null, validation: null });
        check('its format is the default, every key present', empty.format, {
            bold: false, italic: false, underline: false, doubleUnderline: false, strikethrough: false, fontFamily: null, fontSize: null, textColor: null, fill: null,
            horizontalAlign: null, verticalAlign: null, wrap: false, indent: 0, shrinkToFit: false, textRotation: 0, numberFormat: null,
            borders: { top: null, bottom: null, left: null, right: null, diagonal: null, diagonalUp: false, diagonalDown: false } });

        // Formats.
        grid.SelectCell(2, 0);
        grid.ApplyFormat('bold'); grid.ApplyFormat('italic'); grid.ApplyFormat('underline'); grid.ApplyFormat('alignCenter');
        grid.ApplyTextColor('#ff0000'); grid.ApplyBgColor('#00FF00'); grid.ApplyBorder('bottom', 'medium', '#0000ff');
        const formatted = grid.GetCell(2, 0).format;
        check('the format of a formatted cell', [formatted.bold, formatted.italic, formatted.underline, formatted.doubleUnderline, formatted.textColor.toLowerCase(), formatted.fill.toLowerCase(), formatted.horizontalAlign, formatted.borders.bottom && [formatted.borders.bottom.style, formatted.borders.bottom.color.toLowerCase()], formatted.borders.top],
            [true, true, true, false, '#ff0000', '#00ff00', 'center', ['medium', '#0000ff'], null]);

        // Merged cells, comments.
        grid.SetCell(10, 1, 'merged'); grid.MergeCells(10, 1, 11, 3);
        check('the first cell of a merged cell', pick(grid.GetCell(10, 1), ['value', 'isMerged', 'mergeRange']), { value: 'merged', isMerged: true, mergeRange: 'B11:D12' });
        check('another cell of it', pick(grid.GetCell(11, 3), ['value', 'isMerged', 'mergeRange']), { value: null, isMerged: true, mergeRange: 'B11:D12' });
        check('the cell beside it', pick(grid.GetCell(11, 4), ['isMerged', 'mergeRange']), { isMerged: false, mergeRange: null });
        grid.SetComment(2, 0, 'look here', 'Ada');
        check('a comment', grid.GetCell(2, 0).comment, { text: 'look here', author: 'Ada' });
        grid.SetComment(1, 0, 'anonymous', undefined);
        check('a comment without an author', grid.GetCell(1, 0).comment, { text: 'anonymous', author: null });

        // Invalid input.
        check('a row outside the sheet throws', thrown(() => grid.GetCell(-1, 0)), 'GetCell: row -1 is outside the sheet (0 to 999)');
        check('a column outside the sheet throws', /^GetCell: col 26 is outside the sheet/.test(thrown(() => grid.GetCell(0, 26))), true);
        check('a row that is not a whole number throws', [thrown(() => grid.GetCell(1.5, 0)), thrown(() => grid.GetCell('1', 0)), /^GetCell: row must be a whole number/.test(thrown(() => grid.GetCell(undefined, 0)))],
            ['GetCell: row must be a whole number, got 1.5', 'GetCell: row must be a whole number, got "1"', true]);

        // ---- GetRange ----
        const block = grid.GetRange('A1:B3');
        check('a block: rows of cells', [block.length, block.map((row) => row.length), block.map((row) => row.map((cell) => cell.address)), block.map((row) => row[0].value)], [3, [2, 2, 2], [['A1', 'B1'], ['A2', 'B2'], ['A3', 'B3']], [12.5, 25, 'hello']]);
        check('a cell of a range is the cell GetCell gives', block[1][0], grid.GetCell(1, 0));
        check('one cell', grid.GetRange('$A$3').map((row) => row.map((cell) => cell.text)), [['hello']]);
        check('corners in any order', grid.GetRange('B3:A1').map((row) => row.map((cell) => cell.address)), [['A1', 'B1'], ['A2', 'B2'], ['A3', 'B3']]);
        const column = grid.GetRange('A:A');
        check('a whole column is cut down to the used rows', [column.length, column[0].length, column[7][0].text, column[11][0].address], [12, 1, '1,234.50', 'A12']);
        const rows = grid.GetRange('11:12');
        check('whole rows are cut down to the used columns', [rows.length, rows[0].length, rows[0][1].value, rows[1][3].mergeRange], [2, 4, 'merged', 'B11:D12']);
        check('a block past the last column is cut at the sheet', grid.GetRange('Y1:AB2').map((row) => row.map((cell) => cell.address)), [['Y1', 'Z1'], ['Y2', 'Z2']]);
        check('a block outside the sheet is empty', grid.GetRange('AB1:AC2'), []);
        check('the active sheet by name', grid.GetRange('Sheet1!A1')[0][0].value, 12.5);
        check('not a range throws', [thrown(() => grid.GetRange('A:5')), thrown(() => grid.GetRange('')), /^GetRange: the range must be a string/.test(thrown(() => grid.GetRange(7))), thrown(() => grid.GetRange('Nowhere!A1'))],
            ['GetRange: "A:5" is not a range (expected A1 notation such as "B2", "B2:D5", "A:C" or "3:5")', 'GetRange: "" is not a range (expected A1 notation such as "B2", "B2:D5", "A:C" or "3:5")', true, 'GetRange: there is no sheet "Nowhere"']);

        // ---- GetSelectionFormat ----
        grid.SelectCell(2, 0);
        const one = grid.GetSelectionFormat();
        check('one cell: its format, nothing mixed', [one.bold, one.italic, one.underline, one.horizontalAlign, one.fill.toLowerCase(), one.borders.bottom.style, one.numberFormat, JSON.stringify(one).includes('mixed')], [true, true, true, 'center', '#00ff00', 'medium', null, false]);
        check('the same keys as a format', Object.keys(one), Object.keys(empty.format));
        check('asked again for the same selection: the same answer', grid.GetSelectionFormat() === one, true);
        grid.SelectRange(1, 0, 2, 0);
        const two = grid.GetSelectionFormat();
        check('two cells that differ: mixed where they do', pick(two, ['bold', 'italic', 'underline', 'strikethrough', 'fontFamily', 'fontSize', 'textColor', 'fill', 'horizontalAlign', 'verticalAlign', 'wrap', 'numberFormat']),
            { bold: 'mixed', italic: 'mixed', underline: 'mixed', strikethrough: false, fontFamily: null, fontSize: null, textColor: 'mixed', fill: 'mixed', horizontalAlign: 'mixed', verticalAlign: null, wrap: false, numberFormat: null });
        check('borders are summarised side by side', two.borders, { top: null, bottom: 'mixed', left: null, right: null, diagonal: null });
        grid.ApplyFormat('bold');   // the active cell (A2) is not bold: bold for both
        const both = grid.GetSelectionFormat();
        check('after a format command the answer is new', [both === two, both.bold, both.italic], [false, true, 'mixed']);
        grid.Undo();
        check('and after its undo', grid.GetSelectionFormat().bold, 'mixed');
        grid.SelectRange(6, 0, 7, 0);
        check('number formats that differ', grid.GetSelectionFormat().numberFormat, 'mixed');
        grid.SelectRange(20, 2, 30, 6);
        check('empty cells: the default format', grid.GetSelectionFormat(), { ...empty.format, borders: { top: null, bottom: null, left: null, right: null, diagonal: null } });
        grid.SelectAll();
        const all = grid.GetSelectionFormat();
        check('the whole sheet', [all.bold, all.fill, all.numberFormat, all.strikethrough, all.borders.top], ['mixed', 'mixed', 'mixed', false, null]);
        // It is asked for on every selection change.
        const seen = [];
        const off = grid.OnSelectionChange(() => seen.push(grid.GetSelectionFormat().bold));
        grid.SelectCell(2, 0); grid.SelectCell(0, 0); grid.SelectRange(0, 0, 2, 0);
        off();
        check('from a selection listener', seen, [true, false, 'mixed']);

        // Reading changes nothing.
        const revision = grid.GetRevision();
        grid.GetCell(0, 0); grid.GetRange('A:C'); grid.GetSelectionFormat();
        check('reading is not a change', [grid.GetRevision(), grid.IsEditing()], [revision, false]);

        // ---- an opened file: formats of a file, a sheet that is not converted yet ----
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        const names = grid.GetSheetNames();
        const other = names.findIndex((name, index) => index !== grid.GetActiveSheetIndex() && !grid.IsSheetLoaded(index));
        check('a sheet of the file is not converted yet', other >= 0, true);
        const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(names[other]) ? names[other] : `'${names[other].replace(/'/g, "''")}'`;
        const remote = grid.GetRange(`${quoted}!A1:F12`);
        check('GetRange reads it by name (and converts it)', [remote.length, remote[0].length, grid.IsSheetLoaded(other), grid.GetActiveSheetIndex() !== other], [12, 6, true, true]);
        check('reading another sheet is not a change', grid.IsModified(), false);
        grid.SwitchSheet(other);
        await frame();
        const local = grid.GetRange('A1:F12');
        check('the same cells once the sheet is shown', JSON.stringify(remote), JSON.stringify(local));
        const shown = local.flat().filter((cell) => cell.text !== '');
        check('the sheet has content there', shown.length > 0, true);
        const mismatched = shown.filter((cell) => { const node = host.querySelector(`.spreadsheet-cell[data-row="${cell.row}"][data-col="${cell.col}"]`); return node && !cell.isMerged && node.textContent !== cell.text && !/^#+$/.test(node.textContent); });
        check('texts are what the grid shows', mismatched.map((cell) => cell.address), []);
        const whole = grid.GetRange('A:A');
        const everything = grid.GetRange('1:1048576');
        check('a whole column of a file sheet is as long as its used range', [whole.length > 0 && whole.length <= grid.GetModel().RowCount, whole[0].length, whole.length], [true, 1, everything.length]);

        // The cost of GetSelectionFormat on a large sheet.
        await grid.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
        const data = grid.GetSheetNames().indexOf('Data');
        if (data >= 0) {
            grid.SwitchSheet(data);
            await frame();
            grid.SelectAll();
            let t0 = performance.now();
            const big = grid.GetSelectionFormat();
            const first = performance.now() - t0;
            t0 = performance.now();
            for (let i = 0; i < 100; i++) grid.GetSelectionFormat();
            const repeated = (performance.now() - t0) / 100;
            grid.SelectRange(0, 0, 99, 9);
            t0 = performance.now();
            grid.GetSelectionFormat();
            const small = performance.now() - t0;
            grid.SelectRange(0, 2, grid.GetModel().RowCount - 1, 2);
            t0 = performance.now();
            grid.GetSelectionFormat();
            const oneColumn = performance.now() - t0;
            note(`GetSelectionFormat on "Data" (${Object.keys(big).length} keys): whole sheet ${first.toFixed(1)} ms, again ${repeated.toFixed(3)} ms, 100 x 10 cells ${small.toFixed(1)} ms, a whole column ${oneColumn.toFixed(1)} ms`);
            check('the whole sheet of 846 000 cells is summarised within a frame or so', first < 250, true);
            check('asking again costs nothing', repeated < 0.5, true);
            check('a block on screen is immediate', small < 16, true);
        } else lines.push('FAIL the dashboard has no "Data" sheet');
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
