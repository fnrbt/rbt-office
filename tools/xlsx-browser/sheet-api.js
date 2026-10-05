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

// Formatting: the setters, ApplyFormat's commands, ApplyFormatPatch, the format painter,
// decimals, number format codes and the three kinds of Clear.
checks['api-format'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
    const css = (row, col) => getComputedStyle(node(row, col));
    const fmt = (row, col) => grid.GetCell(row, col).format;
    try {
        for (let row = 0; row < 12; row++) { grid.SetCell(row, 0, 'text ' + row); grid.SetCell(row, 1, String(1234.5 + row)); }
        const changes = [];
        grid.OnChange((change) => changes.push(change.Revision));

        // ---- font ----
        grid.SelectRange(0, 0, 1, 1);
        let revision = grid.GetRevision();
        grid.SetFontFamily('Georgia');
        await frame();
        check('SetFontFamily: every cell of the selection', [fmt(0, 0).fontFamily, fmt(1, 1).fontFamily, fmt(2, 0).fontFamily], ['Georgia', 'Georgia', null]);
        check('... is drawn in that font', /Georgia/.test(css(0, 0).fontFamily), true);
        check('... one change, a new revision', [changes.length, grid.GetRevision() !== revision], [1, true]);
        grid.Undo();
        check('... one undo step, back at the revision', [fmt(0, 0).fontFamily, fmt(1, 1).fontFamily, grid.GetRevision()], [null, null, revision]);
        grid.Redo();
        check('... redo', [fmt(0, 0).fontFamily, fmt(1, 1).fontFamily], ['Georgia', 'Georgia']);
        grid.SetFontFamily(null);
        check('SetFontFamily(null): the default font again', fmt(0, 0).fontFamily, null);
        grid.SetFontSize(16);
        await frame();
        check('SetFontSize', [fmt(0, 0).fontSize, fmt(1, 1).fontSize, css(0, 0).fontSize], [16, 16, '16px']);
        grid.SetFontSize(null);
        check('SetFontSize(null): the default size', fmt(0, 0).fontSize, null);
        check('SetFontSize refuses a size that is none', [thrown(() => grid.SetFontSize(0)), thrown(() => grid.SetFontSize('big'))].map((m) => /^SetFontSize: /.test(m ?? '')), [true, true]);
        check('SetFontFamily refuses an empty name', /^SetFontFamily: /.test(thrown(() => grid.SetFontFamily('  ')) ?? ''), true);

        // ---- strikethrough, underline ----
        grid.SelectCell(2, 0);
        grid.SetStrikethrough(true);
        await frame();
        check('SetStrikethrough', [fmt(2, 0).strikethrough, /line-through/.test(css(2, 0).textDecorationLine)], [true, true]);
        grid.ApplyFormat('underline');
        await frame();
        check('with an underline both lines are drawn', [/line-through/.test(css(2, 0).textDecorationLine), /underline/.test(css(2, 0).textDecorationLine)], [true, true]);
        grid.ApplyFormat('doubleUnderline');
        await frame();
        check('a double underline', [fmt(2, 0).doubleUnderline, fmt(2, 0).underline, css(2, 0).textDecorationStyle], [true, true, 'double']);
        grid.ApplyFormat('underline');
        check('underline off takes the double underline with it', [fmt(2, 0).underline, fmt(2, 0).doubleUnderline], [false, false]);
        grid.ApplyFormat('strikethrough');
        check("ApplyFormat('strikethrough') toggles by the active cell", fmt(2, 0).strikethrough, false);
        check('SetStrikethrough needs true or false', [thrown(() => grid.SetStrikethrough()), thrown(() => grid.SetStrikethrough('yes'))].map((m) => /^SetStrikethrough: /.test(m ?? '')), [true, true]);

        // ---- alignment ----
        grid.SelectRange(3, 0, 3, 1);
        grid.SetHorizontalAlign('right');
        await frame();
        check('SetHorizontalAlign', [fmt(3, 0).horizontalAlign, fmt(3, 1).horizontalAlign, css(3, 0).textAlign], ['right', 'right', 'right']);
        grid.SetHorizontalAlign('general');
        check("'general' is the default alignment", fmt(3, 0).horizontalAlign, null);
        for (const [command, want] of [['alignCenter', 'center'], ['alignJustify', 'justify'], ['alignLeft', 'left'], ['alignGeneral', null]]) { grid.ApplyFormat(command); check(`ApplyFormat('${command}')`, fmt(3, 1).horizontalAlign, want); }
        check('SetHorizontalAlign refuses what is no alignment', /^SetHorizontalAlign: .*"middle"/.test(thrown(() => grid.SetHorizontalAlign('middle')) ?? ''), true);
        // Vertical: in a tall row the text is at the top, in the middle, at the bottom.
        const { setRowHeight } = await import('/spreadsheet/Model.js');
        const textTop = (row, col) => { const range = document.createRange(); range.selectNodeContents(node(row, col)); const box = range.getBoundingClientRect(), cell = node(row, col).getBoundingClientRect(); return Math.round(box.top + box.height / 2 - cell.top); };
        grid.InstallWorkbook([setRowHeight(4, 80, grid.GetModel())], ['Sheet1'], 0, undefined);
        grid.SelectCell(4, 0);
        const places = {};
        for (const align of ['top', 'middle', 'bottom']) { grid.SetVerticalAlign(align); await frame(); places[align] = textTop(4, 0); check(`SetVerticalAlign('${align}')`, fmt(4, 0).verticalAlign, align); }
        check('the text is drawn at the top, in the middle, at the bottom of an 80 px row', [places.top < 20, Math.abs(places.middle - 40) <= 4, places.bottom > 60], [true, true, true]);
        note(`centre of the text in the 80 px row: top ${places.top}, middle ${places.middle}, bottom ${places.bottom}`);
        grid.ApplyFormat('alignTop'); check("ApplyFormat('alignTop')", fmt(4, 0).verticalAlign, 'top');
        grid.SetVerticalAlign(null); check('SetVerticalAlign(null)', fmt(4, 0).verticalAlign, null);
        check('SetVerticalAlign refuses what is no alignment', /^SetVerticalAlign: /.test(thrown(() => grid.SetVerticalAlign('left')) ?? ''), true);

        // ---- wrap ----
        grid.SetCell(4, 2, 'several words that are wider than one column of the grid and wrap');
        grid.SelectCell(4, 2);
        grid.SetWrapText(true);
        await frame();
        const wrapped = (() => { const range = document.createRange(); range.selectNodeContents(node(4, 2)); return range.getBoundingClientRect(); })();
        check('SetWrapText: the text takes several lines inside its cell', [fmt(4, 2).wrap, wrapped.height > 30, wrapped.width <= node(4, 2).getBoundingClientRect().width], [true, true, true]);
        check('... at the bottom of the cell, as Excel puts it', wrapped.bottom > node(4, 2).getBoundingClientRect().bottom - 8, true);
        grid.SetVerticalAlign('top');
        await frame();
        const wrappedTop = (() => { const range = document.createRange(); range.selectNodeContents(node(4, 2)); return range.getBoundingClientRect(); })();
        check('... and at the top when aligned so', wrappedTop.top < node(4, 2).getBoundingClientRect().top + 8, true);
        grid.ApplyFormat('wrap');
        check("ApplyFormat('wrap') toggles", fmt(4, 2).wrap, false);

        // ---- indent ----
        grid.SelectCell(5, 0);
        const paddingBefore = parseFloat(css(5, 0).paddingLeft);
        grid.SetIndent(2);
        await frame();
        check('SetIndent', [fmt(5, 0).indent, parseFloat(css(5, 0).paddingLeft) > paddingBefore + 10], [2, true]);
        grid.IncreaseIndent(); grid.IncreaseIndent(); grid.DecreaseIndent();
        check('IncreaseIndent / DecreaseIndent', fmt(5, 0).indent, 3);
        grid.SetIndent(0); grid.DecreaseIndent();
        check('no indent below none', fmt(5, 0).indent, 0);
        grid.SetHorizontalAlign('right'); grid.SetIndent(1);
        await frame();
        check('a right-aligned cell is indented from the right', [parseFloat(css(5, 0).paddingRight) > 10, parseFloat(css(5, 0).paddingLeft)], [true, paddingBefore]);
        check('SetIndent refuses what is no indent', [thrown(() => grid.SetIndent(-1)), thrown(() => grid.SetIndent(1.5)), thrown(() => grid.SetIndent(251))].map((m) => /^SetIndent: /.test(m ?? '')), [true, true, true]);

        // ---- rotation ----
        grid.SelectCell(6, 0);
        grid.SetTextRotation(45);
        await frame();
        const turned = node(6, 0).querySelector('.spreadsheet-rotated');
        check('SetTextRotation: the text is drawn turned', [fmt(6, 0).textRotation, turned?.textContent, /rotate\(-45deg\)/.test(turned?.style.transform ?? '')], [45, 'text 6', true]);
        grid.SetTextRotation(-90);
        await frame();
        check('clockwise', [fmt(6, 0).textRotation, /rotate\(90deg\)/.test(node(6, 0).querySelector('.spreadsheet-rotated')?.style.transform ?? '')], [-90, true]);
        grid.SetTextRotation(255);
        await frame();
        check('stacked', [fmt(6, 0).textRotation, node(6, 0).querySelector('.spreadsheet-rotated')?.style.writingMode], [255, 'vertical-rl']);
        grid.SetTextRotation(0);
        await frame();
        check('and back', [fmt(6, 0).textRotation, node(6, 0).querySelector('.spreadsheet-rotated'), node(6, 0).textContent], [0, null, 'text 6']);
        check('SetTextRotation refuses an angle that is none', [thrown(() => grid.SetTextRotation(120)), thrown(() => grid.SetTextRotation())].map((m) => /^SetTextRotation: /.test(m ?? '')), [true, true]);

        // ---- shrink to fit ----
        grid.SetCell(7, 2, 'a text much wider than the column it is in'); grid.SetCell(7, 3, 'next');
        grid.SelectCell(7, 2);
        grid.SetShrinkToFit(true);
        await frame();
        const shrunkWidth = (() => { const range = document.createRange(); range.selectNodeContents(node(7, 2)); return range.getBoundingClientRect().width; })();
        check('SetShrinkToFit: the text is drawn smaller and fits', [fmt(7, 2).shrinkToFit, parseFloat(css(7, 2).fontSize) < 11, shrunkWidth <= node(7, 2).getBoundingClientRect().width], [true, true, true]);
        grid.ApplyFormat('shrinkToFit');
        await frame();
        check("ApplyFormat('shrinkToFit') toggles", [fmt(7, 2).shrinkToFit, css(7, 2).fontSize], [false, '11px']);

        // ---- number formats, decimals ----
        grid.SelectRange(8, 1, 9, 1);
        grid.SetNumberFormat('#,##0.00');
        await frame();
        check('SetNumberFormat', [fmt(8, 1).numberFormat, grid.GetCell(8, 1).text, node(9, 1).textContent], ['#,##0.00', '1,242.50', '1,243.50']);
        grid.IncreaseDecimals();
        check('IncreaseDecimals', [fmt(8, 1).numberFormat, grid.GetCell(9, 1).text], ['#,##0.000', '1,243.500']);
        grid.DecreaseDecimals(); grid.DecreaseDecimals(); grid.DecreaseDecimals();
        check('DecreaseDecimals down to none', [fmt(8, 1).numberFormat, grid.GetCell(8, 1).text], ['#,##0', '1,243']);
        const beforeNone = grid.GetRevision();
        grid.DecreaseDecimals();
        check('nothing left to take away: no change', [fmt(8, 1).numberFormat, grid.GetRevision()], ['#,##0', beforeNone]);
        grid.SetNumberFormat('General');
        check("'General' is no number format", fmt(8, 1).numberFormat, null);
        grid.ApplyFormat('decimalIncrease');
        check('a General number starts from the decimals it shows', [fmt(8, 1).numberFormat, grid.GetCell(8, 1).text, fmt(8, 0).numberFormat], ['0.00', '1242.50', null]);
        grid.Undo();
        check('... one undo step for the selection', [fmt(8, 1).numberFormat, fmt(9, 1).numberFormat], [null, null]);
        grid.SetNumberFormat('yyyy-mm-dd');
        grid.IncreaseDecimals();
        check('a date format has no decimals to add', fmt(8, 1).numberFormat, 'yyyy-mm-dd');
        check('SetNumberFormat refuses codes no format can be', [thrown(() => grid.SetNumberFormat('0.0 "open')), thrown(() => grid.SetNumberFormat('[Red 0')), thrown(() => grid.SetNumberFormat('')), thrown(() => grid.SetNumberFormat(5)), thrown(() => grid.SetNumberFormat('0;0;0;0;0'))].map((m) => /^SetNumberFormat: /.test(m ?? '')), [true, true, true, true, true]);
        grid.SetNumberFormat(null);

        // ---- ApplyFormatPatch ----
        grid.SelectRange(10, 0, 11, 1);
        revision = grid.GetRevision();
        grid.SelectCell(10, 0); grid.ApplyFormat('italic'); grid.SelectRange(10, 0, 11, 1);
        grid.ApplyFormatPatch({ bold: true, fill: '#fff2cc', textColor: '#C00000', borders: { bottom: { style: 'medium', color: '#000000' } } });
        await frame();
        check('ApplyFormatPatch sets its keys and leaves the others', [fmt(10, 0).bold, fmt(10, 0).italic, fmt(11, 1).bold, fmt(11, 1).italic, fmt(11, 1).fill, fmt(11, 1).textColor, fmt(11, 1).borders.bottom, fmt(11, 1).borders.top],
            [true, true, true, false, '#FFF2CC', '#C00000', { style: 'medium', color: '#000000' }, null]);
        check('... and is drawn', [css(11, 1).fontWeight, css(11, 1).backgroundColor, css(11, 1).color], ['700', 'rgb(255, 242, 204)', 'rgb(192, 0, 0)']);
        grid.Undo();
        check('... as one undo step', [fmt(10, 0).bold, fmt(10, 0).italic, fmt(11, 1).fill], [false, true, null]);
        grid.Redo();
        check('ApplyFormatPatch refuses unknown keys and wrong values', [thrown(() => grid.ApplyFormatPatch({ colour: 'red' })), thrown(() => grid.ApplyFormatPatch({ bold: 'yes' })), thrown(() => grid.ApplyFormatPatch('bold')), thrown(() => grid.ApplyFormatPatch({ fill: 'reddish' }))].map((m) => /^ApplyFormatPatch: /.test(m ?? '')), [true, true, true, true]);
        check("ApplyFormat refuses an unknown command", /^ApplyFormat: unknown command "sparkle"/.test(thrown(() => grid.ApplyFormat('sparkle')) ?? ''), true);

        // ---- format painter ----
        const picked = grid.GetFormatOf('B12');
        check('GetFormatOf gives the format of GetCell', picked, fmt(11, 1));
        grid.SelectCell(0, 3); grid.SetCell(0, 3, 'painted'); grid.SetCell(1, 4, '7'); grid.SelectCell(1, 4); grid.ApplyFormat('italic'); grid.SetNumberFormat('0.0');
        revision = grid.GetRevision();
        grid.ApplyFormatTo('D1:E3', picked);
        await frame();
        check('ApplyFormatTo gives every cell of the range that format, whatever it had', [fmt(0, 3), fmt(1, 4), fmt(2, 3)], [picked, picked, picked]);
        check('... values stay, the selection too', [grid.GetCell(0, 3).value, grid.GetCell(1, 4).value, grid.GetSelectionRef()], ['painted', 7, 'E2']);
        grid.Undo();
        check('... one undo step', [fmt(0, 3).bold, fmt(1, 4).italic, fmt(1, 4).numberFormat, fmt(2, 3).fill, grid.GetRevision()], [false, true, '0.0', null, revision]);
        grid.Redo();
        check('GetFormatOf of a range is its first cell', grid.GetFormatOf('D1:E3'), picked);
        check('the painter refuses what is no range or no format', [thrown(() => grid.ApplyFormatTo('D0', picked)), thrown(() => grid.ApplyFormatTo('D1', null)), thrown(() => grid.GetFormatOf(7)), thrown(() => grid.ApplyFormatTo('D1', { bold: 1 }))].map((m) => /^(ApplyFormatTo|GetFormatOf): /.test(m ?? '')), [true, true, true, true]);
        // A whole column of this 1 000-row sheet is 1 000 cells: each gets the format. (Over 20 000
        // cells - a whole column of a large sheet - only the cells that exist are changed.)
        let t0 = performance.now();
        grid.ApplyFormatTo('A:A', { italic: true });
        const wholeColumn = performance.now() - t0;
        check('a whole column', [fmt(0, 0).italic, fmt(11, 0).italic, fmt(500, 0).italic, fmt(0, 0).bold, fmt(0, 1).italic, wholeColumn < 200], [true, true, true, false, false, true]);
        grid.Undo();

        // ---- export, re-import ----
        grid.SelectCell(2, 0); grid.SetStrikethrough(true); grid.SetFontFamily('Georgia'); grid.SetFontSize(14);
        grid.SelectCell(3, 0); grid.SetHorizontalAlign('right'); grid.SetVerticalAlign('top'); grid.SetWrapText(true); grid.SetIndent(2);
        grid.SelectCell(6, 0); grid.SetTextRotation(-30);
        grid.SelectCell(7, 2); grid.SetShrinkToFit(true);
        grid.SelectCell(9, 1); grid.SetNumberFormat('#,##0.0;[Red]-#,##0.0');
        const keys = ['bold', 'italic', 'strikethrough', 'fontFamily', 'fontSize', 'textColor', 'fill', 'horizontalAlign', 'verticalAlign', 'wrap', 'indent', 'shrinkToFit', 'textRotation', 'numberFormat'];
        const cells = [[2, 0], [3, 0], [6, 0], [7, 2], [9, 1], [11, 1], [0, 3]];
        // (A file names the default font where the editor leaves it out.)
        const snapshot = (g) => cells.map(([row, col]) => { const format = pick(g.GetCell(row, col).format, keys); if (format.fontFamily === 'Calibri') format.fontFamily = null; if (format.fontSize === 11) format.fontSize = null; return [format, g.GetCell(row, col).format.borders.bottom]; });
        const before = snapshot(grid);
        const again = await reopened(grid);
        check('every format comes back from an exported file', snapshot(again.grid), before);
        check('... and the texts they give', cells.map(([row, col]) => again.grid.GetCell(row, col).text), cells.map(([row, col]) => grid.GetCell(row, col).text));
        again.close();

        // ---- clear ----
        grid.SetCell(14, 0, '5'); grid.SetCell(15, 0, '=A15*2'); grid.SetCell(16, 0, '=A16+1');
        grid.SelectRange(14, 0, 15, 0); grid.ApplyFormat('bold'); grid.SetComment(14, 0, 'a note', 'Ada');
        revision = grid.GetRevision();
        grid.ClearContents();
        await frame();
        check('ClearContents: values and formulas go, formats and comments stay', [grid.GetCell(14, 0).value, grid.GetCell(15, 0).formula, fmt(14, 0).bold, fmt(15, 0).bold, grid.GetCell(14, 0).comment?.text], [null, null, true, true, 'a note']);
        check('... what read them is recalculated', grid.GetCell(16, 0).value, 1);
        grid.Undo();
        check('... one undo step', [grid.GetCell(14, 0).value, grid.GetCell(15, 0).formula, grid.GetCell(16, 0).value, grid.GetRevision()], [5, '=A15*2', 11, revision]);
        grid.ClearFormats();
        check('ClearFormats: formats go, values and comments stay', [fmt(14, 0).bold, fmt(15, 0).bold, grid.GetCell(14, 0).value, grid.GetCell(15, 0).value, grid.GetCell(14, 0).comment?.text], [false, false, 5, 10, 'a note']);
        grid.Undo();
        check('... one undo step', [fmt(14, 0).bold, fmt(15, 0).bold], [true, true]);
        grid.ClearAll();
        check('ClearAll: contents, formats and comments', [grid.GetCell(14, 0).value, fmt(14, 0).bold, grid.GetCell(14, 0).comment, grid.GetCell(15, 0).value, grid.GetCell(16, 0).value], [null, false, null, null, 1]);
        grid.Undo();
        check('... one undo step', [grid.GetCell(14, 0).value, fmt(14, 0).bold, grid.GetCell(14, 0).comment?.text, grid.GetCell(16, 0).value], [5, true, 'a note', 11]);
        grid.SelectCell(20, 5);
        grid.ClearContents('A15:A16');
        check('Clear takes a range instead of the selection', [grid.GetCell(14, 0).value, grid.GetCell(15, 0).value, fmt(14, 0).bold, grid.GetSelectionRef()], [null, null, true, 'F21']);
        grid.Undo();
        const quiet = grid.GetRevision();
        grid.ClearAll('H30:J40');
        check('clearing empty cells changes nothing', [grid.GetRevision(), grid.CanRedo()], [quiet, true]);
        check('Clear refuses what is no range', [thrown(() => grid.ClearContents('nowhere')), thrown(() => grid.ClearFormats(3)), thrown(() => grid.ClearAll('A0'))].map((m) => /^Clear(Contents|Formats|All): /.test(m ?? '')), [true, true, true]);
        // A spilled array goes with its formula.
        grid.SetCell(20, 0, '=SEQUENCE(3)');
        check('(a formula spilled)', [grid.GetCell(21, 0).value, grid.GetCell(22, 0).value], [2, 3]);
        grid.ClearContents('A21');
        check('ClearContents of a spilling formula takes its array back', [grid.GetCell(20, 0).value, grid.GetCell(21, 0).value, grid.GetCell(22, 0).value], [null, null, null]);
        grid.Undo();
        check('... and undo spills again', [grid.GetCell(20, 0).value, grid.GetCell(21, 0).value, grid.GetCell(22, 0).value], [1, 2, 3]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
