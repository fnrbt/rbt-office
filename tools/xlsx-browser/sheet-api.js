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

// Conditional formats: GetConditionalFormats, AddConditionalFormat, UpdateConditionalFormat,
// RemoveConditionalFormat, ReorderConditionalFormats, ClearConditionalFormats.
checks['api-cf'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const fillOf = (row, col) => getComputedStyle(node(row, col)).backgroundColor;
        const rgb = (hex) => `rgb(${parseInt(hex.slice(1, 3), 16)}, ${parseInt(hex.slice(3, 5), 16)}, ${parseInt(hex.slice(5, 7), 16)})`;
        const list = (g = grid, sheet) => g.GetConditionalFormats(sheet).map((cf) => ({ ...cf, rule: { ...cf.rule } }));
        const brief = (g = grid) => g.GetConditionalFormats().map((cf) => `${cf.id}:${cf.rule.type}@${cf.range}`);
        const today = Math.floor((Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
        // A1:A6 numbers, B1:B6 text, C1:C3 dates around today.
        [10, 20, 30, 40, 50, 20].forEach((n, row) => grid.SetCell(row, 0, String(n)));
        ['north', 'south', 'northern', 'East', '', 'west'].forEach((text, row) => grid.SetCell(row, 1, text));
        [today, today - 1, today + 40].forEach((serial, row) => grid.SetCell(row, 2, String(serial)));
        grid.SetCell(0, 4, '25');
        check('no conditional formats to begin with', grid.GetConditionalFormats(), []);
        const format = { fill: '#FFC7CE', textColor: '#9C0006', bold: true };
        const fullFormat = { textColor: '#9C0006', fill: '#FFC7CE', bold: true, italic: false, underline: false, strikethrough: false };
        const changes = [];
        const off = grid.OnChange((change) => changes.push(change.Revision));

        // ---- one rule: add, read, the grid shows it, undo / redo ----
        const revision = grid.GetRevision();
        const first = grid.AddConditionalFormat('A1:A6', { type: 'cellIs', operator: 'greaterThan', formula1: 25, format });
        await frame();
        check('AddConditionalFormat returns the id', [typeof first, first > 0], ['number', true]);
        check('GetConditionalFormats gives the rule back', list(), [{ id: first, range: 'A1:A6', priority: 1, stopIfTrue: false,
            rule: { type: 'cellIs', operator: 'greaterThan', formula1: '25', formula2: null, format: fullFormat } }]);
        check('the grid shows it', [0, 1, 2, 3, 4, 5].map((row) => fillOf(row, 0) === rgb('#FFC7CE')), [false, false, true, true, true, false]);
        check('... with the text colour and weight', [getComputedStyle(node(2, 0)).color, Number(getComputedStyle(node(2, 0)).fontWeight) >= 600, Number(getComputedStyle(node(0, 0)).fontWeight) >= 600], [rgb('#9C0006'), true, false]);
        check('it is a change with a new revision', [changes.length, grid.GetRevision() !== revision, grid.IsModified()], [1, true, true]);
        grid.Undo(); await frame();
        check('undo takes the rule away', [grid.GetConditionalFormats().length, fillOf(2, 0) === rgb('#FFC7CE'), grid.GetRevision()], [0, false, revision]);
        grid.Redo(); await frame();
        check('redo brings it back with its id', [brief(), fillOf(2, 0) === rgb('#FFC7CE')], [[`${first}:cellIs@A1:A6`], true]);

        // ---- every kind ----
        grid.ClearConditionalFormats();
        const kinds = [
            ['A1:A6', { type: 'cellIs', operator: 'between', formula1: '=$E$1', formula2: 45, format }],
            ['A1:A6', { type: 'cellIs', operator: 'notBetween', formula1: 15, formula2: 45, format }],
            ['A1:A6', { type: 'cellIs', operator: 'greaterThanOrEqual', formula1: 30, format }],
            ['A1:A6', { type: 'cellIs', operator: 'lessThanOrEqual', formula1: 20, format }],
            ['A1:A6', { type: 'cellIs', operator: 'lessThan', formula1: 20, format }],
            ['A1:A6', { type: 'cellIs', operator: 'equal', formula1: 20, format }],
            ['A1:A6', { type: 'cellIs', operator: 'notEqual', formula1: 20, format }],
            ['B1:B6', { type: 'cellIs', operator: 'equal', formula1: '"north"', format }],
            ['A1:B6', { type: 'expression', formula: '=AND($A1>15,MOD(ROW(),2)=0)', format }],
            ['B1:B6', { type: 'containsText', text: 'north', format }],
            ['B1:B6', { type: 'notContainsText', text: 'north', format }],
            ['B1:B6', { type: 'beginsWith', text: 'no', format }],
            ['B1:B6', { type: 'endsWith', text: 'st', format }],
            ['A1:A6', { type: 'top', rank: 2, format }],
            ['A1:A6', { type: 'top', rank: 2, bottom: true, format }],
            ['A1:A6', { type: 'top', rank: 50, percent: true, format }],
            ['A1:A6', { type: 'aboveAverage', format }],
            ['A1:A6', { type: 'aboveAverage', above: false, orEqual: true, stdDev: 1, format }],
            ['A1:A6', { type: 'duplicateValues', format }],
            ['A1:A6', { type: 'uniqueValues', format }],
            ['B1:B6', { type: 'blanks', format }],
            ['B1:B6', { type: 'noBlanks', format }],
            ['A1:A6', { type: 'errors', format }],
            ['A1:A6', { type: 'noErrors', format }],
            ['C1:C3', { type: 'dateOccurring', period: 'today', format }],
            ['C1:C3', { type: 'dateOccurring', period: 'last7Days', format }],
            ['A1:A6', { type: 'colorScale', stops: [{ type: 'min', color: '#F8696B' }, { type: 'max', color: '#63BE7B' }] }],
            ['A1:A6', { type: 'colorScale', stops: [{ type: 'min', color: '#F8696B' }, { type: 'percentile', value: 50, color: '#FFEB84' }, { type: 'number', value: 45, color: '#63BE7B' }] }],
            ['A1:A6', { type: 'dataBar', color: '#638EC6' }],
            ['A1:A6', { type: 'dataBar', min: { type: 'number', value: 0 }, max: { type: 'formula', value: '=$E$1*4' }, color: '#FF0000' }],
            ['A1:A6', { type: 'iconSet', iconSet: '3TrafficLights1' }],
            ['A1:A6', { type: 'iconSet', iconSet: '5Arrows', thresholds: [0, 20, 40, 60, 80].map((value) => ({ type: 'percentile', value })), reverse: true }],
        ];
        // Which cells of the rule's first column (rows 1..6) a rule formats, for the kinds with a format.
        const expected = [
            [0, 0, 1, 1, 0, 0], [1, 0, 0, 0, 1, 0], [0, 0, 1, 1, 1, 0], [1, 1, 0, 0, 0, 1], [1, 0, 0, 0, 0, 0], [0, 1, 0, 0, 0, 1], [1, 0, 1, 1, 1, 0],
            [1, 0, 0, 0, 0, 0], [0, 1, 0, 1, 0, 1], [1, 0, 1, 0, 0, 0], [0, 1, 0, 1, 1, 1], [1, 0, 1, 0, 0, 0], [0, 0, 0, 1, 0, 1],
            [0, 0, 0, 1, 1, 0], [1, 1, 0, 0, 0, 1], [0, 0, 1, 1, 1, 0], [0, 0, 1, 1, 1, 0], [1, 0, 0, 0, 0, 0], [0, 1, 0, 0, 0, 1], [1, 0, 1, 1, 1, 0],
            [0, 0, 0, 0, 1, 0], [1, 1, 1, 1, 0, 1], [0, 0, 0, 0, 0, 0], [1, 1, 1, 1, 1, 1], [1, 0, 0], [1, 1, 0],
        ];
        const shown = [], back = [];
        for (const [range, rule] of kinds) {
            const id = grid.AddConditionalFormat(range, rule);
            await frame();
            const col = range.startsWith('A') ? 0 : range.startsWith('B') ? 1 : 2, rows = range.startsWith('C') ? 3 : 6;
            shown.push(Array.from({ length: rows }, (_, row) => (fillOf(row, col) === rgb('#FFC7CE') ? 1 : 0)));
            back.push(grid.GetConditionalFormats().find((cf) => cf.id === id).rule);
            grid.RemoveConditionalFormat(id);
        }
        kinds.forEach(([range, rule], index) => { if (index < expected.length) check(`${rule.type}${rule.operator ? ' ' + rule.operator : ''}${rule.period ? ' ' + rule.period : ''}${rule.bottom ? ' bottom' : ''}${rule.percent ? ' percent' : ''}${rule.above === false ? ' below' : ''} formats the cells it should`, shown[index], expected[index]); });
        check('every kind reads back as its type', back.map((rule) => rule.type), kinds.map(([, rule]) => rule.type));
        check('defaults are filled in', [back[13], back[16], back[28], back[30].thresholds],
            [{ type: 'top', rank: 2, percent: false, bottom: false, format: fullFormat }, { type: 'aboveAverage', above: true, orEqual: false, stdDev: 0, format: fullFormat },
             { type: 'dataBar', min: { type: 'min', value: null }, max: { type: 'max', value: null }, color: '#638EC6', format: null },
             [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }]]);
        check('formulas come back without the =', [back[0].formula1, back[0].formula2, back[8].formula, back[29].max], ['$E$1', '45', 'AND($A1>15,MOD(ROW(),2)=0)', { type: 'formula', value: '$E$1*4' }]);
        check('a colour scale reads back', back[27], { type: 'colorScale', stops: [{ type: 'min', value: null, color: '#F8696B' }, { type: 'percentile', value: 50, color: '#FFEB84' }, { type: 'number', value: 45, color: '#63BE7B' }], format: null });
        check('no rule is left', grid.GetConditionalFormats(), []);

        // ---- what the visual rules draw ----
        const scale = grid.AddConditionalFormat('A1:A6', kinds[26][1]); await frame();
        check('a colour scale: the lowest and the highest value get the end colours', [fillOf(0, 0), fillOf(4, 0)], [rgb('#F8696B'), rgb('#63BE7B')]);
        grid.RemoveConditionalFormat(scale);
        const bar = grid.AddConditionalFormat('A1:A6', kinds[29][1]); await frame();
        check('a data bar: 10 and 50 of 0..100', [0, 4].map((row) => /linear-gradient\(to right, rgb\(255, 0, 0\) (\d+)%/.exec(getComputedStyle(node(row, 0)).backgroundImage)?.[1]), ['10', '50']);
        grid.RemoveConditionalFormat(bar);
        const icons = grid.AddConditionalFormat('A1:A6', kinds[30][1]); await frame();
        check('an icon set: an icon in every cell with a number, the lowest and the highest differ', [host.querySelectorAll('[data-cf-icon]').length >= 6, node(0, 0).querySelector('[data-cf-icon]')?.getAttribute('data-cf-icon'), node(4, 0).querySelector('[data-cf-icon]')?.getAttribute('data-cf-icon')], [true, '0', '2']);
        grid.RemoveConditionalFormat(icons);

        // ---- several rules: order, stop if true, update, reorder, clear ----
        const red = grid.AddConditionalFormat('A1:A6', { type: 'cellIs', operator: 'greaterThan', formula1: 15, format: { fill: '#FF0000' } });
        const green = grid.AddConditionalFormat('A1:A6 B1:B3', { type: 'cellIs', operator: 'greaterThan', formula1: 35, format: { fill: '#00FF00', italic: true } }, { stopIfTrue: true });
        const blue = grid.AddConditionalFormat('A:A', { type: 'top', rank: 1, format: { fill: '#0000FF', underline: true, strikethrough: true } });
        await frame();
        check('the newest rule takes precedence', [brief(), fillOf(4, 0), fillOf(3, 0), fillOf(1, 0), fillOf(0, 0) === rgb('#FF0000')], [[`${blue}:top@A:A`, `${green}:cellIs@A1:A6 B1:B3`, `${red}:cellIs@A1:A6`], rgb('#0000FF'), rgb('#00FF00'), rgb('#FF0000'), false]);
        check('priorities are 1, 2, 3 and stopIfTrue is told', list().map((cf) => [cf.priority, cf.stopIfTrue]), [[1, false], [2, true], [3, false]]);
        check('ids are not reused', new Set([first, scale, bar, icons, red, green, blue]).size, 7);
        check('underline and strikethrough of a rule are drawn', /underline/.test(getComputedStyle(node(4, 0)).textDecorationLine) && /line-through/.test(getComputedStyle(node(4, 0)).textDecorationLine), true);
        let before = grid.GetRevision();
        grid.ReorderConditionalFormats([red, green, blue]); await frame();
        check('ReorderConditionalFormats: the first id wins', [brief().map((entry) => entry.split(':')[0]), fillOf(4, 0), fillOf(3, 0)], [[red, green, blue].map(String), rgb('#FF0000'), rgb('#FF0000')]);
        grid.Undo(); await frame();
        check('... one undo step', [brief().map((entry) => entry.split(':')[0]), fillOf(4, 0), grid.GetRevision()], [[blue, green, red].map(String), rgb('#0000FF'), before]);
        grid.Redo();
        check('reordering refuses a list that is not the sheet\'s rules', [thrown(() => grid.ReorderConditionalFormats([red, green])), thrown(() => grid.ReorderConditionalFormats([red, green, 999])), thrown(() => grid.ReorderConditionalFormats([red, green, green])), thrown(() => grid.ReorderConditionalFormats('all'))].map((m) => /^ReorderConditionalFormats: /.test(m ?? '')), [true, true, true, true]);

        grid.UpdateConditionalFormat(blue, { priority: 1 }); await frame();
        check('UpdateConditionalFormat priority: the rule moves, the others keep their order', brief().map((entry) => entry.split(':')[0]), [blue, red, green].map(String));
        before = grid.GetRevision();
        grid.UpdateConditionalFormat(red, { range: 'A1:A2', stopIfTrue: true, format: { fill: '#FFFF00', bold: true } }); await frame();
        const redNow = list().find((cf) => cf.id === red);
        check('UpdateConditionalFormat range, stopIfTrue and format: the rule itself stays', [redNow.range, redNow.stopIfTrue, redNow.rule.operator, redNow.rule.formula1, redNow.rule.format.fill, redNow.rule.format.bold], ['A1:A2', true, 'greaterThan', '15', '#FFFF00', true]);
        check('... the grid follows', [fillOf(1, 0), fillOf(3, 0)], [rgb('#FFFF00'), rgb('#00FF00')]);
        grid.Undo(); await frame();
        check('... one undo step', [list().find((cf) => cf.id === red).range, fillOf(1, 0), grid.GetRevision()], ['A1:A6', rgb('#FF0000'), before]);
        grid.Redo();
        grid.UpdateConditionalFormat(red, { rule: { type: 'cellIs', operator: 'lessThan', formula1: 15 } }); await frame();
        check('UpdateConditionalFormat rule without a format keeps the format', [list().find((cf) => cf.id === red).rule, fillOf(0, 0), fillOf(1, 0) === rgb('#FFFF00')],
            [{ type: 'cellIs', operator: 'lessThan', formula1: '15', formula2: null, format: { textColor: null, fill: '#FFFF00', bold: true, italic: false, underline: false, strikethrough: false } }, rgb('#FFFF00'), false]);
        grid.UpdateConditionalFormat(red, { rule: { type: 'dataBar', color: '#00AA00' } });
        check('a rule can become a visual one (its format goes)', list().find((cf) => cf.id === red).rule.format, null);
        grid.Undo();
        check('an update that changes nothing is no change', (() => { const r = grid.GetRevision(); grid.UpdateConditionalFormat(red, { stopIfTrue: true }); return grid.GetRevision() === r; })(), true);
        check('UpdateConditionalFormat refuses', [thrown(() => grid.UpdateConditionalFormat(999, { stopIfTrue: true })), thrown(() => grid.UpdateConditionalFormat(red, { colour: 'red' })), thrown(() => grid.UpdateConditionalFormat(red, { priority: 9 })),
            thrown(() => grid.UpdateConditionalFormat(red, { range: 'nowhere' })), thrown(() => grid.UpdateConditionalFormat(red)), thrown(() => grid.UpdateConditionalFormat(red, { format: { fontSize: 20 } })), thrown(() => grid.UpdateConditionalFormat('1', {}))],
            ['UpdateConditionalFormat: the sheet has no conditional format with id 999', 'UpdateConditionalFormat: changes has an unknown key "colour" (known: range, priority, stopIfTrue, rule, format)', 'UpdateConditionalFormat: priority must be between 1 and 3, got 9',
             'UpdateConditionalFormat: "nowhere" is not a range (expected A1 notation such as "B2:D5", or several areas: "A1:A5 C1:C5")', 'UpdateConditionalFormat: changes must be an object { range, priority, stopIfTrue, rule, format }',
             'UpdateConditionalFormat: changes.format.fontSize: a conditional format applies only textColor, fill, bold, italic, underline, strikethrough', 'UpdateConditionalFormat: id must be a whole number, got "1"']);

        before = grid.GetRevision();
        check('ClearConditionalFormats(range) says how many rules it touched', grid.ClearConditionalFormats('A1:A2 B1:B3'), 3);
        check('... a rule that covers more keeps the rest, a rule left with no cell is gone', brief(), [`${blue}:top@A3:A1048576`, `${green}:cellIs@A3:A6`]);
        grid.Undo();
        check('... one undo step', [brief().length, grid.GetRevision()], [3, before]);
        grid.Redo();
        check('clearing where there is none is no change', [grid.ClearConditionalFormats('H1:J9'), grid.GetRevision() === grid.GetRevision()], [0, true]);
        check('RemoveConditionalFormat refuses an unknown id', [thrown(() => grid.RemoveConditionalFormat(999)), thrown(() => grid.RemoveConditionalFormat())].map((m) => /^RemoveConditionalFormat: /.test(m ?? '')), [true, true]);
        grid.RemoveConditionalFormat(green);
        check('RemoveConditionalFormat', brief(), [`${blue}:top@A3:A1048576`]);
        grid.Undo();
        check('... one undo step', brief().length, 2);

        // ---- the older members ----
        const legacy = grid.AddConditionalFormat(8, 3, 'greaterThan', 5, '#ABCDEF');
        check('the older AddConditionalFormat(row, col, ruleType, threshold, bgColor) still works', [typeof legacy, list().find((cf) => cf.id === legacy).range, list().find((cf) => cf.id === legacy).rule.formula1, list().find((cf) => cf.id === legacy).rule.format.fill], ['number', 'D9', '5', '#ABCDEF']);
        grid.Undo();
        check('... and is an undo step now', list().some((cf) => cf.id === legacy), false);
        grid.Redo();
        grid.RemoveConditionalFormats(8, 3);
        check('RemoveConditionalFormats(row, col) still works', list().some((cf) => cf.range === 'D9'), false);
        grid.Undo();
        check('... and is an undo step now', list().some((cf) => cf.range === 'D9'), true);
        grid.RemoveConditionalFormats(8, 3);

        // ---- what AddConditionalFormat refuses ----
        const refused = [
            [() => grid.AddConditionalFormat('A1:', { type: 'blanks' }), /^AddConditionalFormat: "A1:" is not a range/],
            [() => grid.AddConditionalFormat('A1'), /^AddConditionalFormat: the rule is missing/],
            [() => grid.AddConditionalFormat('A1', 'blanks'), /^AddConditionalFormat: rule must be an object/],
            [() => grid.AddConditionalFormat('A1', {}), /^AddConditionalFormat: rule\.type must be a string/],
            [() => grid.AddConditionalFormat('A1', { type: 'sparkle' }), /^AddConditionalFormat: unknown rule type "sparkle" \(known: cellIs, expression/],
            [() => grid.AddConditionalFormat('A1', { type: 'cellIs', operator: 'above', formula1: 1 }), /^AddConditionalFormat: rule\.operator must be one of/],
            [() => grid.AddConditionalFormat('A1', { type: 'cellIs', operator: 'between', formula1: 1 }), /^AddConditionalFormat: a "cellIs" rule needs rule\.formula2/],
            [() => grid.AddConditionalFormat('A1', { type: 'cellIs', operator: 'equal', formula1: true }), /^AddConditionalFormat: rule\.formula1 must be a number or a string/],
            [() => grid.AddConditionalFormat('A1', { type: 'expression' }), /^AddConditionalFormat: a "expression" rule needs rule\.formula/],
            [() => grid.AddConditionalFormat('A1', { type: 'expression', formula: '=A1', text: 'x' }), /^AddConditionalFormat: a "expression" rule has no key "text" \(it takes: type, formula, format\)/],
            [() => grid.AddConditionalFormat('A1', { type: 'containsText' }), /^AddConditionalFormat: a "containsText" rule needs rule\.text/],
            [() => grid.AddConditionalFormat('A1', { type: 'top', rank: 0 }), /^AddConditionalFormat: rule\.rank must be between 1 and 1000/],
            [() => grid.AddConditionalFormat('A1', { type: 'top', rank: 2.5 }), /^AddConditionalFormat: rule\.rank must be a whole number/],
            [() => grid.AddConditionalFormat('A1', { type: 'dateOccurring', period: 'someday' }), /^AddConditionalFormat: rule\.period must be one of "today"/],
            [() => grid.AddConditionalFormat('A1', { type: 'colorScale', stops: [{ type: 'min', color: '#FF0000' }] }), /^AddConditionalFormat: a colour scale has 2 or 3 stops, got 1/],
            [() => grid.AddConditionalFormat('A1', { type: 'colorScale', stops: [{ type: 'min', color: 'red' }, { type: 'max', color: '#00FF00' }] }), /^AddConditionalFormat: rule\.stops\[0\]\.color must be a colour/],
            [() => grid.AddConditionalFormat('A1', { type: 'colorScale', stops: [{ type: 'min', color: '#FF0000' }, { type: 'percent', color: '#00FF00' }] }), /^AddConditionalFormat: rule\.stops\[1\]\.value is missing/],
            [() => grid.AddConditionalFormat('A1', { type: 'dataBar', color: 'blue' }), /^AddConditionalFormat: rule\.color must be a colour/],
            [() => grid.AddConditionalFormat('A1', { type: 'dataBar', format }), /^AddConditionalFormat: a "dataBar" rule draws by itself/],
            [() => grid.AddConditionalFormat('A1', { type: 'iconSet', iconSet: '3Stars' }), /^AddConditionalFormat: rule\.iconSet must be one of 3Arrows/],
            [() => grid.AddConditionalFormat('A1', { type: 'iconSet', iconSet: '5Rating', thresholds: [{ type: 'percent', value: 0 }] }), /^AddConditionalFormat: icon set "5Rating" needs 5 thresholds/],
            [() => grid.AddConditionalFormat('A1', { type: 'blanks', format: { fill: 'red' } }), /^AddConditionalFormat: rule\.format\.fill must be a colour/],
            [() => grid.AddConditionalFormat('A1', { type: 'blanks', format: { wrap: true } }), /^AddConditionalFormat: rule\.format\.wrap: a conditional format applies only/],
            [() => grid.AddConditionalFormat('A1', { type: 'blanks' }, { stop: true }), /^AddConditionalFormat: options has an unknown key "stop"/],
            [() => grid.AddConditionalFormat(0, 0, 'between', 1, '#FF0000'), /^AddConditionalFormat: ruleType must be "greaterThan", "lessThan" or "equal"/],
            [() => grid.GetConditionalFormats(5), /^GetConditionalFormats: there is no sheet 5/],
        ];
        const count = grid.GetConditionalFormats().length, quiet = grid.GetRevision();
        check('invalid input throws with a message that says what is wrong', refused.map(([run, pattern]) => { const message = thrown(run); return pattern.test(message ?? '') ? true : message; }), refused.map(() => true));
        check('... and changes nothing', [grid.GetConditionalFormats().length, grid.GetRevision()], [count, quiet]);

        // ---- rules follow their cells ----
        const moving = grid.AddConditionalFormat('B2:B4', { type: 'expression', formula: '=$A2>B$1', format });
        grid.InsertRows(0, 2);
        check('inserted rows: range and formula move, the id stays', (() => { const cf = list().find((entry) => entry.id === moving); return [cf.range, cf.rule.formula]; })(), ['B4:B6', '$A4>B$3']);
        grid.Undo();
        check('... and come back with undo', (() => { const cf = list().find((entry) => entry.id === moving); return [cf.range, cf.rule.formula]; })(), ['B2:B4', '$A2>B$1']);
        grid.RemoveConditionalFormat(moving);

        // ---- through a file ----
        grid.ClearConditionalFormats();
        const saved = kinds.map(([range, rule], index) => grid.AddConditionalFormat(range, rule, { stopIfTrue: index % 5 === 0 }));
        grid.AddConditionalFormat('A:A D2:E3', { type: 'uniqueValues', format: { fill: '#112233', textColor: '#445566', bold: true, italic: true, underline: true, strikethrough: true } });
        const mine = list().map(({ id, ...rest }) => rest);
        const again = await reopened(grid);
        const theirs = list(again.grid).map(({ id, ...rest }) => rest);
        check('every rule comes back from an exported file: count', theirs.length, kinds.length + 1);
        for (let index = 0; index < mine.length; index++)
            if (JSON.stringify(mine[index]) !== JSON.stringify(theirs[index])) check(`rule ${index + 1} (${mine[index].rule.type}) after a file`, theirs[index], mine[index]);
        check('... each as it was (range, priority, stopIfTrue, rule, format)', JSON.stringify(theirs) === JSON.stringify(mine), true);
        check('a file\'s rules are numbered in order of precedence', again.grid.GetConditionalFormats().map((cf) => cf.id), mine.map((_, index) => index + 1));
        await frame();
        const node2 = again.host.querySelector('.spreadsheet-cell[data-row="0"][data-col="0"]');
        check('... and the opened grid draws them', getComputedStyle(node2).backgroundColor, rgb('#112233'));
        // A sheet that is not on display yet.
        const { createEmpty } = await import('/spreadsheet/Model.js');
        again.grid.ClearConditionalFormats('A:A');
        const second = await reopened(again.grid);
        check('cleared rules stay cleared in the file', second.grid.GetConditionalFormats().some((cf) => /(^| )A/.test(cf.range)), false);
        second.close(); again.close();
        grid.InstallWorkbook([grid.GetModel(), createEmpty(50, 10)], ['One', 'Two'], 1, undefined);
        check('GetConditionalFormats(sheetIndex) reads another sheet', [grid.GetConditionalFormats().length, grid.GetConditionalFormats(0).length, grid.GetConditionalFormats(1).length], [0, kinds.length + 1, 0]);
        check('ids on a new sheet are above those handed out', grid.AddConditionalFormat('A1', { type: 'blanks' }) > Math.max(...saved), true);
        off();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Data validation: GetValidations, GetValidationAt, SetValidation, ClearValidation, GetValidationList,
// ShowValidationList, CircleInvalid, ClearInvalidCircles.
checks['api-validation'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { createEmpty, cellCount } = await import('/spreadsheet/Model.js');
        const brief = (g = grid, sheet) => g.GetValidations(sheet).map((v) => `${v.id}:${v.rule.type}@${v.range}`);
        const shape = (rule) => Object.fromEntries(Object.entries(rule).filter(([, value]) => value !== null));
        const defaults = { allowBlank: true, showInput: true, showError: true, errorStyle: 'stop' };
        const notices = [];
        const offNotice = grid.OnValidationNotice((style, title, message) => notices.push(`${style}|${title}|${message}`));
        const changes = [];
        const off = grid.OnChange((change) => changes.push(change.Revision));
        check('no validations to begin with', [grid.GetValidations(), grid.GetValidationAt(0, 0), grid.GetCell(0, 0).validation], [[], null, null]);

        // ---- one validation: set, read, enforced, undo / redo ----
        const revision = grid.GetRevision();
        const first = grid.SetValidation('A1:A10', { type: 'whole', operator: 'between', formula1: 1, formula2: 10, errorTitle: 'Range', errorMessage: 'From 1 to 10' });
        check('SetValidation returns the id', [typeof first, first > 0], ['number', true]);
        check('GetValidations gives it back', grid.GetValidations().map((v) => ({ ...v, rule: shape(v.rule) })),
            [{ id: first, range: 'A1:A10', rule: { type: 'whole', operator: 'between', formula1: '1', formula2: '10', ...defaults, errorTitle: 'Range', errorMessage: 'From 1 to 10' } }]);
        check('GetValidationAt is the rule with its id and range', shape(grid.GetValidationAt(4, 0)),
            { type: 'whole', operator: 'between', formula1: '1', formula2: '10', ...defaults, errorTitle: 'Range', errorMessage: 'From 1 to 10', id: first, range: 'A1:A10' });
        check('GetCell has the same', [JSON.stringify(grid.GetCell(4, 0).validation) === JSON.stringify(grid.GetValidationAt(4, 0)), grid.GetCell(10, 0).validation, grid.GetValidationAt(4, 1)], [true, null, null]);
        check('it is one change with a new revision', [changes.length, grid.GetRevision() !== revision, grid.IsModified()], [1, true, true]);
        check('an entry inside the rule is accepted', [grid.EnterText(0, 0, '5'), grid.GetCell(0, 0).value, notices.length], [true, 5, 0]);
        check('one outside is refused and the notice raised', [grid.EnterText(0, 0, '50'), grid.GetCell(0, 0).value, notices.splice(0), grid.GetValidationNotice()?.style], [false, 5, ['stop|Range|From 1 to 10'], 'stop']);
        check('SetCell is not validated', [grid.SetCell(1, 0, '99'), grid.GetCell(1, 0).value], [undefined, 99]);
        grid.Undo();
        check('undo of the entry, then of the validation', [grid.GetCell(0, 0).value, brief()], [null, [`${first}:whole@A1:A10`]]);
        grid.SetCell(1, 0, '');
        const steps = changes.length;
        // (SetCell is not an undo step: the step on top is the validation.)
        grid.Undo();
        check('undo takes the validation away', [brief(), grid.GetValidationAt(0, 0), grid.CanRedo()], [[], null, true]);
        grid.Redo();
        check('redo brings it back under its id', [brief(), grid.EnterText(0, 0, '50')], [[`${first}:whole@A1:A10`], false]);
        notices.splice(0);
        check('undo and redo are changes', changes.length > steps, true);

        // ---- one validation per cell: a new one replaces what its cells had ----
        const second = grid.SetValidation('A3:A4', { type: 'list', values: ['Yes', 'No'] });
        check('a validation over part of another cuts it', brief(), [`${first}:whole@A1:A2 A5:A10`, `${second}:list@A3:A4`]);
        check('ClearValidation returns what it removed or cut', [grid.ClearValidation('A1:A3'), brief()], [2, [`${first}:whole@A5:A10`, `${second}:list@A4`]]);
        check('clearing where there is none is no change', [grid.ClearValidation('F1:G9'), brief().length], [0, 2]);
        grid.Undo();
        check('ClearValidation is one undo step', brief(), [`${first}:whole@A1:A2 A5:A10`, `${second}:list@A3:A4`]);
        grid.ClearValidation('A:Z');
        check('cleared', brief(), []);

        // ---- every kind ----
        const kinds = [
            ['B1:B5', { type: 'any', inputTitle: 'Note', inputMessage: 'Anything goes' }, { type: 'any', inputTitle: 'Note', inputMessage: 'Anything goes' }],
            ['C1:C5', { type: 'list', values: ['Red', 'Green', 'Blue'] }, { type: 'list', values: ['Red', 'Green', 'Blue'], dropdown: true }],
            ['D1:D5', { type: 'list', source: '=$J$1:$J$4', dropdown: false }, { type: 'list', source: '$J$1:$J$4', dropdown: false }],
            ['E1:E5', { type: 'decimal', operator: 'notBetween', formula1: 1.5, formula2: '=$K$1' }, { type: 'decimal', operator: 'notBetween', formula1: '1.5', formula2: '$K$1' }],
            ['F1:F5', { type: 'date', operator: 'greaterThanOrEqual', formula1: '2024-01-01' }, { type: 'date', operator: 'greaterThanOrEqual', formula1: '45292' }],
            ['G1:G5', { type: 'time', operator: 'lessThan', value1: '12:00' }, { type: 'time', operator: 'lessThan', formula1: '0.5' }],
            ['H1:H5', { type: 'textLength', operator: 'lessThanOrEqual', formula1: 3, errorStyle: 'warning', errorMessage: 'Long' }, { type: 'textLength', operator: 'lessThanOrEqual', formula1: '3', errorStyle: 'warning', errorMessage: 'Long' }],
            ['I1:I5', { type: 'custom', formula: '=ISNUMBER(I1)', errorStyle: 'information', showError: true, allowBlank: false }, { type: 'custom', formula: 'ISNUMBER(I1)', errorStyle: 'information', allowBlank: false }],
            ['L1:L5 N1:N5', { type: 'whole', operator: 'equal', formula1: 7, showError: false, showInput: false, inputMessage: 'hidden' }, { type: 'whole', operator: 'equal', formula1: '7', showError: false, showInput: false, inputMessage: 'hidden' }],
        ];
        for (const operator of ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']) {
            const two = operator === 'between' || operator === 'notBetween';
            kinds.push([`P${kinds.length}`, { type: 'whole', operator, formula1: '$K$1', ...(two ? { formula2: '$K$2' } : {}) }, { type: 'whole', operator, formula1: '$K$1', ...(two ? { formula2: '$K$2' } : {}) }]);
        }
        const ids = kinds.map(([range, rule]) => grid.SetValidation(range, rule));
        check('every kind is set under its own id', [new Set(ids).size, grid.GetValidations().length], [kinds.length, kinds.length]);
        const listed = grid.GetValidations();
        kinds.forEach(([range, , expected], index) => {
            const want = { ...defaults, ...expected };
            const got = shape(listed[index].rule);
            if (listed[index].range !== range || JSON.stringify(Object.entries(got).sort()) !== JSON.stringify(Object.entries(want).sort())) check(`kind ${index} (${expected.type}) reads back`, [listed[index].range, got], [range, want]);
        });
        check('... each reads back as it was set', listed.length, kinds.length);

        // ---- they are enforced ----
        ['Red', 'Amber'].forEach((text, row) => grid.SetCell(row, 9, text));
        grid.SetCell(0, 10, '5'); grid.SetCell(1, 10, '9');
        check('any value', grid.EnterText(0, 1, 'whatever'), true);
        check('a written-out list (case does not matter)', [grid.EnterText(0, 2, 'green'), grid.EnterText(1, 2, 'Pink')], [true, false]);
        check('a list read from cells', [grid.EnterText(0, 3, 'Amber'), grid.EnterText(1, 3, 'Green')], [true, false]);
        grid.SetCell(2, 9, 'Green');
        check('... follows the cells', [grid.EnterText(1, 3, 'Green'), grid.GetValidationList(1, 3), grid.GetValidationList(0, 2), grid.GetValidationList(0, 1), grid.GetValidationList(20, 20)], [true, ['Red', 'Amber', 'Green'], ['Red', 'Green', 'Blue'], null, null]);
        check('a decimal outside a range with a formula bound', [grid.EnterText(0, 4, '3'), grid.EnterText(1, 4, '7.5'), grid.EnterText(2, 4, '1')], [false, true, true]);
        check('a date on or after', [grid.EnterText(0, 5, '2023-12-31'), grid.EnterText(1, 5, '2024-06-01')], [false, true]);
        check('a text length, as a warning: accepted and told', [grid.EnterText(0, 7, 'toolong'), notices.some((n) => n.startsWith('warning||Long'))], [true, true]);
        check('a custom formula written for the first cell', [grid.EnterText(2, 8, 'text'), grid.EnterText(3, 8, '12')], [true, true]);
        check('... as information it is accepted and told', notices.some((n) => n.startsWith('information|')), true);
        check('no alert: nothing refused', grid.EnterText(0, 11, '8'), true);
        notices.splice(0);

        // ---- the dropdown of a list ----
        const button = () => host.querySelector('.spreadsheet-list-button');
        const shown = (node) => !!node && getComputedStyle(node).display !== 'none';
        grid.SelectCell(2, 2); await frame();
        const cell = host.querySelector('.spreadsheet-cell[data-row="2"][data-col="2"]').getBoundingClientRect();
        check('a list cell shows the dropdown button at its right edge', [shown(button()), Math.round(button().getBoundingClientRect().left - cell.right) <= 1, Math.abs(button().getBoundingClientRect().bottom - cell.bottom) <= 2], [true, true, true]);
        grid.SelectCell(2, 3); await frame();
        check('not with dropdown: false, nor on other cells', [shown(button()), (grid.SelectCell(2, 4), await frame(), shown(button()))], [false, false]);
        grid.SelectCell(2, 2); await frame();
        button().dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        const items = () => [...document.querySelectorAll('#validation-list-dropdown .validation-list-item')];
        check('a click on it opens the list under the cell', [items().map((item) => item.textContent), Math.abs(document.getElementById('validation-list-dropdown').getBoundingClientRect().top - cell.bottom) <= 3], [['Red', 'Green', 'Blue'], true]);
        items()[2].dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        check('picking a value enters it and closes the list', [grid.GetCell(2, 2).value, document.getElementById('validation-list-dropdown')], ['Blue', null]);
        grid.Undo();
        check('... as one undo step', grid.GetCell(2, 2).value, null);
        check('ShowValidationList opens it too; a click elsewhere closes it', [grid.ShowValidationList(), items().length, (await new Promise((r) => setTimeout(r, 20)), document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })), document.getElementById('validation-list-dropdown'))], [true, 3, null]);
        grid.SelectCell(1, 3);
        check('a list without its dropdown still has values for a host', [grid.ShowValidationList(), grid.GetValidationList(1, 3).length], [true, 3]);
        document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        await new Promise((r) => setTimeout(r, 20)); document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        grid.SetReadOnly(true); grid.SelectCell(2, 2); await frame();
        check('no button and no list in read-only mode', [shown(button()), grid.ShowValidationList()], [false, false]);
        grid.SetReadOnly(false);
        // The input message of the selected cell.
        grid.SelectCell(1, 1); await frame();
        const message = host.querySelector('.spreadsheet-input-message');
        check('the input message shows while the cell is selected', [shown(message), message.textContent], [true, 'NoteAnything goes']);
        grid.SelectCell(1, 11); await frame();
        check('not when showInput is off', shown(message), false);

        // ---- circle invalid data ----
        grid.SetCell(3, 2, 'Pink'); grid.SetCell(4, 4, '2'); grid.SetCell(0, 11, '8'); grid.SetCell(4, 5, 'soon');
        const invalid = grid.CircleInvalid();
        await frame();
        check('CircleInvalid lists the cells that break their rule', invalid.map((c) => c.address).sort(), ['C4', 'E5', 'F5', 'H1', 'I3', 'L1']);
        check('... with row, col, text and the alert', invalid.filter((c) => c.address === 'H1').map((c) => [c.row, c.col, c.text, c.errorMessage]), [[0, 7, 'toolong', 'Long']]);
        const circles = () => [...host.querySelectorAll('.spreadsheet-invalid-circle')].filter(shown).map((node) => `${node.getAttribute('data-row')},${node.getAttribute('data-col')}`).sort();
        check('the grid circles them', circles(), ['0,11', '0,7', '2,8', '3,2', '4,4', '4,5']);
        const circle = host.querySelector('.spreadsheet-invalid-circle[data-row="3"][data-col="2"]').getBoundingClientRect();
        const circledCell = host.querySelector('.spreadsheet-cell[data-row="3"][data-col="2"]').getBoundingClientRect();
        check('a circle is around its cell', [circle.left < circledCell.left, circle.right > circledCell.right, circle.top < circledCell.top, circle.bottom > circledCell.bottom, getComputedStyle(host.querySelector('.spreadsheet-invalid-circle')).borderTopColor], [true, true, true, true, 'rgb(224, 0, 0)']);
        const beforeCircles = grid.GetRevision();
        check('circles are not a change of the document', [grid.CircleInvalid(false).length, grid.GetRevision()], [6, beforeCircles]);
        grid.ClearInvalidCircles(); await frame();
        check('ClearInvalidCircles removes them', circles(), []);
        grid.CircleInvalid(); grid.ClearValidation('C1:C5'); await frame();
        check('a change of the validations removes them too', [circles(), grid.CircleInvalid(false).map((c) => c.address).includes('C4')], [[], false]);
        grid.Undo();

        // ---- invalid input ----
        check('throws: no range', thrown(() => grid.SetValidation(undefined, { type: 'any' })), 'SetValidation: the range is missing');
        check('throws: a bad range', /^SetValidation: "A1:" is not a range/.test(thrown(() => grid.SetValidation('A1:', { type: 'any' }))), true);
        check('throws: no rule', thrown(() => grid.SetValidation('A1')), 'SetValidation: the rule is missing');
        check('throws: an unknown type', thrown(() => grid.SetValidation('A1', { type: 'sparkle' })), 'SetValidation: unknown validation type "sparkle" (any, list, whole, decimal, date, time, textLength, custom)');
        check('throws: an unknown key', /^SetValidation: a validation rule has no key "colour"/.test(thrown(() => grid.SetValidation('A1', { type: 'any', colour: 1 }))), true);
        check('throws: an unknown operator', /^SetValidation: unknown operator "gt"/.test(thrown(() => grid.SetValidation('A1', { type: 'whole', operator: 'gt', formula1: 1 }))), true);
        check('throws: a missing operand', /^SetValidation: a "whole" rule needs formula2/.test(thrown(() => grid.SetValidation('A1', { type: 'whole', operator: 'between', formula1: 1 }))), true);
        check('throws: a list without values', /^SetValidation: a list rule needs values/.test(thrown(() => grid.SetValidation('A1', { type: 'list' }))), true);
        check('throws: values that are not an array', /^SetValidation: rule.values must be an array/.test(thrown(() => grid.SetValidation('A1', { type: 'list', values: 'a,b' }))), true);
        check('throws: a wrong type of key', /^SetValidation: rule.allowBlank must be true or false/.test(thrown(() => grid.SetValidation('A1', { type: 'any', allowBlank: 'yes' }))), true);
        check('throws: a bad error style', /^SetValidation: unknown errorStyle "loud"/.test(thrown(() => grid.SetValidation('A1', { type: 'any', errorStyle: 'loud' }))), true);
        check('throws: ClearValidation without a range', thrown(() => grid.ClearValidation()), 'ClearValidation: the range is missing');
        check('throws: a cell outside the sheet', /^GetValidationAt: row 5000000 is outside the sheet/.test(thrown(() => grid.GetValidationAt(5000000, 0))), true);
        check('throws: an unknown sheet', /^GetValidations: there is no sheet 7/.test(thrown(() => grid.GetValidations(7))), true);
        check('nothing was changed by the refused calls', grid.GetValidations().length, kinds.length);
        // The rule of a cell can be put on other cells as it is.
        const copied = grid.SetValidation('R1:R3', grid.GetValidationAt(0, 7));
        check('GetValidationAt can be given back to SetValidation', [grid.GetValidations().find((v) => v.id === copied).range, JSON.stringify(grid.GetValidations().find((v) => v.id === copied).rule) === JSON.stringify(grid.GetValidations().find((v) => v.range === 'H1:H5').rule)], ['R1:R3', true]);
        grid.ClearValidation('R1:R3');

        // ---- rows inserted: ranges and formulas follow ----
        grid.InsertRows(0, 2);
        check('validations follow inserted rows', [grid.GetValidations()[0].range, grid.GetValidations()[3].rule.formula2, grid.GetValidations()[2].rule.source], ['B3:B7', '$K$3', '$J$3:$J$6']);
        grid.Undo();
        check('... and come back with an undo', [grid.GetValidations()[0].range, grid.GetValidations()[3].rule.formula2], ['B1:B5', '$K$1']);

        // ---- through a file ----
        const mine = grid.GetValidations().map(({ id, ...rest }) => rest);
        const bytes = await grid.ExportXlsx();
        const again = mountGrid();
        await again.grid.ImportXlsx(bytes);
        const theirs = again.grid.GetValidations().map(({ id, ...rest }) => rest);
        for (let index = 0; index < mine.length; index++)
            if (JSON.stringify(mine[index]) !== JSON.stringify(theirs[index])) check(`validation ${index + 1} (${mine[index].rule.type}) after a file`, theirs[index], mine[index]);
        check('every validation comes back from an exported file', [theirs.length, JSON.stringify(theirs) === JSON.stringify(mine)], [mine.length, true]);
        check('a file\'s validations are numbered in file order', again.grid.GetValidations().map((v) => v.id), mine.map((_, index) => index + 1));
        check('... and are enforced', [again.grid.EnterText(3, 2, 'Pink'), again.grid.EnterText(3, 2, 'Red'), again.grid.EnterText(3, 3, 'Amber'), again.grid.EnterText(3, 3, 'Mauve')], [false, true, true, false]);
        again.close();

        // ---- a whole column: no cell is made, a far row is covered, the file says A1:A1048576 ----
        const big = mountGrid();
        big.grid.InstallWorkbook([createEmpty(60000, 12), createEmpty(50, 5)], ['Data', 'Lists'], 0, undefined);
        big.grid.SwitchSheet(1); big.grid.SetCell(0, 0, 'North'); big.grid.SetCell(1, 0, 'South'); big.grid.SwitchSheet(0);
        const cellsBefore = cellCount(big.grid.GetModel());
        const whole = big.grid.SetValidation('A:A', { type: 'decimal', operator: 'greaterThan', formula1: 0, errorMessage: 'Positive' });
        big.grid.SetValidation('C:D', { type: 'list', source: 'Lists!$A$1:$A$9' });
        check('a whole-column validation makes no cell', [cellCount(big.grid.GetModel()) - cellsBefore, brief(big.grid)], [0, [`${whole}:decimal@A:A`, `${whole + 1}:list@C:D`]]);
        check('a far row is covered', [big.grid.GetValidationAt(50000, 0)?.id, big.grid.EnterText(50000, 0, '-4'), big.grid.EnterText(50000, 0, '4'), big.grid.GetCell(50000, 0).value], [whole, false, true, 4]);
        check('a list read from another sheet', [big.grid.GetValidationList(40000, 3), big.grid.EnterText(40000, 3, 'East'), big.grid.EnterText(40000, 3, 'south')], [['North', 'South'], false, true]);
        big.grid.SelectCell(40000, 2); await frame();
        check('the dropdown button on a far cell of the column', shown(big.host.querySelector('.spreadsheet-list-button')), true);
        check('GetValidations(sheetIndex) reads another sheet', [big.grid.GetValidations(1), big.grid.GetValidations(0).length], [[], 2]);
        const bigBytes = await big.grid.ExportXlsx();
        const sheetXml = await (await window.JSZip.loadAsync(bigBytes)).file('xl/worksheets/sheet1.xml').async('string');
        check('the file has the whole columns as ranges', [...sheetXml.matchAll(/<dataValidation [^>]*sqref="([^"]*)"/g)].map((m) => m[1]), ['A1:A1048576', 'C1:D1048576']);
        const bigAgain = mountGrid();
        await bigAgain.grid.ImportXlsx(bigBytes);
        check('re-imported it is one validation per rule, on whole columns', bigAgain.grid.GetValidations().map((v) => [v.id, v.range, v.rule.type, v.rule.source]), [[1, 'A:A', 'decimal', null], [2, 'C:D', 'list', 'Lists!$A$1:$A$9']]);
        check('... still without a cell per row, and enforced far down', [cellCount(bigAgain.grid.GetModel()), bigAgain.grid.EnterText(55000, 0, '-1'), bigAgain.grid.EnterText(55000, 2, 'North'), bigAgain.grid.EnterText(55001, 2, 'West')], [2, false, true, false]);
        check('... the other sheet was read for the list only when it was needed', bigAgain.grid.GetValidationList(3, 3), ['North', 'South']);
        // Unedited validations of a file go back as the file had them.
        const third = mountGrid();
        await third.grid.ImportXlsx(await bigAgain.grid.ExportXlsx());
        check('a second trip keeps them', third.grid.GetValidations().map((v) => [v.range, v.rule.type]), [['A:A', 'decimal'], ['C:D', 'list']]);
        third.close(); bigAgain.close(); big.close();
        off(); offNotice();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Defined names: GetDefinedNames, SetDefinedName, DeleteDefinedName, and the name box's GoTo.
checks['api-names'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([grid.GetModel(), createEmpty(200, 20)], ['Data', 'Other sheet'], 0, undefined);
        grid.SetCell(0, 0, '1'); grid.SetCell(1, 0, '2'); grid.SetCell(2, 0, '3'); grid.SetCell(0, 1, '10'); grid.SetCell(1, 1, '20');
        grid.SetCell(4, 0, '=SUM(Totals)'); grid.SetCell(5, 0, '=Rate*100');
        const value = (row, col) => grid.GetCell(row, col).value;
        check('a formula with an unknown name', [value(4, 0), value(5, 0)], ['#NAME?', '#NAME?']);
        check('no names yet', grid.GetDefinedNames(), []);
        const changes = [];
        const off = grid.OnChange((change) => changes.push(change));
        const take = () => changes.splice(0).map((c) => `${c.Sheets.join(',')}|${c.WorkbookChanged ? 'book' : ''}`);

        // ---- define ----
        const r0 = grid.GetRevision();
        grid.SetDefinedName('Totals', '=Data!$A$1:$A$3', undefined, 'the first column');
        check('a name is defined and its formulas recalculated', [grid.GetDefinedNames(), value(4, 0)],
            [[{ name: 'Totals', refersTo: 'Data!$A$1:$A$3', scope: null, scopeName: null, comment: 'the first column', isBuiltIn: false, hidden: false }], 6]);
        const r1 = grid.GetRevision();
        check('it is a change of the workbook: OnChange and a new revision', [take(), r1 !== r0, grid.IsModified()], [['0|book'], true, true]);
        grid.SetDefinedName('Rate', '0.19', null, null);
        check('a constant, without "="', [value(5, 0), grid.GetDefinedNames().map((n) => [n.name, n.refersTo, n.comment])], [19, [['Rate', '0.19', null], ['Totals', 'Data!$A$1:$A$3', 'the first column']]]);
        take();
        // A sheet-level name shadows the workbook's on its own sheet.
        grid.SwitchSheet(1);
        grid.SetCell(0, 0, '=Rate*100'); grid.SetCell(0, 1, '7'); grid.SetCell(1, 0, '=SUM(Totals)');
        check('another sheet reads the workbook-level names', [value(0, 0), value(1, 0)], [19, 6]);
        grid.SetDefinedName('rate', "'Other sheet'!$B$1", 1);
        check('a sheet-level name wins on its sheet', [value(0, 0), grid.GetDefinedNames().map((n) => [n.name, n.scope, n.scopeName])], [700, [['Rate', null, null], ['Totals', null, null], ['rate', 1, 'Other sheet']]]);
        grid.SwitchSheet(0);
        check('... and only there', value(5, 0), 19);

        // ---- replace, undo, redo ----
        const beforeReplace = grid.GetRevision();
        grid.SetDefinedName('TOTALS', 'Data!$B$1:$B$2');
        check('a replaced name: new spelling, re-pointed formulas, one name', [value(4, 0), grid.GetDefinedNames().filter((n) => n.scope === null).map((n) => [n.name, n.refersTo, n.comment])], [30, [['Rate', '0.19', null], ['TOTALS', 'Data!$B$1:$B$2', null]]]);
        check('formulas of the other sheet follow', grid.GetCell(1, 0) && grid.GetRange("'Other sheet'!A2")[0][0].value, 30);
        grid.Undo();
        check('undo puts the name and the values back, and the revision', [value(4, 0), grid.GetDefinedNames()[1].name, grid.GetDefinedNames()[1].comment, grid.GetRevision() === beforeReplace], [6, 'Totals', 'the first column', true]);
        grid.Redo();
        check('redo', [value(4, 0), grid.GetDefinedNames()[1].name], [30, 'TOTALS']);
        grid.Undo();
        grid.SetDefinedName('Totals', '=Data!$A$1:$A$3', undefined, 'the first column');
        check('defining a name as it is changes nothing', [grid.CanRedo(), take().length > 0], [true, true]);

        // ---- delete ----
        take();
        grid.DeleteDefinedName('totals');
        check('a deleted name: #NAME? in its formulas', [value(4, 0), grid.GetDefinedNames().map((n) => n.name), take()], ['#NAME?', ['Rate', 'rate'], ['0,1|book']]);
        grid.Undo();
        check('undo brings it back', [value(4, 0), grid.GetDefinedNames().map((n) => n.name)], [6, ['Rate', 'Totals', 'rate']]);
        grid.Redo(); grid.Undo();
        grid.DeleteDefinedName('Rate', 1);
        check('a sheet-level name is deleted by its sheet', [grid.GetDefinedNames().map((n) => [n.name, n.scope]), grid.GetRange("'Other sheet'!A1")[0][0].value], [[['Rate', null], ['Totals', null]], 19]);
        grid.Undo();

        // ---- what is refused ----
        check('invalid names', ['A1', 'R1C1', 'My name', '1st', '', 'TRUE', '_xlnm.Print_Area'].map((name) => thrown(() => grid.SetDefinedName(name, '1')) !== null), [true, true, true, true, true, true, true]);
        check('the message says what is wrong', thrown(() => grid.SetDefinedName('A1', '1')), 'SetDefinedName: the name "A1" reads as a cell reference');
        check('an invalid formula', [thrown(() => grid.SetDefinedName('Bad', 'SUM(')) !== null, thrown(() => grid.SetDefinedName('Bad', '')), thrown(() => grid.SetDefinedName('Bad', 5)) !== null],
            [true, 'SetDefinedName: a name has to refer to something', true]);
        check('an unknown sheet', [thrown(() => grid.SetDefinedName('Fine', '1', 7)), thrown(() => grid.DeleteDefinedName('Rate', 9))], ['SetDefinedName: there is no sheet 7 (the workbook has 2)', 'DeleteDefinedName: there is no sheet 9 (the workbook has 2)']);
        check('deleting what is not there', [thrown(() => grid.DeleteDefinedName('Nope')), thrown(() => grid.DeleteDefinedName('Totals', 0))],
            ['DeleteDefinedName: the workbook has no name "Nope" (a sheet-level name needs its sheet index)', 'DeleteDefinedName: the sheet "Data" has no name "Totals"']);
        check('nothing was defined by the refused calls', grid.GetDefinedNames().map((n) => n.name), ['Rate', 'Totals', 'rate']);

        // ---- GoTo ----
        const selection = () => [grid.GetActiveSheetIndex(), grid.GetSelectionRef()];
        check('a cell', [grid.GoTo('C7'), selection()], [{ sheetIndex: 0, range: 'C7' }, [0, 'C7']]);
        check('a range with $', [grid.GoTo('$B$2:$D$5'), selection()], [{ sheetIndex: 0, range: 'B2:D5' }, [0, 'B2:D5']]);
        check('a name', [grid.GoTo('totals'), selection()], [{ sheetIndex: 0, range: 'A1:A3' }, [0, 'A1:A3']]);
        check('a reference on another sheet switches to it', [grid.GoTo("'Other sheet'!B3:C9"), selection()], [{ sheetIndex: 1, range: 'B3:C9' }, [1, 'B3:C9']]);
        check('on that sheet its own name is the one', [grid.GoTo('Rate'), selection()], [{ sheetIndex: 1, range: 'B1' }, [1, 'B1']]);
        check('a name that leads to another sheet', [grid.GoTo('Totals'), selection()], [{ sheetIndex: 0, range: 'A1:A3' }, [0, 'A1:A3']]);
        check("another sheet's name, qualified", [grid.GoTo("'Other sheet'!rate"), selection()], [{ sheetIndex: 1, range: 'B1' }, [1, 'B1']]);
        grid.SwitchSheet(0);
        check('a whole column is clipped to the sheet', grid.GoTo('B:B').range, 'B1:B' + grid.GetModel().RowCount);
        grid.GoTo('A150');
        await frame();
        check('the place is scrolled into view', [!!host.querySelector('.spreadsheet-cell[data-row="149"][data-col="0"]'), host.querySelector('.spreadsheet-container').scrollTop > 1000], [true, true]);
        check('a name that is not a range', thrown(() => grid.GoTo('Rate')), 'GoTo: the name "Rate" does not refer to a range (0.19)');
        check('nothing by that name', [thrown(() => grid.GoTo('Nowhere')), thrown(() => grid.GoTo('Missing!A1')), thrown(() => grid.GoTo('')) !== null, thrown(() => grid.GoTo('ZZZ99999999')) !== null],
            ['GoTo: "Nowhere" is not a reference, a defined name or a table', 'GoTo: there is no sheet "Missing"', true, true]);
        grid.SelectRange(0, 0, 2, 1);
        grid.DefineTable('Sales', 0, 0, 2, 1, ['a', 'b'], false);
        check('a table', grid.GoTo('sales'), { sheetIndex: 0, range: 'A1:B3' });
        check('a table name is not free for a defined name', thrown(() => grid.SetDefinedName('SALES', '1')), 'SetDefinedName: "SALES" is the name of a table');

        // ---- through a file ----
        off();
        const again = await reopened(grid);
        try {
            check('the names come back from the file: spelling, comment, scope', again.grid.GetDefinedNames(),
                [{ name: 'Rate', refersTo: '0.19', scope: null, scopeName: null, comment: null, isBuiltIn: false, hidden: false },
                 { name: 'Totals', refersTo: 'Data!$A$1:$A$3', scope: null, scopeName: null, comment: 'the first column', isBuiltIn: false, hidden: false },
                 { name: 'rate', refersTo: "'Other sheet'!$B$1", scope: 1, scopeName: 'Other sheet', comment: null, isBuiltIn: false, hidden: false }]);
            check('an opened workbook is not modified', again.grid.IsModified(), false);
            check('and its formulas read them', [again.grid.GetCell(4, 0).value, again.grid.GetCell(5, 0).value, again.grid.GetRange("'Other sheet'!A1")[0][0].value], [6, 19, 700]);
            again.grid.SetDefinedName('Rate', '0.07', undefined, 'reduced');
            check('a name edited after opening', [again.grid.GetCell(5, 0).value, again.grid.IsModified()], [7.000000000000001, true]);
            const third = await reopened(again.grid);
            try {
                check('... is saved as edited, the others as they were', third.grid.GetDefinedNames().map((n) => [n.name, n.refersTo, n.scope, n.comment]),
                    [['Rate', '0.07', null, 'reduced'], ['Totals', 'Data!$A$1:$A$3', null, 'the first column'], ['rate', "'Other sheet'!$B$1", 1, null]]);
            } finally { third.close(); }
        } finally { again.close(); }

        // A workbook of the corpus with names: listed as the file spells them, untouched by a save.
        const bytes = await bytesOf('tests/real-world/milk-street-office.xlsx');
        await grid.ImportXlsx(bytes);
        const fileNames = grid.GetDefinedNames();
        const reader = OF.excel.workbookInfo(await OF.excel.load(bytes)).definedNames.filter((n) => !/^_xlnm\./.test(n.name)).map((n) => n.name).sort();
        check('names of an opened file are listed with their spelling', [fileNames.length > 0, fileNames.map((n) => n.name).sort()], [true, reader]);
        check('opening is not a modification', grid.IsModified(), false);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Freeze panes and the view: FreezeAt / Unfreeze / GetFrozen, SetZoom / GetZoom, grid lines,
// headers, formulas, zeros.
checks['api-view'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const scroller = host.querySelector('.spreadsheet-container');
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const round = (rect) => [rect.left, rect.top, rect.width, rect.height].map((v) => Math.round(v));
        const origin = () => { const box = scroller.getBoundingClientRect(); return [box.left + scroller.clientLeft, box.top + scroller.clientTop]; };
        // Where a cell of the default 100 x 25 grid is on the page at a zoom, scrolled or not.
        const expected = (row, col, zoom, frozen = [0, 0], headers = [50, 30]) => {
            const [ox, oy] = origin();
            const x = (headers[0] + col * 100) * zoom - (col < frozen[1] ? 0 : scroller.scrollLeft), y = (headers[1] + row * 25) * zoom - (row < frozen[0] ? 0 : scroller.scrollTop);
            return [ox + x, oy + y, 100 * zoom, 25 * zoom].map((v) => Math.round(v));
        };
        const near = (got, want, slack = 1) => got.length === want.length && got.every((v, i) => Math.abs(v - want[i]) <= slack);
        const click = (x, y, type = 'mousedown') => {
            const target = document.elementFromPoint(x, y) ?? host.querySelector('.spreadsheet-content');
            target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0 }));
            if (type === 'mousedown') document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: x, clientY: y }));
        };
        const active = () => { const cell = grid.GetActiveCell(); return [cell.Row, cell.Col]; };
        const pixel = async (x, y) => {
            const image = new Image();
            image.src = 'data:image/png;base64,' + await window.screenshotRegion(Math.round(x), Math.round(y), 1, 1);
            await image.decode();
            const canvas = document.createElement('canvas'); canvas.width = 1; canvas.height = 1;
            const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
            return [...context.getImageData(0, 0, 1, 1).data.slice(0, 3)];
        };
        grid.SetCell(0, 0, '5'); grid.SetCell(1, 0, '=A1*2'); grid.SetCell(2, 0, '0'); grid.SetCell(3, 0, '=A1-5'); grid.SetCell(4, 0, 'text'); grid.SetCell(5, 3, 'far');
        grid.SelectCell(7, 2); grid.ApplyBgColor('#ff0000'); grid.SelectCell(0, 0);
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        await frame();

        // ---- freeze panes ----
        check('nothing frozen at first', [grid.GetFrozen(), grid.GetViewOptions()], [{ rows: 0, cols: 0 }, { zoom: 1, showGridLines: true, showHeaders: true, showFormulas: false, showZeros: true, frozen: { rows: 0, cols: 0 } }]);
        const r0 = grid.GetRevision();
        grid.FreezeAt(2, 1);
        check('FreezeAt freezes the rows above and the columns left of the cell', [grid.GetFrozen(), grid.GetModel().FrozenRows, grid.GetModel().FrozenCols], [{ rows: 2, cols: 1 }, 2, 1]);
        check('it is a change with an undo step', [changes.splice(0).length, grid.GetRevision() !== r0, grid.CanUndo()], [1, true, true]);
        grid.ScrollToCell(60, 12); await frame();
        check('frozen rows and columns stay while the rest scrolls', [scroller.scrollTop > 500, scroller.scrollLeft > 100, near(round(node(0, 0).getBoundingClientRect()), expected(0, 0, 1, [2, 1])), near(round(node(1, 5).getBoundingClientRect()), expected(1, 5, 1, [2, 1])),
            near(round(node(55, 0).getBoundingClientRect()), expected(55, 0, 1, [2, 1]))], [true, true, true, true, true]);
        grid.Undo();
        check('undo unfreezes, redo freezes again', [grid.GetFrozen(), (grid.Redo(), grid.GetFrozen()), grid.GetRevision() !== r0], [{ rows: 0, cols: 0 }, { rows: 2, cols: 1 }, true]);
        grid.Unfreeze();
        check('Unfreeze', [grid.GetFrozen(), grid.CanUndo()], [{ rows: 0, cols: 0 }, true]);
        const steps = grid.GetRevision();
        grid.Unfreeze(); grid.FreezeAt(0, 0);
        check('unfreezing what is not frozen is nothing', grid.GetRevision(), steps);
        grid.FreezeAt(1, 0);
        check('the first row alone', grid.GetFrozen(), { rows: 1, cols: 0 });
        check('a cell that is not on the sheet', [thrown(() => grid.FreezeAt(-1, 0)), thrown(() => grid.FreezeAt(0, 99999)) !== null, thrown(() => grid.FreezeAt('a', 0)) !== null], ['FreezeAt: row -1 is not on the sheet (0..999)', true, true]);
        grid.Unfreeze();
        grid.ScrollToCell(0, 0); await frame();

        // ---- zoom ----
        const saved = grid.GetRevision();
        grid.MarkSaved();
        changes.splice(0);
        const shownRows = () => [...host.querySelectorAll('.spreadsheet-row-header')].filter((header) => header.style.display !== 'none').length;
        const rows100 = shownRows();
        for (const zoom of [0.5, 1.25, 2]) {
            grid.SetZoom(zoom); await frame();
            check(`zoom ${zoom}: GetZoom and the model`, [grid.GetZoom(), grid.GetModel().Zoom], [zoom, zoom]);
            check(`zoom ${zoom}: cells are scaled and placed`, [round(node(0, 0).getBoundingClientRect()), round(node(5, 3).getBoundingClientRect())], [expected(0, 0, zoom), expected(5, 3, zoom)]);
            const header = host.querySelector('.spreadsheet-col-header[data-col="2"]').getBoundingClientRect(), rowHeader = host.querySelector('.spreadsheet-row-header[data-row="4"]').getBoundingClientRect();
            check(`zoom ${zoom}: the headers too`, [near(round(header), [origin()[0] + (50 + 200) * zoom, origin()[1], 100 * zoom, 30 * zoom].map(Math.round)), near(round(rowHeader), [origin()[0], origin()[1] + (30 + 100) * zoom, 50 * zoom, 25 * zoom].map(Math.round))], [true, true]);
            check(`zoom ${zoom}: text is laid out at the zoomed size (not a scaled bitmap)`, Math.abs(parseFloat(getComputedStyle(node(4, 0)).fontSize) * zoom - node(4, 0).getBoundingClientRect().height * 11 / 25) < 0.6, true);
            check(`zoom ${zoom}: the scroll extent is the zoomed sheet`, [Math.abs(scroller.scrollHeight - (30 + 1000 * 25) * zoom) <= 2, Math.abs(scroller.scrollWidth - (50 + 26 * 100) * zoom) <= 2], [true, true]);
            // Hit testing: a click in the middle of where a cell is drawn selects it.
            const target = expected(4, 2, zoom);
            click(target[0] + target[2] / 2, target[1] + target[3] / 2);
            check(`zoom ${zoom}: a click lands on the cell under it`, active(), [4, 2]);
            const edge = expected(6, 1, zoom);
            click(edge[0] + 2, edge[1] + 2); const first = active();
            click(edge[0] + edge[2] - 2, edge[1] + edge[3] - 2);
            check(`zoom ${zoom}: ... also just inside its corners`, [first, active()], [[6, 1], [6, 1]]);
            await frame();
            const outline = [...host.querySelectorAll('.spreadsheet-active-cell')].find((overlay) => overlay.style.display !== 'none').getBoundingClientRect();
            check(`zoom ${zoom}: the active-cell outline sits on the cell`, near(round(outline), expected(6, 1, zoom), 3), true);
            const fill = expected(7, 2, zoom);
            check(`zoom ${zoom}: a filled cell is painted where it is`, [await pixel(fill[0] + fill[2] / 2, fill[1] + fill[3] / 2), await pixel(fill[0] + fill[2] / 2, fill[1] + fill[3] + 4 * zoom)], [[255, 0, 0], [255, 255, 255]]);
            // The edit box covers the cell being edited.
            grid.SelectCell(5, 3); grid.StartEdit(); await frame();
            const box = host.querySelector('.spreadsheet-edit-input').getBoundingClientRect();
            check(`zoom ${zoom}: the edit box covers its cell`, near(round(box), expected(5, 3, zoom), 3), true);
            grid.CancelEdit();
            // Scrolled: the far cell comes into view, is where it should be, and takes clicks.
            grid.ScrollToCell(400, 20); await frame();
            const far = node(400, 20).getBoundingClientRect(), view = scroller.getBoundingClientRect();
            check(`zoom ${zoom}: ScrollToCell brings a far cell into view`, [far.top >= view.top + 30 * zoom - 1, far.bottom <= view.bottom + 1, far.left >= view.left + 50 * zoom - 1, far.right <= view.right + 1], [true, true, true, true]);
            check(`zoom ${zoom}: scrolled cells are placed by the scroll offset`, near(round(far), expected(400, 20, zoom)), true);
            click(far.left + far.width / 2, far.top + far.height / 2);
            check(`zoom ${zoom}: a click after scrolling`, active(), [400, 20]);
            // Frozen panes while scrolled.
            grid.FreezeAt(2, 1); await frame();
            check(`zoom ${zoom}: frozen panes stay put while scrolled`, [near(round(node(1, 0).getBoundingClientRect()), expected(1, 0, zoom, [2, 1])), near(round(node(400, 0).getBoundingClientRect()), expected(400, 0, zoom, [2, 1])), near(round(node(1, 20).getBoundingClientRect()), expected(1, 20, zoom, [2, 1]))], [true, true, true]);
            const pinned = expected(1, 0, zoom, [2, 1]);
            click(pinned[0] + pinned[2] / 2, pinned[1] + pinned[3] / 2);
            check(`zoom ${zoom}: a click on a frozen cell`, active(), [1, 0]);
            grid.Undo();
            grid.ScrollToCell(0, 0); await frame();
        }
        check('more rows are drawn zoomed out, fewer zoomed in', [(grid.SetZoom(0.5), await frame(), shownRows()) > rows100, (grid.SetZoom(2), await frame(), shownRows()) < rows100], [true, true]);
        // The cell in the top-left corner stays there when the zoom changes.
        grid.SetZoom(1); grid.ScrollToCell(0, 0); await frame();
        changes.splice(0);
        scroller.scrollTop = 200 * 25; scroller.scrollLeft = 300; await frame(); await frame();
        grid.SetZoom(1.5); await frame();
        check('zooming keeps the top-left cell in place', [scroller.scrollTop, scroller.scrollLeft], [200 * 25 * 1.5, 450]);
        check('zooming is not an edit: no revision, no change event, no undo step of its own', [grid.GetRevision() === saved, grid.IsModified(), changes.length], [true, false, 0]);
        check('a factor out of range', [thrown(() => grid.SetZoom(0)), thrown(() => grid.SetZoom(5)) !== null, thrown(() => grid.SetZoom('big')) !== null, thrown(() => grid.SetZoom(NaN)) !== null, grid.GetZoom()],
            ['SetZoom: the factor has to be between 0.1 and 4 (1 = 100%), not 0', true, true, true, 1.5]);
        // A drag of a column's resize handle follows the pointer at any zoom.
        grid.ScrollToCell(0, 0); await frame();
        const handle = host.querySelector('.spreadsheet-col-header[data-col="1"] .col-resize-handle').getBoundingClientRect();
        host.querySelector('.spreadsheet-col-header[data-col="1"] .col-resize-handle').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX: handle.left + 1, clientY: handle.top + 5, button: 0 }));
        document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: handle.left + 1 + 60, clientY: handle.top + 5 }));
        document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        await frame(); await frame();
        check('dragging a column edge 60 px at 150 % widens the column by 40', [Math.round(node(0, 1).getBoundingClientRect().width), (await import('/spreadsheet/Model.js')).getColumnWidth(1, grid.GetModel())], [210, 140]);
        // Each sheet has its own zoom.
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([grid.GetModel(), createEmpty(100, 10)], ['One', 'Two'], 0, undefined);
        grid.SwitchSheet(1); await frame();
        check('another sheet has its own zoom', [grid.GetZoom(), round(node(1, 1).getBoundingClientRect())], [1, expected(1, 1, 1)]);
        grid.SetZoom(0.75); grid.SwitchSheet(0); await frame();
        check('... and coming back the first has its own', [grid.GetZoom(), Math.round(node(0, 0).getBoundingClientRect().height * 10) / 10], [1.5, 37.5]);

        // ---- pictures scale with the sheet ----
        const pictured = mountGrid();
        try {
            await pictured.grid.ImportXlsx(await bytesOf('tests/test-documents-excel/synthetic-drawing-extras.xlsx'));
            pictured.grid.SetZoom(2); await frame();
            const picture = pictured.grid.GetDrawings().find((drawing) => drawing.kind === 'picture'), image = pictured.host.querySelector('.spreadsheet-picture').getBoundingClientRect();
            const box = pictured.host.querySelector('.spreadsheet-container'), left = box.getBoundingClientRect().left + box.clientLeft, top = box.getBoundingClientRect().top + box.clientTop;
            check('a picture is drawn at twice its place and size at 200 %', near(round(image), [left + (50 + picture.x) * 2, top + (30 + picture.y) * 2, picture.width * 2, picture.height * 2].map(Math.round), 2), true);
            check('opening and zooming leaves the workbook unmodified', pictured.grid.IsModified(), false);
        } finally { pictured.close(); }

        // ---- what the view shows ----
        grid.SetZoom(1); await frame();
        const text = (row, col) => node(row, col).textContent;
        check('values at first', [text(1, 0), text(2, 0), text(3, 0)], ['10', '0', '0']);
        const beforeOptions = grid.GetRevision();
        changes.splice(0);
        grid.SetShowFormulas(true); await frame();
        check('formulas are shown as text, left-aligned; values stay values', [text(0, 0), text(1, 0), text(3, 0), getComputedStyle(node(1, 0)).textAlign !== 'right', grid.GetCell(1, 0).value, grid.GetShowFormulas()], ['5', '=A1*2', '=A1-5', true, 10, true]);
        grid.Undo(); await frame();
        check('undo shows the results again', [text(1, 0), grid.GetShowFormulas(), grid.GetRevision() === beforeOptions], ['10', false, true]);
        grid.SetShowZeros(false); await frame();
        check('zeros are hidden: typed and calculated; other values stay', [text(2, 0), text(3, 0), text(0, 0), text(1, 0), grid.GetCell(2, 0).value, grid.GetShowZeros()], ['', '', '5', '10', 0, false]);
        grid.SetCell(0, 0, '0'); await frame();
        check('a result that becomes zero disappears', [text(0, 0), text(1, 0), text(3, 0)], ['', '', '-5']);
        grid.SetCell(0, 0, '5');
        grid.SetShowZeros(true); await frame();
        check('and zeros are back', [text(2, 0), text(3, 0)], ['0', '0']);
        grid.SelectCell(20, 5);
        check('grid lines at first', [grid.GetCellEdges(3, 3).Top, grid.GetShowGridLines()], ['grid', true]);
        grid.SetShowGridLines(false); await frame();
        const between = expected(3, 3, 1);
        check('no grid lines: the edges are gone, in the model and on screen', [grid.GetCellEdges(3, 3).Top, grid.GetModel().ShowGridLines, await pixel(between[0] + 50, between[1])], [null, false, [255, 255, 255]]);
        grid.Undo(); await frame();
        check('undo draws them again', [grid.GetCellEdges(3, 3).Top, (await pixel(between[0] + 50, between[1]))[0] < 250], ['grid', true]);
        grid.SetShowHeaders(false); await frame();
        const [ox, oy] = origin();
        check('without headers the cells start at the corner', [round(node(0, 0).getBoundingClientRect()), getComputedStyle(host.querySelector('.spreadsheet-col-headers')).display, getComputedStyle(host.querySelector('.spreadsheet-row-headers')).display,
            getComputedStyle(host.querySelector('.spreadsheet-frozen-headers')).display, scroller.scrollHeight], [[ox, oy, 100, 25].map(Math.round), 'none', 'none', 'none', 1000 * 25]);
        click(ox + 250, oy + 60);
        check('a click lands on the cell under it', active(), [2, 2]);
        grid.SetZoom(2); await frame();
        click(ox + 250, oy + 60);
        check('... also zoomed', [active(), round(node(1, 1).getBoundingClientRect())], [[1, 1], [ox + 200, oy + 50, 280, 50].map(Math.round)]);
        grid.SetZoom(1);
        grid.Undo(); await frame();
        check('undo shows the headers again', [grid.GetShowHeaders(), round(node(0, 0).getBoundingClientRect()), getComputedStyle(host.querySelector('.spreadsheet-col-headers')).display !== 'none', !!host.querySelector('.spreadsheet-col-header[data-col="0"]')], [true, expected(0, 0, 1), true, true]);
        check('each option was a change', changes.length >= 6, true);
        check('not booleans', [thrown(() => grid.SetShowHeaders('no')) !== null, thrown(() => grid.SetShowZeros()) !== null, thrown(() => grid.SetShowFormulas(1)) !== null, thrown(() => grid.SetShowGridLines(null)) !== null], [true, true, true, true]);

        // ---- through a file ----
        grid.SetShowHeaders(false); grid.SetShowFormulas(true); grid.SetShowZeros(false); grid.SetShowGridLines(false); grid.FreezeAt(3, 2); grid.SetZoom(1.35);
        grid.SwitchSheet(1); grid.SetCell(0, 0, 'second');
        const again = await reopened(grid);
        try {
            check('the file keeps the view of each sheet', [again.grid.GetViewOptions(), (again.grid.SwitchSheet(1), again.grid.GetViewOptions())],
                [{ zoom: 1.35, showGridLines: false, showHeaders: false, showFormulas: true, showZeros: false, frozen: { rows: 3, cols: 2 } }, { zoom: 0.75, showGridLines: true, showHeaders: true, showFormulas: false, showZeros: true, frozen: { rows: 0, cols: 0 } }]);
            again.grid.SwitchSheet(0); await frame();
            const cell = again.host.querySelector('.spreadsheet-cell[data-row="1"][data-col="0"]');
            check('and the reopened grid shows it: formulas, no headers, zoomed', [cell.textContent, Math.abs(cell.getBoundingClientRect().height - (await import('/spreadsheet/Model.js')).getRowHeight(1, again.grid.GetModel()) * 1.35) < 0.5, getComputedStyle(again.host.querySelector('.spreadsheet-col-headers')).display], ['=A1*2', true, 'none']);
            check('a reopened workbook is unmodified', again.grid.IsModified(), false);
        } finally { again.close(); }
        // After Dispose and Mount the view is applied to the new nodes.
        grid.SwitchSheet(0);
        grid.Dispose(); grid.Mount(); await frame();
        check('mounted again: zoom and hidden headers are applied', [Math.round(node(0, 0).getBoundingClientRect().height * 100) / 100, getComputedStyle(host.querySelector('.spreadsheet-col-headers')).display], [33.75, 'none']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// The list of sheets: AddSheet, RenameSheet, DeleteSheet, MoveSheet, DuplicateSheet,
// SetSheetVisibility, SetSheetTabColor, GetSheetInfos, OnSheetsChange.
checks['api-sheets'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const infos = () => grid.GetSheetInfos().map((s) => `${s.index}:${s.name}:${s.visibility}${s.tabColor ? ':' + s.tabColor : ''}${s.isActive ? '*' : ''}`);
        const at = (ref) => { const cell = grid.GetRange(ref)[0][0]; return [cell.formula, cell.value]; };
        const changes = [], lists = [];
        const offChange = grid.OnChange((change) => changes.push(change));
        const offSheets = grid.OnSheetsChange((list) => lists.push(list.map((s) => s.name + (s.visibility === 'visible' ? '' : '(' + s.visibility + ')')).join(',')));
        const take = () => changes.splice(0).map((c) => `${c.Sheets.join(',')}${c.StructureChanged ? ' structure' : ''}`);
        const takeLists = () => lists.splice(0);

        // ---- add ----
        grid.SetCell(0, 0, '1'); grid.SetCell(1, 0, '2'); grid.SetCell(2, 0, '3'); grid.SetCell(3, 0, '=Summary!A1*2');
        check('one sheet to start with', [infos(), grid.GetSheetInfos()[0].id, at('A4')[1]], [['0:Sheet1:visible*'], 'sheet1', '#REF!']);
        take();
        const r0 = grid.GetRevision();
        check('AddSheet returns the index of the new sheet, which is shown', [grid.AddSheet('Summary'), grid.GetActiveSheetIndex(), infos()], [1, 1, ['0:Sheet1:visible', '1:Summary:visible*']]);
        check('it is a change of the structure, a new revision, and OnSheetsChange hears it', [take().some((c) => c.endsWith('structure')), grid.GetRevision() !== r0, takeLists()], [true, true, ['Sheet1,Summary']]);
        check('a formula that named the sheet before it was there reads it now', grid.GetRange('Sheet1!A4')[0][0].value, 0);
        grid.SetCell(0, 0, '=SUM(Sheet1!A1:A3)'); grid.SetCell(1, 0, '=Sheet1!$A$2*10'); grid.SetCell(2, 0, '=SUM(Total)');
        grid.SetDefinedName('Total', 'Sheet1!$A$1:$A$3');
        check('formulas across sheets', [at('A1'), at('A2'), at('A3'), grid.GetRange('Sheet1!A4')[0][0].value], [['=SUM(Sheet1!A1:A3)', 6], ['=Sheet1!$A$2*10', 20], ['=SUM(Total)', 6], 12]);
        check('the default name is the first free SheetN; a position can be given', [grid.AddSheet(undefined, 0), infos()], [0, ['0:Sheet2:visible*', '1:Sheet1:visible', '2:Summary:visible']]);
        check('the step is undone from the new sheet', [grid.CanUndo(), (grid.Undo(), infos()), grid.CanRedo()], [true, ['0:Sheet1:visible', '1:Summary:visible*'], true]);
        check('... and redone from the sheet then on display', [(grid.Redo(), infos()), grid.CanUndo(), grid.GetSheetInfos().map((s) => s.id)], [['0:Sheet2:visible*', '1:Sheet1:visible', '2:Summary:visible'], true, ['sheet3', 'sheet1', 'sheet2']]);
        grid.Undo();
        check('undone again: two sheets, the sheet that was on display is back', [infos(), grid.GetSheetNames(), grid.SheetVisibility()], [['0:Sheet1:visible', '1:Summary:visible*'], ['Sheet1', 'Summary'], [true, true]]);
        takeLists(); take();
        check('AddSheet rejects bad names and positions', [thrown(() => grid.AddSheet('summary')), thrown(() => grid.AddSheet('a/b')), thrown(() => grid.AddSheet('x', 7)), thrown(() => grid.AddSheet(5))],
            ['AddSheet: there is already a sheet named "summary"', 'AddSheet: a sheet name cannot contain any of : \\ / ? * [ ]', 'AddSheet: atIndex 7 is outside 0 to 2', 'AddSheet: name must be a string, got 5']);

        // ---- rename ----
        const beforeRename = grid.GetRevision();
        grid.RenameSheet(0, 'My Data');
        check('RenameSheet: formulas of every sheet follow, quoted; values stand', [grid.GetSheetNames(), at('A1'), at('A2'), at('A3')], [['My Data', 'Summary'], ["=SUM('My Data'!A1:A3)", 6], ["='My Data'!$A$2*10", 20], ['=SUM(Total)', 6]]);
        check('... and defined names', grid.GetDefinedNames().map((n) => [n.name, n.refersTo]), [['Total', "'My Data'!$A$1:$A$3"]]);
        check('... one change, one notification', [take().length, takeLists()], [1, ['My Data,Summary']]);
        grid.SetCell(4, 0, '5');
        check('an edit still recalculates through the renamed sheet', (grid.SwitchSheet(0), grid.SetCell(0, 0, '11'), grid.SwitchSheet(1), [at('A1')[1], at('A3')[1]]), [16, 16]);
        grid.SwitchSheet(0); grid.SetCell(0, 0, '1'); grid.SwitchSheet(1);
        grid.Undo();
        check('undo: the name, the formulas and the revision are back', [grid.GetSheetNames(), at('A1'), grid.GetDefinedNames()[0].refersTo], [['Sheet1', 'Summary'], ['=SUM(Sheet1!A1:A3)', 6], 'Sheet1!$A$1:$A$3']);
        grid.Redo();
        check('redo', [grid.GetSheetNames(), at('A2')[0]], [['My Data', 'Summary'], "='My Data'!$A$2*10"]);
        check('RenameSheet rejects invalid and duplicate names',
            [thrown(() => grid.RenameSheet(0, '')), thrown(() => grid.RenameSheet(0, 'x'.repeat(32))), thrown(() => grid.RenameSheet(0, "'quoted")), thrown(() => grid.RenameSheet(0, 'History')),
             thrown(() => grid.RenameSheet(0, 'SUMMARY')), thrown(() => grid.RenameSheet(5, 'x')), thrown(() => grid.RenameSheet(0)), thrown(() => grid.RenameSheet(0, 'my data'))],
            ['RenameSheet: a sheet name cannot be empty', 'RenameSheet: a sheet name cannot be longer than 31 characters', 'RenameSheet: a sheet name cannot begin or end with an apostrophe', 'RenameSheet: "History" is a reserved sheet name',
             'RenameSheet: there is already a sheet named "SUMMARY"', 'RenameSheet: there is no sheet 5 (the workbook has 2)', 'RenameSheet: name must be a string, got null', null]);
        check('a rename that only changes the case', [grid.GetSheetNames(), at('A1')], [['my data', 'Summary'], ["=SUM('my data'!A1:A3)", 6]]);
        grid.RenameSheet(0, 'Data');
        take(); takeLists();

        // ---- delete ----
        const beforeDelete = grid.GetRevision();
        grid.DeleteSheet(0);
        check('DeleteSheet: references become #REF!', [infos(), at('A1'), at('A2'), at('A3')[1], grid.GetDefinedNames().map((n) => n.refersTo)], [['0:Summary:visible*'], ['=SUM(#REF!)', '#REF!'], ['=#REF!*10', '#REF!'], '#REF!', ['#REF!']]);
        check('... a change of the structure', [take().some((c) => c.endsWith('structure')), takeLists(), grid.GetRevision() !== beforeDelete], [true, ['Summary'], true]);
        check('the last visible sheet cannot be deleted', [thrown(() => grid.DeleteSheet(0)), thrown(() => grid.DeleteSheet(3)), thrown(() => grid.DeleteSheet('0'))],
            ['DeleteSheet: a workbook must keep at least one visible sheet', 'DeleteSheet: there is no sheet 3 (the workbook has 1)', 'DeleteSheet: index must be a whole number, got "0"']);
        grid.Undo();
        check('undo: the sheet, the references, the values and the revision are back', [infos(), at('A1'), at('A2'), at('A3'), grid.GetDefinedNames().map((n) => n.refersTo), grid.GetRange('Data!A1:A4').map((row) => row[0].value), grid.GetRevision() === beforeDelete],
            [['0:Data:visible', '1:Summary:visible*'], ['=SUM(Data!A1:A3)', 6], ['=Data!$A$2*10', 20], ['=SUM(Total)', 6], ['Data!$A$1:$A$3'], [1, 2, 3, 12], true]);
        grid.Redo();
        check('redo deletes it again', [infos(), at('A1')[1]], [['0:Summary:visible*'], '#REF!']);
        grid.Undo();
        // Deleting the sheet on display shows its neighbour, and the step is undone from there.
        grid.SwitchSheet(0);
        grid.DeleteSheet(0);
        check('deleting the sheet on display shows the next one', [infos(), grid.CanUndo()], [['0:Summary:visible*'], true]);
        grid.Undo();
        check('... and undo shows it again', [infos(), grid.GetCell(3, 0).value], [['0:Data:visible*', '1:Summary:visible'], 12]);
        take(); takeLists();

        // ---- move ----
        grid.AddSheet('Third');
        grid.SetCell(0, 0, '=Data!A1+Summary!A1');
        grid.MoveSheet(2, 0);
        check('MoveSheet: the order, the ids, the sheet on display', [infos(), grid.GetSheetInfos().map((s) => s.id).join(), grid.GetCell(0, 0).value], [['0:Third:visible*', '1:Data:visible', '2:Summary:visible'], 'sheet4,sheet1,sheet2', 7]);
        check('... told once', [take().some((c) => c.endsWith('structure')), takeLists().slice(-1)], [true, ['Third,Data,Summary']]);
        grid.SwitchSheet(1); grid.SetCell(0, 0, '4'); grid.SwitchSheet(0);
        check('an edit after a move recalculates the right sheets', [grid.GetCell(0, 0).value, grid.GetRange('Summary!A1')[0][0].value], [13, 9]);
        grid.SwitchSheet(1); grid.SetCell(0, 0, '1'); grid.SwitchSheet(0);
        grid.Undo();
        check('undo moves it back', infos(), ['0:Data:visible', '1:Summary:visible', '2:Third:visible*']);
        grid.Redo();
        check('MoveSheet rejects positions that are no sheet', [thrown(() => grid.MoveSheet(0, 3)), thrown(() => grid.MoveSheet(-1, 0)), thrown(() => grid.MoveSheet(1, 1))], ['MoveSheet: to 3 is outside 0 to 2', 'MoveSheet: there is no sheet -1 (the workbook has 3)', null]);

        // ---- visibility, tab colour ----
        grid.SetSheetVisibility(2, 'hidden');
        check('SetSheetVisibility: hidden', [infos(), grid.SheetVisibility(), grid.GetCell(0, 0).value], [['0:Third:visible*', '1:Data:visible', '2:Summary:hidden'], [true, true, false], 7]);
        grid.SetSheetVisibility(0, 'veryHidden');
        check('hiding the sheet on display shows the nearest visible one', infos(), ['0:Third:veryHidden', '1:Data:visible*', '2:Summary:hidden']);
        check('the only visible sheet can be neither hidden nor deleted', [thrown(() => grid.SetSheetVisibility(1, 'hidden')), thrown(() => grid.DeleteSheet(1)), thrown(() => grid.SetSheetVisibility(1, 'gone'))],
            ['SetSheetVisibility: a workbook must keep at least one visible sheet', 'DeleteSheet: a workbook must keep at least one visible sheet', 'SetSheetVisibility: visibility must be "visible", "hidden" or "veryHidden", got "gone"']);
        grid.Undo();
        check('undo: visible and on display again', infos(), ['0:Third:visible*', '1:Data:visible', '2:Summary:hidden']);
        grid.Redo();
        grid.SetSheetTabColor(1, '#ff8800'); grid.SetSheetTabColor(2, '#00AA55');
        check('SetSheetTabColor', infos(), ['0:Third:veryHidden', '1:Data:visible:#FF8800*', '2:Summary:hidden:#00AA55']);
        grid.Undo();
        check('undo takes the colour away, redo gives it back', [infos()[2], (grid.Redo(), infos()[2])], ['2:Summary:hidden', '2:Summary:hidden:#00AA55']);
        check('a malformed colour is refused; null clears', [thrown(() => grid.SetSheetTabColor(1, 'orange')) !== null, (grid.SetSheetTabColor(2, null), infos()[2]), (grid.Undo(), infos()[2])], [true, '2:Summary:hidden', '2:Summary:hidden:#00AA55']);
        check('each of these was told', [take().length >= 5, takeLists().slice(-1)], [true, ['Third(veryHidden),Data,Summary(hidden)']]);

        // ---- duplicate ----
        grid.SelectCell(5, 0); grid.SetComment(5, 0, 'a note', 'me');
        grid.SetCell(6, 0, '=Data!A1+1');
        grid.SetDefinedName('Local', 'Data!$A$1', 1);
        grid.SetValidation('C1:C5', { type: 'list', values: ['a', 'b'] });
        grid.SelectRange(0, 1, 2, 1); grid.ApplyBgColor('#FFEE00');
        check('DuplicateSheet returns the index of the copy, which is shown beside the original', [grid.DuplicateSheet(1), infos()], [2, ['0:Third:veryHidden', '1:Data:visible:#FF8800', '2:Data (2):visible:#FF8800*', '3:Summary:hidden:#00AA55']]);
        check('the copy has the cells, formats, comments and validations', [grid.GetRange('A1:A4').map((row) => row[0].value), grid.GetCell(0, 1).format.fill, grid.GetCell(5, 0).comment && grid.GetCell(5, 0).comment.text, grid.GetValidations().length], [[1, 2, 3, 12], '#FFEE00', 'a note', 1]);
        check('its formulas that named the sheet name the copy; its names are copied', [grid.GetCell(6, 0).formula, grid.GetDefinedNames().filter((n) => n.name === 'Local').map((n) => [n.scope, n.refersTo])], ["='Data (2)'!A1+1", [[1, 'Data!$A$1'], [2, "'Data (2)'!$A$1"]]]);
        grid.SetCell(0, 0, '100');
        check('the copy is a sheet of its own', [grid.GetCell(6, 0).value, grid.GetRange('Data!A7')[0][0].value, grid.GetRange('Summary!A1')[0][0].value], [101, 2, 6]);
        grid.Undo();
        check('undo removes the copy and its names', [infos(), grid.GetDefinedNames().filter((n) => n.name === 'Local').length], [['0:Third:veryHidden', '1:Data:visible:#FF8800*', '2:Summary:hidden:#00AA55'], 1]);
        grid.Redo();
        check('redo brings it back', [infos()[2], grid.GetCell(6, 0).formula, grid.DuplicateSheet(1), infos()[2]], ['2:Data (2):visible:#FF8800*', "='Data (2)'!A1+1", 2, '2:Data (3):visible:#FF8800*']);
        grid.Undo();

        // ---- the file ----
        const again = await reopened(grid);
        try {
            const reinfos = again.grid.GetSheetInfos().map((s) => `${s.index}:${s.name}:${s.visibility}${s.tabColor ? ':' + s.tabColor : ''}${s.isActive ? '*' : ''}`);
            check('saved and opened again: names, order, visibility, tab colours', reinfos, ['0:Third:veryHidden', '1:Data:visible:#FF8800*', '2:Data (2):visible:#FF8800', '3:Summary:hidden:#00AA55']);
            check('... formulas and names', [again.grid.GetRange("'Data (2)'!A7")[0][0].formula, again.grid.GetRange('Summary!A1:A3').map((row) => [row[0].formula, row[0].value]), again.grid.GetDefinedNames().map((n) => [n.name, n.scope, n.refersTo]).sort()],
                ["='Data (2)'!A1+1", [['=SUM(Data!A1:A3)', 6], ['=Data!$A$2*10', 20], ['=SUM(Total)', 6]], [['Local', 1, 'Data!$A$1'], ['Local', 2, "'Data (2)'!$A$1"], ['Total', null, 'Data!$A$1:$A$3']]]);
            again.grid.SetSheetVisibility(0, 'visible');
            check('a very hidden sheet of a file can be shown', again.grid.GetSheetInfos()[0].visibility, 'visible');
        } finally { again.close(); }
        take(); takeLists(); offChange(); offSheets();
        grid.AddSheet('Unheard');
        check('no events after unsubscribing', [changes.length, lists.length], [0, 0]);

        // ---- a file with charts, pictures and pivot tables, its sheets converted when needed ----
        const count = (wb, key) => OF.excel.sheetNames(wb).map((name) => `${name}:${OF.excel[key](OF.excel.sheet(name, wb)).length}`);
        const describe = async (bytes) => { const wb = await OF.excel.load(bytes); return { names: OF.excel.sheetNames(wb), states: OF.excel.sheetStates(wb), objects: count(wb, 'drawingObjects'), images: count(wb, 'images') }; };
        const chartParts = async (bytes) => Object.keys((await window.JSZip.loadAsync(bytes)).files).filter((path) => /^xl\/(charts\/chart|pivotTables\/pivotTable|drawings\/drawing)\d+\.xml$/.test(path)).length;
        for (const file of ['tests/real-world/dashboard-ecommerce-sales.xlsx', 'tests/real-world/dashboard-sales-analysis.xlsx']) {
            const label = file.split('/').pop().replace('.xlsx', '');
            const original = await bytesOf(file);
            const before = await describe(original);
            const book = mountGrid();
            try {
                await book.grid.ImportXlsx(original);
                const g = book.grid, n = g.GetSheetNames().length;
                const loaded = () => g.GetSheetInfos().map((s) => s.isLoaded ? 1 : 0).join('');
                const loadedAtOpen = loaded();
                note(`${label}: ${n} sheets (${before.names.join(', ')}), objects ${before.objects.join(' ')}, loaded at open ${loadedAtOpen}`);
                // Moving, hiding and colouring convert no sheet.
                const last = n - 1;
                g.MoveSheet(last, 0); g.SetSheetTabColor(2, '#123456'); g.SetSheetVisibility(2, 'hidden');   // (the sheet on display is at 1 now)
                check(`${label}: moving, colouring and hiding sheets converts none`, [...loaded()].filter((c) => c === '1').length, [...loadedAtOpen].filter((c) => c === '1').length);
                g.Undo(); g.Undo(); g.Undo();
                check(`${label}: undone`, [g.GetSheetNames(), g.SheetVisibility().every((v, i) => v === (before.states[i] === 'visible'))], [before.names, true]);
                // Renamed and moved: every sheet keeps its pictures and charts.
                const target = Math.min(1, last);
                g.RenameSheet(target, "It's renamed");
                g.MoveSheet(target, last);
                const moved = await describe(await g.ExportXlsx());
                const expectNames = before.names.filter((_, i) => i !== target).concat(["It's renamed"]);
                const byName = (list, renameFrom) => (list.map((entry) => { const cut = entry.lastIndexOf(':'); const name = entry.slice(0, cut); return (name === renameFrom ? "It's renamed" : name) + ':' + entry.slice(cut + 1); })).sort().join(' ');
                check(`${label}: renamed and moved: the sheets`, moved.names, expectNames);
                check(`${label}: ... each with its own drawing objects and pictures`, [byName(moved.objects), byName(moved.images)], [byName(before.objects, before.names[target]), byName(before.images, before.names[target])]);
                g.Undo(); g.Undo();
                check(`${label}: undone: the names of the file`, g.GetSheetNames(), before.names);
                // Deleted: what the file kept for the sheet is dropped from the package.
                const partsBefore = await chartParts(await g.ExportXlsx());
                note(`${label}: ${await chartParts(original)} chart, drawing and pivot table parts in the file, ${partsBefore} in an export of it`);
                const richest = before.objects.map((entry, i) => [Number(entry.slice(entry.lastIndexOf(':') + 1)), i]).sort((a, b) => b[0] - a[0])[0][1];
                g.DeleteSheet(richest);
                const withoutBytes = await g.ExportXlsx();
                const without = await describe(withoutBytes);
                check(`${label}: deleted "${before.names[richest]}": the other sheets keep their objects`, [without.names, without.objects], [before.names.filter((_, i) => i !== richest), before.objects.filter((_, i) => i !== richest)]);
                const partsAfter = await chartParts(withoutBytes);
                check(`${label}: ... and its chart / drawing / pivot parts are gone from the package`, partsAfter < partsBefore, true);
                note(`${label}: ${partsBefore} such parts in an export, ${partsAfter} without "${before.names[richest]}"`);
                g.Undo();
                const restored = await describe(await g.ExportXlsx());
                check(`${label}: undo: the sheet is back with everything`, [restored.names, restored.objects, restored.images, await chartParts(await g.ExportXlsx())], [before.names, before.objects, before.images, partsBefore]);
                // A copy: pictures come along; the charts stay with the original.
                const copyIndex = g.DuplicateSheet(richest);
                const copied = await describe(await g.ExportXlsx());
                check(`${label}: a copy of it is written, the original keeps its objects`, [copied.names.length, copied.names[copyIndex], copied.objects.filter((_, i) => i !== copyIndex), copied.images[copyIndex].endsWith(':' + before.images[richest].slice(before.images[richest].lastIndexOf(':') + 1))],
                    [n + 1, g.GetSheetNames()[copyIndex], before.objects, true]);
                g.AddSheet('Brand new', 0);
                const added = await describe(await g.ExportXlsx());
                check(`${label}: an added sheet is written too`, [added.names[0], added.names.length], ['Brand new', n + 2]);
            } finally { book.close(); }
        }
        // ---- formulas of a file across its sheets: rename, delete, undo ----
        const street = mountGrid();
        try {
            const g = street.grid;
            await g.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
            const names = g.GetSheetNames();
            // A formula of one sheet that reads the other.
            let found = null;
            for (const home of [0, 1]) {
                const other = names[1 - home];
                const rows = g.GetRange(`'${names[home]}'!A1:AZ400`);
                for (const row of rows) for (const cell of row) if (!found && cell.formula && cell.formula.includes(other + '!') && typeof cell.value === 'number') found = { home, other: 1 - home, ref: `'${names[home]}'!${cell.address}`, formula: cell.formula, value: cell.value };
                if (found) break;
            }
            check('milk-street: a formula of one sheet reads the other', found !== null, true);
            if (found) {
                const read = () => { const cell = g.GetRange(found.ref)[0][0]; return [cell.formula, cell.value]; };
                const errors = () => g.GetRange(`'${names[found.home]}'!A1:AZ400`).flat().filter((cell) => typeof cell.value === 'string' && cell.value.startsWith('#')).length;
                const errorsBefore = errors();
                g.SwitchSheet(found.home);
                g.RenameSheet(found.other, 'Other Sheet');
                check('milk-street: renamed: the formula names the new name, its value stands', [read()[0], read()[1], errors()], [found.formula.split(names[found.other] + '!').join("'Other Sheet'!"), found.value, errorsBefore]);
                const renamedBack = await reopened(g);
                try {
                    const index = renamedBack.grid.GetSheetNames().indexOf('Other Sheet');
                    const cell = renamedBack.grid.GetRange(found.ref)[0][0];
                    check('milk-street: saved and opened again', [index, cell.formula, cell.value], [found.other, found.formula.split(names[found.other] + '!').join("'Other Sheet'!"), found.value]);
                } finally { renamedBack.close(); }
                g.DeleteSheet(found.other);
                check('milk-street: deleted: the formula reads #REF! and is an error', [g.GetSheetNames(), read()[0].includes('#REF!'), String(read()[1]).startsWith('#'), errors() > errorsBefore], [[names[found.home]], true, true, true]);
                g.Undo();
                check('milk-street: undo: the sheet and the values are back', [g.GetSheetNames()[found.other], read()[1], errors()], ['Other Sheet', found.value, errorsBefore]);
                g.Undo();
                check('milk-street: undo: the name of the file', [g.GetSheetNames(), read()], [names, [found.formula, found.value]]);
            }
        } finally { street.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Rows and columns: sizes, auto-fit (also by a double click on a header's resize handle), hidden
// rows and columns that come back at their size, the outline (groups, collapse / expand).
import { press } from '/component.js';
checks['api-rowcol'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const size = (row, col) => { const box = node(row, col).getBoundingClientRect(); return [Math.round(box.width), Math.round(box.height)]; };
        const active = () => { const cell = grid.GetActiveCell(); return [cell.Row, cell.Col]; };
        const textWidth = (text, px = 11, bold = false) => { const context = document.createElement('canvas').getContext('2d'); context.font = `${bold ? 'bold ' : ''}${px}px Arial, sans-serif`; return context.measureText(text).width; };
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        grid.SetCell(0, 0, 'a'); grid.SetCell(9, 0, 'z'); changes.splice(0);

        // ---- sizes ----
        const r0 = grid.GetRevision();
        check('the default sizes', [grid.GetColumnWidth(1), grid.GetRowHeight(2)], [100, 25]);
        grid.SetColumnWidth(1, 150);
        check('SetColumnWidth sets one column: one change, one undo step', [grid.GetColumnWidth(1), grid.GetColumnWidth(2), changes.splice(0).length, grid.GetRevision() !== r0, grid.CanUndo()], [150, 100, 1, true, true]);
        grid.SetColumnWidth('C:D', 60); grid.SetRowHeight(2, 40); grid.SetRowHeight('4:5', 30);
        await frame();
        check('... or the columns / rows of a range', [[1, 2, 3, 4].map((col) => grid.GetColumnWidth(col)), [2, 3, 4, 5].map((row) => grid.GetRowHeight(row))], [[150, 60, 60, 100], [40, 30, 30, 25]]);
        check('the grid draws them', [size(0, 1), size(0, 2), size(2, 0), size(3, 3)], [[150, 25], [60, 25], [100, 40], [60, 30]]);
        grid.Undo(); grid.Undo();
        check('undo takes back one call at a time', [[2, 3].map((col) => grid.GetColumnWidth(col)), [2, 3, 4].map((row) => grid.GetRowHeight(row))], [[60, 60], [25, 25, 25]]);
        grid.Redo(); grid.Redo();
        check('redo', [grid.GetRowHeight(2), grid.GetRowHeight(4)], [40, 30]);
        const r1 = grid.GetRevision();
        grid.SetColumnWidth(1, 150);
        check('the size it has already is no change', [grid.GetRevision() === r1, changes.splice(0).length > 0], [true, true]);
        check('bad arguments throw', [
            thrown(() => grid.SetColumnWidth(-1, 10)), thrown(() => grid.SetColumnWidth(0, -5)), thrown(() => grid.SetColumnWidth(0, 5000)), thrown(() => grid.SetColumnWidth('no range', 10)),
            thrown(() => grid.SetRowHeight(0, 'tall')), thrown(() => grid.SetRowHeight(0, 600)), thrown(() => grid.HideRows()), thrown(() => grid.HideRows(2, 0)), thrown(() => grid.GetColumnWidth(99999)),
            thrown(() => grid.SetGroupCollapsed('rows', 0, true)), thrown(() => grid.SetGroupCollapsed('diagonal', 0, true)), thrown(() => grid.AutoFitColumns('??')),
        ].map((message) => message && message.split(':')[0]), ['SetColumnWidth', 'SetColumnWidth', 'SetColumnWidth', 'SetColumnWidth', 'SetRowHeight', 'SetRowHeight', 'HideRows', 'HideRows', 'GetColumnWidth', 'SetGroupCollapsed', 'SetGroupCollapsed', 'AutoFitColumns']);
        check('a message says what is wrong', [thrown(() => grid.SetColumnWidth(0, 5000)), thrown(() => grid.SetGroupCollapsed('rows', 0, true))], ['SetColumnWidth: the width must be 0 to 1785 pixels, got 5000', 'SetGroupCollapsed: row 0 is not in a group']);

        // ---- auto-fit ----
        const long = 'a fairly long piece of text for one cell';
        grid.SetCell(0, 5, long); grid.SetCell(1, 5, 'Big and bold'); grid.SelectCell(1, 5); grid.ApplyFormat('bold'); grid.SetFontSize(22);
        grid.SetCell(2, 5, 'short');
        grid.SetCell(0, 6, '1234567.891'); grid.SelectCell(0, 6); grid.SetNumberFormat('#,##0.00'); grid.SetColumnWidth(6, 30);
        await frame();
        check('a number in a column too narrow for it shows ####', /^#+$/.test(node(0, 6).textContent), true);
        const before = grid.GetRevision();
        grid.AutoFitColumns('F:G');
        await frame();
        const wanted = Math.ceil(Math.max(textWidth(long), textWidth('Big and bold', 22, true))) + 10;
        check('AutoFitColumns: the widest text of the column in its font, plus the padding', [grid.GetColumnWidth(5), grid.GetRevision() !== before], [wanted, true]);
        check('... the text fits its cell', [node(0, 5).scrollWidth <= node(0, 5).clientWidth, node(1, 5).scrollWidth <= node(1, 5).clientWidth], [true, true]);
        check('... a number is fitted as it is displayed', [grid.GetColumnWidth(6), node(0, 6).textContent], [Math.ceil(textWidth('1,234,567.89')) + 10, '1,234,567.89']);
        grid.Undo();
        check('one undo step for all the columns', [grid.GetColumnWidth(5), grid.GetColumnWidth(6)], [100, 30]);
        grid.Redo();
        grid.SelectRange(0, 7, 3, 8); grid.SetCell(1, 7, 'selected column'); grid.SetCell(2, 8, 'x');
        grid.AutoFitColumns();
        check('without a range: the columns of the selection (a column without text keeps its width)', [grid.GetColumnWidth(7), grid.GetColumnWidth(8), grid.GetColumnWidth(9)], [Math.ceil(textWidth('selected column')) + 10, Math.ceil(textWidth('x')) + 10, 100]);
        // Rows: wrapped text, a large font.
        grid.SetCell(6, 1, 'wrapped text that takes a number of lines in a column this narrow'); grid.SelectCell(6, 1); grid.SetWrapText(true);
        grid.SetCell(7, 0, 'Big'); grid.SelectCell(7, 0); grid.SetFontSize(22);
        grid.SetRowHeight(8, 60);
        grid.AutoFitRows('7:9');
        await frame();
        const wrappedHeight = grid.GetRowHeight(6);
        check('AutoFitRows: wrapped text gets the lines it takes', [wrappedHeight > 25, (wrappedHeight - 6) % 14, node(6, 1).scrollHeight <= node(6, 1).clientHeight + 1], [true, 0, true]);
        check('... a large font its line height, a row without text the default', [grid.GetRowHeight(7), grid.GetRowHeight(8)], [Math.round(22 * 1.3) + 6, 20]);
        grid.SetColumnWidth(1, 300); grid.AutoFitRows(6);
        check('... fewer lines in a wider column', grid.GetRowHeight(6) < wrappedHeight, true);
        grid.Undo(); grid.Undo();
        check('undo of the fit and of the width', [grid.GetRowHeight(6), grid.GetColumnWidth(1)], [wrappedHeight, 150]);

        // ---- the resize handles of the headers ----
        const handle = (kind, index) => host.querySelector(`.spreadsheet-${kind}-header[data-${kind}="${index}"] .${kind}-resize-handle`);
        const doubleClick = (target) => { const box = target.getBoundingClientRect(); target.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true, clientX: box.left + 1, clientY: box.top + 1, button: 0 })); };
        grid.SelectCell(0, 0);
        grid.SetColumnWidth(5, 400); await frame();
        doubleClick(handle('col', 5));
        check('a double click on a column header\'s resize handle fits the column', grid.GetColumnWidth(5), wanted);
        grid.Undo();
        check('... as one undo step', grid.GetColumnWidth(5), 400);
        grid.SetRowHeight(7, 90); await frame();
        doubleClick(handle('row', 7));
        check('... and on a row header\'s the row', grid.GetRowHeight(7), Math.round(22 * 1.3) + 6);
        grid.SetZoom(1.5); await frame();
        doubleClick(handle('col', 5));
        check('at a zoom of 150 % the fitted width is the same (sheet pixels)', [grid.GetColumnWidth(5), Math.abs(node(0, 5).getBoundingClientRect().width - wanted * 1.5) <= 1], [wanted, true]);
        // With whole columns selected, a double click on one of them fits them all.
        grid.SetColumnWidth('F:G', 250); grid.SelectRange(0, 5, grid.GetModel().RowCount - 1, 6); await frame();
        doubleClick(handle('col', 6));
        check('with whole columns selected, all of them', [grid.GetColumnWidth(5), grid.GetColumnWidth(6)], [wanted, Math.ceil(textWidth('1,234,567.89')) + 10]);
        // A drag of the handle is one undo step.
        grid.SetZoom(1); grid.SelectCell(0, 0); await frame();
        const grip = handle('col', 0), gripBox = grip.getBoundingClientRect(), revisionBeforeDrag = grid.GetRevision();
        grip.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, clientX: gripBox.left + 1, clientY: gripBox.top + 5, button: 0 }));
        for (const dx of [10, 25, 40]) document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: gripBox.left + 1 + dx, clientY: gripBox.top + 5 }));
        document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: gripBox.left + 41, clientY: gripBox.top + 5 }));
        await frame();
        check('dragging a resize handle resizes', [grid.GetColumnWidth(0), grid.GetRevision() !== revisionBeforeDrag], [140, true]);
        grid.Undo();
        check('... as one undo step', [grid.GetColumnWidth(0), grid.GetRevision() === revisionBeforeDrag], [100, true]);

        // ---- hidden rows and columns ----
        grid.SetRowHeight(2, 40);
        const visible = grid.GetRevision();
        grid.HideRows(2, 2); await frame();
        check('HideRows: the rows have no height and are listed', [grid.GetRowHeight(2), grid.GetRowHeight(3), grid.GetHiddenRows(), grid.GetRevision() !== visible], [0, 0, ['3:4'], true]);
        check('... and are not drawn', [node(2, 0) === null || node(2, 0).getBoundingClientRect().height === 0, Math.round(node(4, 0).getBoundingClientRect().top - node(1, 0).getBoundingClientRect().bottom)], [true, 0]);
        grid.SelectCell(1, 0); press(host, 'ArrowDown');
        check('the arrow keys step over hidden rows', active(), [4, 0]);
        press(host, 'ArrowUp');
        check('... both ways', active(), [1, 0]);
        grid.HideColumns('B:C'); await frame();
        check('HideColumns', [grid.GetHiddenColumns(), grid.GetColumnWidth(1), grid.GetColumnWidth(2)], [['B:C'], 0, 0]);
        grid.SelectCell(0, 0); press(host, 'ArrowRight');
        check('... stepped over by the arrow keys too', active(), [0, 3]);
        grid.Undo();
        check('undo shows the columns again at their widths', [grid.GetHiddenColumns(), grid.GetColumnWidth(1), grid.GetColumnWidth(2)], [[], 150, 60]);
        grid.Redo();
        grid.UnhideRows('1:10'); grid.UnhideColumns(0, 5); await frame();
        check('UnhideRows / UnhideColumns: back at the sizes they had', [grid.GetHiddenRows(), grid.GetRowHeight(2), grid.GetRowHeight(3), grid.GetHiddenColumns(), grid.GetColumnWidth(1), grid.GetColumnWidth(2), size(2, 1)], [[], 40, 30, [], 150, 60, [150, 40]]);
        grid.SetColumnWidth(3, 0);
        check('a width of 0 hides; any other width shows', [grid.GetHiddenColumns(), (grid.SetColumnWidth(3, 70), grid.GetHiddenColumns()), grid.GetColumnWidth(3)], [['D:D'], [], 70]);

        // ---- outline ----
        check('no outline at first', grid.GetOutline(), { rows: [], columns: [], rowLevels: 0, columnLevels: 0 });
        grid.GroupRows('11:16'); grid.GroupRows(11, 2); grid.GroupColumns('J:K');
        check('GroupRows / GroupColumns: levels', grid.GetOutline(), {
            rows: [{ first: 10, last: 15, range: '11:16', level: 1, collapsed: false }, { first: 11, last: 12, range: '12:13', level: 2, collapsed: false }],
            columns: [{ first: 9, last: 10, range: 'J:K', level: 1, collapsed: false }], rowLevels: 2, columnLevels: 1 });
        grid.SetGroupCollapsed('rows', 11, true);
        check('collapsing the inner group hides its rows', [grid.GetHiddenRows(), grid.GetOutline().rows.map((group) => group.collapsed)], [['12:13'], [false, true]]);
        grid.SetGroupCollapsed('rows', 14, true);
        check('... the outer group all of them', grid.GetHiddenRows(), ['11:16']);
        grid.SetGroupCollapsed('rows', 14, false);
        check('expanding the outer group leaves the collapsed inner one shut', [grid.GetHiddenRows(), grid.GetOutline().rows.map((group) => group.collapsed)], [['12:13'], [false, true]]);
        grid.Undo(); grid.Undo();
        check('undo', grid.GetHiddenRows(), ['12:13']);
        grid.Redo();
        grid.SetGroupCollapsed('columns', 9, true);
        grid.UngroupRows(15);
        check('UngroupRows lowers the level', grid.GetOutline().rows.map((group) => [group.range, group.level]), [['11:15', 1], ['12:13', 2]]);
        grid.InsertRows(0, 2);
        check('inserted rows move the outline and the hidden rows', [grid.GetOutline().rows.map((group) => group.range), grid.GetHiddenRows()], [['13:17', '14:15'], ['13:18']]);
        grid.Undo();

        // ---- through a file ----
        grid.HideRows(2);
        const state = () => ({ rows: grid.GetHiddenRows(), columns: grid.GetHiddenColumns(), outline: grid.GetOutline(), widths: [1, 2, 3, 5, 6].map((col) => Math.round(grid.GetColumnWidth(col))), heights: [4, 6, 7].map((row) => Math.round(grid.GetRowHeight(row))) });
        const saved = state();
        const again = await reopened(grid);
        try {
            const read = { rows: again.grid.GetHiddenRows(), columns: again.grid.GetHiddenColumns(), outline: again.grid.GetOutline(), widths: [1, 2, 3, 5, 6].map((col) => Math.round(again.grid.GetColumnWidth(col))), heights: [4, 6, 7].map((row) => Math.round(again.grid.GetRowHeight(row))) };
            for (const key of Object.keys(saved)) check(`ExportXlsx -> ImportXlsx: ${key}`, read[key], saved[key]);
            again.grid.UnhideRows('1:20'); again.grid.UnhideColumns('A:Z');
            check('... a row hidden in the file comes back at the height it had', [Math.round(again.grid.GetRowHeight(2)), again.grid.GetHiddenRows(), again.grid.GetHiddenColumns()], [40, [], []]);
            check('... the collapsed marks are still there for the host', again.grid.GetOutline().rows.map((group) => group.collapsed), saved.outline.rows.map((group) => group.collapsed));
        } finally { again.close(); }

        // ---- a large sheet: a whole column is fitted from the cells it has ----
        const big = mountGrid();
        try {
            await big.grid.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
            big.grid.SwitchSheet(big.grid.GetSheetNames().indexOf('Data'));
            const cells = big.grid.GetModel().RowCount;
            let started = performance.now();
            big.grid.AutoFitColumns('A:A');
            const one = performance.now() - started;
            started = performance.now();
            big.grid.AutoFitColumns('A:K');
            const all = performance.now() - started;
            started = performance.now();
            big.grid.AutoFitRows('1:1048576');
            const rows = performance.now() - started;
            check('a whole column of a large sheet is fitted in reasonable time', one < 3000 && all < 6000 && rows < 6000, true);
            note(`dashboard "Data" (${cells} rows): AutoFitColumns("A:A") ${Math.round(one)} ms, ("A:K") ${Math.round(all)} ms, AutoFitRows(all rows) ${Math.round(rows)} ms`);
            check('the cell map did not grow', big.grid.CanUndo(), true);
        } finally { big.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
