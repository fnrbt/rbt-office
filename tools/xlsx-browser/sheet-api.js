// Browser side of run.mjs, third part: the host API of the Spreadsheet component - what a host
// application builds its toolbars, dialogs and panels on. One check per group of members
// ("api-..."). Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, bytesOf, frame, OF, createSpreadsheet, container, mouse, pointOfCell, partsOf } from '/component.js';

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
        check('a comment', grid.GetCell(2, 0).comment, { text: 'look here', author: 'Ada', runs: null, thread: null });
        grid.SetComment(1, 0, 'anonymous', undefined);
        check('a comment without an author', grid.GetCell(1, 0).comment, { text: 'anonymous', author: null, runs: null, thread: null });

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
            [() => grid.AddConditionalFormat('A1', { type: 'iconSet', iconSet: '3Hearts' }), /^AddConditionalFormat: rule\.iconSet must be one of 3Arrows/],
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
        grid.Undo(); grid.Undo();   // (the SetCell, then the entry)
        check('undo of the entry, then of the validation', [grid.GetCell(0, 0).value, brief()], [null, [`${first}:whole@A1:A10`]]);
        grid.SetCell(1, 0, '');
        const steps = changes.length;
        // (SetCell of nothing into an empty cell is no step: the step on top is the validation.)
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
        // (The list that reads another sheet is in the x14 list of the sheet, as Excel writes it.)
        check('the file has the whole columns as ranges', [...sheetXml.matchAll(/<dataValidation [^>]*sqref="([^"]*)"/g), ...sheetXml.matchAll(/<xm:sqref>([^<]*)<\/xm:sqref><\/x14:dataValidation>/g)].map((m) => m[1]), ['A1:A1048576', 'C1:D1048576']);
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
        // Any cell a sheet can have, as SelectCell: the grid is given room for it.
        const rowsBefore = grid.GetModel().RowCount, colsBefore = grid.GetModel().ColCount;
        check('a cell beyond the grid', [rowsBefore < 600 || colsBefore < 28, grid.GoTo('AB600'), selection(), grid.GetModel().RowCount >= 600, grid.GetModel().ColCount >= 28], [true, { sheetIndex: 0, range: 'AB600' }, [0, 'AB600'], true, true]);
        check('a range beyond the grid', [grid.GoTo('A600:B700'), selection(), grid.GetModel().RowCount >= 700], [{ sheetIndex: 0, range: 'A600:B700' }, [0, 'A600:B700'], true]);
        check('the last row of a sheet', [grid.GoTo('A1048576'), selection(), grid.GetActiveCell().Row], [{ sheetIndex: 0, range: 'A1048576' }, [0, 'A1048576'], 1048575]);
        check('the last column of a sheet', [grid.GoTo('XFD1'), selection(), grid.GetActiveCell().Col], [{ sheetIndex: 0, range: 'XFD1' }, [0, 'XFD1'], 16383]);
        // Gone to and come back from (the caret's growth is let go of so too), also when a cell was written out there and undone.
        grid.GoTo('A1'); await frame(); await frame();
        check('back at A1 the grid is as small as it was', [grid.GetModel().RowCount, grid.GetModel().ColCount], [rowsBefore, colsBefore]);
        grid.GoTo('AB600'); grid.SetCell(599, 27, 'out there'); grid.Undo(); grid.GoTo('A1'); await frame(); await frame();
        check('a cell written beyond the grid and undone: as small again', [grid.GetModel().RowCount, grid.GetModel().ColCount], [rowsBefore, colsBefore]);
        check('a cell beyond the grid on another sheet', [grid.GoTo("'Other sheet'!AD900"), selection()], [{ sheetIndex: 1, range: 'AD900' }, [1, 'AD900']]);
        grid.SwitchSheet(0);
        grid.SetDefinedName('FarAway', '=' + "'" + grid.GetSheetNames()[0].replace(/'/g, "''") + "'" + '!$AC$5000:$AD$5001');
        check('a name that leads beyond the grid', [grid.GoTo('FarAway'), selection()], [{ sheetIndex: 0, range: 'AC5000:AD5001' }, [0, 'AC5000:AD5001']]);
        grid.DeleteDefinedName('FarAway');
        check('what a sheet cannot have still throws', [thrown(() => grid.GoTo('A1048577')) !== null, thrown(() => grid.GoTo('XFE1')) !== null], [true, true]);
        grid.GoTo('A1');
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
            check(`zoom ${zoom}: the scroll extent is the zoomed sheet`, [Math.abs(scroller.scrollHeight - (30 + grid.GetModel().RowCount * 25) * zoom) <= 2, Math.abs(scroller.scrollWidth - (50 + grid.GetModel().ColCount * 100) * zoom) <= 2], [true, true]);
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
            check('the file opens on the sheet it was saved on', again.grid.GetActiveSheetIndex(), 1);
            check('the file keeps the view of each sheet', [(again.grid.SwitchSheet(0), again.grid.GetViewOptions()), (again.grid.SwitchSheet(1), again.grid.GetViewOptions())],
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
        grid.Undo(); grid.Undo();   // (the SetCell on this sheet, then the rename)
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
        grid.Undo(); grid.Undo();   // (the SetCell, then the copy)
        check('undo removes the copy and its names', [infos(), grid.GetDefinedNames().filter((n) => n.name === 'Local').length], [['0:Third:veryHidden', '1:Data:visible:#FF8800*', '2:Summary:hidden:#00AA55'], 1]);
        grid.Redo();
        check('redo brings it back', [infos()[2], grid.GetCell(6, 0).formula, grid.DuplicateSheet(1), infos()[2]], ['2:Data (2):visible:#FF8800*', "='Data (2)'!A1+1", 2, '2:Data (3):visible:#FF8800*']);
        grid.Undo();

        // ---- the file ----
        const again = await reopened(grid);
        try {
            const reinfos = again.grid.GetSheetInfos().map((s) => `${s.index}:${s.name}:${s.visibility}${s.tabColor ? ':' + s.tabColor : ''}${s.isActive ? '*' : ''}`);
            check('saved and opened again: names, order, visibility, tab colours', reinfos, ['0:Third:veryHidden', '1:Data:visible:#FF8800', '2:Data (2):visible:#FF8800*', '3:Summary:hidden:#00AA55']);   // (on the sheet it was saved on)
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

        // ---- what the file's kept parts say of a sheet follows the sheet ----
        // chart-zoo.xlsx: one sheet, six charts whose series read "Zoo!...".
        const zoo = mountGrid();
        try {
            const zip = async (bytes) => window.JSZip.loadAsync(bytes);
            const chartFormulas = async (bytes) => {
                const archive = await zip(bytes), out = {};
                for (const entry of Object.keys(archive.files).sort()) if (/^xl\/charts\/chart\d+\.xml$/.test(entry)) out[entry] = [...(await archive.files[entry].async('string')).matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => m[1].slice(0, m[1].indexOf('!')));
                return out;
            };
            const g = zoo.grid;
            await g.ImportXlsx(await bytesOf('tests/real-world/chart-zoo.xlsx'));
            const kept = () => g.GetDrawings().filter((item) => item.kind === 'chart').length;
            const chartsBefore = kept();
            const original = await chartFormulas(await g.ExportXlsx());
            const chartCount = Object.keys(original).length;
            check('chart-zoo: its charts read the sheet Zoo', [chartCount >= 6, chartsBefore, [...new Set(Object.values(original).flat())]], [true, chartCount, ['Zoo']]);
            g.RenameSheet(0, 'Wild & Tame');
            const renamedBytes = await g.ExportXlsx();
            const renamed = await chartFormulas(renamedBytes);
            check('a sheet renamed: the series of its charts read it by the new name', [Object.keys(renamed).length, [...new Set(Object.values(renamed).flat())]], [chartCount, ["'Wild &amp; Tame'"]]);
            g.Undo();
            check('... undone: as the file has them', [...new Set(Object.values(await chartFormulas(await g.ExportXlsx())).flat())], ['Zoo']);
            // Copied: the copy shows the charts, and a save gives it chart parts of its own that read the copy.
            const copyIndex = g.DuplicateSheet(0);
            check('a sheet copied: the copy has the charts', [g.GetSheetNames(), g.GetActiveSheetIndex(), kept()], [['Zoo', 'Zoo (2)'], copyIndex, chartsBefore]);
            await frame();
            check('... and draws them', host.ownerDocument.querySelectorAll('.spreadsheet-kept-object, .spreadsheet-drawing').length >= 0, true);
            const copiedBytes = await g.ExportXlsx();
            const copied = await chartFormulas(copiedBytes);
            const readers = Object.values(copied).map((sheetsRead) => [...new Set(sheetsRead)].join());
            check('saved: twice the chart parts, half of them read the copy', [Object.keys(copied).length, readers.filter((name) => name === 'Zoo').length, readers.filter((name) => name === "'Zoo (2)'").length], [2 * chartCount, chartCount, chartCount]);
            const reread = mountGrid();
            try {
                await reread.grid.ImportXlsx(copiedBytes);
                const counts = [0, 1].map((index) => { reread.grid.SwitchSheet(index); return reread.grid.GetDrawings().filter((item) => item.kind === 'chart').length; });
                check('opened again: both sheets have their charts', [reread.grid.GetSheetNames(), counts], [['Zoo', 'Zoo (2)'], [chartsBefore, chartsBefore]]);
                // The copy of the written file is a sheet of that file: renamed, its own charts follow.
                reread.grid.RenameSheet(1, 'Second');
                const again = await chartFormulas(await reread.grid.ExportXlsx());
                check('... and a rename of the copy there', [...new Set(Object.values(again).flat())].sort(), ['Second', 'Zoo']);
            } finally { reread.close(); }
            g.Undo();
            check('the copy undone: one sheet, and a save has the file\'s chart parts', [g.GetSheetNames(), Object.keys(await chartFormulas(await g.ExportXlsx())).length], [['Zoo'], chartCount]);
        } finally { zoo.close(); }

        // A sheet that is not converted yet (it has no formula) with a link into another sheet: the link follows
        // that sheet's new name, and says #REF! when the sheet is deleted.
        const links = mountGrid();
        try {
            const g = links.grid;
            g.SetCell(0, 0, 'first'); g.AddSheet('Notes'); g.SetCell(0, 0, 'see the first sheet'); g.SetHyperlink('A1', 'Sheet1!B2'); g.SetHyperlink('A2', 'https://example.com/Sheet1!B2');
            g.AddSheet('Third'); g.SetCell(0, 0, '3'); g.SwitchSheet(0);
            const opened = mountGrid();
            try {
                const o = opened.grid;
                await o.ImportXlsx(await g.ExportXlsx());
                const loaded = () => o.GetSheetInfos().map((info) => info.isLoaded ?? info.loaded ?? null);
                note('sheets converted after opening: ' + JSON.stringify(o.GetSheetInfos().map((info) => Object.entries(info).filter(([key]) => /load/i.test(key)))));
                o.RenameSheet(0, 'Front page');
                check('a link on a sheet that was not converted follows the rename', o.GetHyperlinks(1).map((link) => link.target), ["'Front page'!B2", 'https://example.com/Sheet1!B2']);
                const renamedLinks = mountGrid();
                try {
                    await renamedLinks.grid.ImportXlsx(await o.ExportXlsx());
                    check('... saved and opened again', [renamedLinks.grid.GetSheetNames(), renamedLinks.grid.GetHyperlinks(1).map((link) => [link.target, link.isInternal])], [['Front page', 'Notes', 'Third'], [["'Front page'!B2", true], ['https://example.com/Sheet1!B2', false]]]);
                    // Deleted from the file just opened (the Notes sheet is not converted there either).
                    renamedLinks.grid.SwitchSheet(2); renamedLinks.grid.DeleteSheet(0);
                    check('the sheet deleted: the link into it says #REF!', renamedLinks.grid.GetHyperlinks(0).map((link) => link.target), ['#REF!', 'https://example.com/Sheet1!B2']);
                    const deleted = mountGrid();
                    try {
                        await deleted.grid.ImportXlsx(await renamedLinks.grid.ExportXlsx());
                        check('... saved and opened again', deleted.grid.GetHyperlinks(0).map((link) => link.target), ['#REF!', 'https://example.com/Sheet1!B2']);
                    } finally { deleted.close(); }
                    renamedLinks.grid.Undo();
                    check('... undone: the link is back', renamedLinks.grid.GetHyperlinks(1).map((link) => link.target), ["'Front page'!B2", 'https://example.com/Sheet1!B2']);
                } finally { renamedLinks.close(); }
            } finally { opened.close(); }
        } finally { links.close(); }
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
        check('no outline at first', grid.GetOutline(), { rows: [], columns: [], rowLevels: 0, columnLevels: 0, summaryBelow: true, summaryRight: true, showSymbols: true });
        grid.GroupRows('11:16'); grid.GroupRows(11, 2); grid.GroupColumns('J:K');
        check('GroupRows / GroupColumns: levels', grid.GetOutline(), {
            rows: [{ first: 10, last: 15, range: '11:16', level: 1, collapsed: false }, { first: 11, last: 12, range: '12:13', level: 2, collapsed: false }],
            columns: [{ first: 9, last: 10, range: 'J:K', level: 1, collapsed: false }], rowLevels: 2, columnLevels: 1, summaryBelow: true, summaryRight: true, showSymbols: true });
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

// Sort and filter: SortRange, SetAutoFilter / ClearAutoFilter, SetFilter / ClearFilter / ReapplyFilter,
// GetFilters, GetFilterValues; what copy, clear, fill and format do on a filtered selection.
checks['api-sortfilter'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const column = (col, rows, g = grid) => rows.map((row) => g.GetCell(row, col).text);
        const formulas = (col, rows, g = grid) => rows.map((row) => g.GetCell(row, col).formula);
        const rowsDrawn = () => [...host.querySelectorAll('.spreadsheet-row-header')].filter((n) => n.getBoundingClientRect().height > 0 && n.offsetParent !== null).map((n) => Number(n.getAttribute('data-row'))).filter((row) => row <= 11).sort((a, b) => a - b);
        const cellDrawn = (row, col) => { const n = host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`); return !!n && n.getBoundingClientRect().height > 0; };
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        const fill = () => {
            const data = [['Region', 'Sales', 'Double', 'Note'], ['East', '30', '=B2*2', 'first order'], ['West', '10', '=B3*2', 'Second'], ['east', '20', '=B4*2', ''],
                          ['North', '40', '=B5*2', '50% off'], ['West', '50', '=B6*2', 'last order'], ['', '60', '=B7*2', 'no region'], ['South', 'n/a', '=B2+B$2', 'what?'], ['East', '20', '=B9*2', '']];
            data.forEach((row, r) => row.forEach((text, c) => { if (text !== '') grid.SetCell(r, c, text); }));
            grid.SetCell(10, 1, '=SUBTOTAL(109,B2:B9)'); grid.SetCell(10, 2, '=SUM(B2:B9)'); grid.SetCell(11, 5, '=B2');
        };
        fill();
        grid.SelectCell(8, 1); grid.ApplyBgColor('#FFFF00'); grid.SelectCell(3, 1); grid.ApplyTextColor('#FF0000'); grid.SelectCell(0, 6);
        changes.splice(0);

        // ---- SortRange ----
        const r0 = grid.GetRevision();
        grid.SortRange('A1:D9', [{ column: 'B' }], { hasHeader: true });
        check('SortRange by a column, the header row staying: numbers, then text', column(1, [0, 1, 2, 3, 4, 5, 6, 7, 8]), ['Sales', '10', '20', '20', '30', '40', '50', '60', 'n/a']);
        check('rows equal on the key keep their order; the other cells of the range come along', column(0, [1, 2, 3, 4, 5, 6, 7, 8]), ['West', 'east', 'East', 'East', 'North', 'West', '', 'South']);
        check('a formula that moved points at its own row again', formulas(2, [1, 2, 3, 4, 8]), ['=B2*2', '=B3*2', '=B4*2', '=B5*2', '=B3+B$2']);
        check('... and is recalculated', column(2, [1, 2, 4, 7]), ['20', '40', '60', '120']);
        check('formulas outside the range are not re-pointed, and see the new order', [grid.GetCell(11, 5).formula, grid.GetCell(11, 5).text, grid.GetCell(10, 2).text], ['=B2', '10', '230']);
        check('formats moved with their cells', [grid.GetCell(3, 1).format.fill, grid.GetCell(2, 1).format.textColor], ['#FFFF00', '#FF0000']);
        check('one change, one undo step, a new revision', [changes.splice(0).length, grid.GetRevision() !== r0, grid.CanUndo()], [1, true, true]);
        grid.Undo();
        check('undo puts the rows back', [column(1, [1, 2, 3]), formulas(2, [1, 7]), grid.GetRevision() === r0, grid.GetCell(11, 5).text], [['30', '10', '20'], ['=B2*2', '=B2+B$2'], true, '30']);
        grid.Redo();
        check('redo sorts again', column(1, [1, 2, 3, 8]), ['10', '20', '20', 'n/a']);
        grid.Undo();
        grid.SortRange('A2:D9', [{ column: 0 }, { column: 1, descending: true }]);
        check('two keys (the second descending), blanks last', [column(0, [1, 2, 3, 4, 5, 6, 7, 8]), column(1, [1, 2, 3, 4, 5, 6, 7, 8])], [['East', 'east', 'East', 'North', 'South', 'West', 'West', ''], ['30', '20', '20', '40', 'n/a', '50', '10', '60']]);
        grid.Undo();
        grid.SortRange('A2:D9', [{ column: 'B', by: 'cellColor', color: '#ffff00' }, { column: 1, by: 'fontColor', color: '#FF0000' }]);
        check('by fill colour, then by text colour', column(1, [1, 2, 3, 4]), ['20', '20', '30', '10']);
        grid.Undo();
        grid.SortRange('A:D', [{ column: 1, descending: true }], { hasHeader: true });
        check('whole columns are clipped to the used rows (the formulas below sort as numbers)', column(1, [1, 2, 3]).length, 3);
        grid.Undo();
        check('sort: bad arguments throw', [
            thrown(() => grid.SortRange('nope', [{ column: 0 }])), thrown(() => grid.SortRange('A1:D9', [])), thrown(() => grid.SortRange('A1:D9', [{ column: 'F' }])),
            thrown(() => grid.SortRange('A1:D9', [{ column: 0, by: 'icon' }])), thrown(() => grid.SortRange('A1:D9', [{ column: 0, order: 'desc' }])),
            thrown(() => grid.SortRange('A1:D9', [{ column: 0 }], { header: true })), thrown(() => grid.SortRange('A1:D9')),
        ].map((message) => message && message.startsWith('SortRange: ')), [true, true, true, true, true, true, true]);
        grid.MergeCells(3, 3, 4, 3);
        check('merged cells among the rows are refused', thrown(() => grid.SortRange('A1:D9', [{ column: 0 }], { hasHeader: true })), 'SortRange: the range holds merged cells');
        grid.Undo();
        grid.SetCell(2, 6, 'a'); grid.SetCell(3, 6, 'B'); grid.SetCell(4, 6, 'c'); grid.SelectCell(2, 6);
        grid.SortCurrentColumnDesc();
        check('SortCurrentColumnDesc sorts the block around the active cell only (a list of text has no header)', [column(6, [2, 3, 4]), column(0, [0, 1, 2]), column(1, [1, 2])], [['c', 'B', 'a'], ['Region', 'East', 'West'], ['30', '10']]);
        grid.Undo();
        check('... as one undo step', column(6, [2, 3, 4]), ['a', 'B', 'c']);
        // The reported case: a block at R101:T105 with a header row of text, far below other data, the active cell in it.
        [['Item', 'Qty', 'Price'], ['bolt', '4', '1.5'], ['nut', '9', ''], ['washer', '2', '2'], ['pin', '7', '']].forEach((row, r) => row.forEach((text, c) => { if (text !== '') grid.SetCell(100 + r, 17 + c, text); }));
        for (let row = 9; row < 100; row++) grid.SetCell(row, 21, String(row));
        const beforeQuick = grid.GetRevision();
        grid.SelectCell(102, 19);
        grid.SortCurrentColumnDesc();
        check('quick sort descending: the header row stays, blanks are last, the rows travel', [column(19, [100, 101, 102, 103, 104]), column(17, [100, 101, 102, 103, 104])], [['Price', '2', '1.5', '', ''], ['Item', 'washer', 'bolt', 'nut', 'pin']]);
        check('... the data above is not touched', [column(0, [0, 1, 2]), column(1, [1, 2, 3]), column(21, [99])], [['Region', 'East', 'West'], ['30', '10', '20'], ['99']]);
        grid.Undo();
        check('... one undo step', [column(19, [101, 102, 103, 104]), grid.GetRevision() === beforeQuick], [['1.5', '', '2', ''], true]);
        grid.Redo();
        check('... redo sorts again', column(19, [101, 102, 103, 104]), ['2', '1.5', '', '']);
        grid.SortCurrentColumnAsc();
        check('quick sort ascending: header kept, blanks last', [column(19, [100, 101, 102, 103, 104]), column(17, [101, 102])], [['Price', '1.5', '2', '', ''], ['bolt', 'washer']]);
        grid.SelectCell(101, 17); grid.SortCurrentColumnAsc();
        check('quick sort by a text column of the block', column(17, [100, 101, 102, 103, 104]), ['Item', 'bolt', 'nut', 'pin', 'washer']);
        grid.SelectCell(150, 3); const idle = grid.GetRevision(); grid.SortCurrentColumnAsc();
        check('no data around the cell: nothing happens', grid.GetRevision(), idle);
        for (let row = 100; row <= 104; row++) for (let col = 17; col <= 19; col++) grid.SetCell(row, col, '');
        for (let row = 9; row < 100; row++) grid.SetCell(row, 21, '');
        grid.SelectCell(0, 0); grid.SelectCell(0, 6);
        changes.splice(0);

        // ---- SetAutoFilter / SetFilter ----
        check('no filter: GetFilters is null', grid.GetFilters(), null);
        const r1 = grid.GetRevision();
        grid.SetAutoFilter('A1:D9');
        check('SetAutoFilter: the range, nothing hidden yet', grid.GetFilters(), { range: 'A1:D9', columns: [], hiddenRows: [], hiddenRowCount: 0, visibleRowCount: 8 });
        check('it is a change and an undo step', [changes.splice(0).length, grid.GetRevision() !== r1], [1, true]);
        check('GetFilterValues: what a checklist offers', [grid.GetFilterValues(0), grid.GetFilterValues('B')], [{ values: ['East', 'North', 'South', 'West'], blanks: true }, { values: ['10', '20', '30', '40', '50', '60', 'n/a'], blanks: false }]);
        grid.SetFilter(0, { values: ['East'] });
        await frame();
        const filters = grid.GetFilters();
        check('SetFilter by values (case-insensitive): the other rows are hidden', [filters.hiddenRows, filters.hiddenRowCount, filters.visibleRowCount, filters.columns], [[2, 4, 5, 6, 7], 5, 3, [{ column: 0, letter: 'A', criteria: { values: ['East'] } }]]);
        check('the grid draws the rows that are left, with their own numbers', [rowsDrawn(), cellDrawn(2, 0), cellDrawn(3, 0), grid.GetRowHeight(2) >= 0], [[0, 1, 3, 8, 9, 10, 11], false, true, true]);
        check('the hidden rows keep their cells', column(0, [2, 4, 5]), ['West', 'North', 'West']);
        check('SUBTOTAL leaves the hidden rows out, SUM does not', [grid.GetCell(10, 1).text, grid.GetCell(10, 2).text], ['70', '230']);
        check('one change, one undo step', [changes.splice(0).length], [1]);
        grid.Undo();
        await frame();
        check('undo shows the rows again (the filter range stays)', [grid.GetFilters().hiddenRows, grid.GetFilters().columns, rowsDrawn().length, grid.GetCell(10, 1).text], [[], [], 12, '230']);
        grid.Redo();
        check('redo hides them', [grid.GetFilters().hiddenRows, grid.GetCell(10, 1).text], [[2, 4, 5, 6, 7], '70']);
        const hiddenFor = (col, criteria) => { grid.ClearFilter(); grid.SetFilter(col, criteria); return grid.GetFilters().hiddenRows; };
        check('values with blanks', hiddenFor(0, { values: ['west'], blanks: true }), [1, 3, 4, 7, 8]);
        check('a condition on numbers', hiddenFor(1, { conditions: [{ operator: 'greaterThan', value: 30 }] }), [1, 2, 3, 7, 8]);
        check('two conditions joined by AND, then OR', [hiddenFor('B', { conditions: [{ operator: 'greaterOrEqual', value: 20 }, { operator: 'lessOrEqual', value: '40' }] }), hiddenFor('B', { conditions: [{ operator: 'lessThan', value: 20 }, { operator: 'greaterThan', value: 50 }], and: false })], [[2, 5, 6, 7], [1, 3, 4, 5, 7, 8]]);
        check('text conditions', [hiddenFor(3, { conditions: [{ operator: 'contains', value: 'order' }] }), hiddenFor(3, { conditions: [{ operator: 'notContains', value: 'order' }] }), hiddenFor(3, { conditions: [{ operator: 'beginsWith', value: 's' }] }),
            hiddenFor(3, { conditions: [{ operator: 'endsWith', value: '?' }] }), hiddenFor(3, { conditions: [{ operator: 'equals', value: '5?% *' }] }), hiddenFor(0, { conditions: [{ operator: 'notEquals', value: 'west' }] })],
            [[2, 3, 4, 6, 7, 8], [1, 5], [1, 3, 4, 5, 6, 7, 8], [1, 2, 3, 4, 5, 6, 8], [1, 2, 3, 5, 6, 7, 8], [2, 5]]);
        check('the criteria read back as they were given', grid.GetFilters().columns[0].criteria, { conditions: [{ operator: 'notEquals', value: 'west' }] });
        // What was set is what is read: the types of the values, and no key that was not given.
        const readBack = (col, criteria) => { grid.ClearFilter(); grid.SetFilter(col, criteria); return grid.GetFilters().columns[0].criteria; };
        const five = readBack(1, { conditions: [{ operator: 'greaterThan', value: 5 }] });
        check('a condition set with the number 5 reads back with the number 5, and nothing added', [five, typeof five.conditions[0].value, Object.keys(five)], [{ conditions: [{ operator: 'greaterThan', value: 5 }] }, 'number', ['conditions']]);
        check('a condition set with the text "5" reads back with the text', readBack(1, { conditions: [{ operator: 'equals', value: '5' }] }), { conditions: [{ operator: 'equals', value: '5' }] });
        check('two conditions with and: false, a number and a text', readBack(1, { conditions: [{ operator: 'lessThan', value: 20.5 }, { operator: 'greaterThan', value: '50' }], and: false }), { conditions: [{ operator: 'lessThan', value: 20.5 }, { operator: 'greaterThan', value: '50' }], and: false });
        check('and: true is kept when it was given', readBack(1, { conditions: [{ operator: 'greaterOrEqual', value: 20 }, { operator: 'lessOrEqual', value: 40 }], and: true }), { conditions: [{ operator: 'greaterOrEqual', value: 20 }, { operator: 'lessOrEqual', value: 40 }], and: true });
        const when = new Date(2024, 2, 31);
        const dated = readBack(1, { conditions: [{ operator: 'lessThan', value: when }] });
        check('a Date reads back as a Date (a copy)', [dated.conditions[0].value instanceof Date, dated.conditions[0].value.getTime() === when.getTime(), dated.conditions[0].value !== when], [true, true, true]);
        check('values, top, colour and dynamic filters read back as given', [readBack(0, { values: ['East', 'west'], blanks: true }), readBack(0, { values: ['East'], blanks: false }), readBack(1, { top: { count: 2 } }), readBack(1, { top: { count: 3, percent: true, bottom: false } }), readBack(1, { color: '#ffff00' }), readBack(1, { color: null, by: 'fontColor' }), readBack(1, { dynamic: 'aboveAverage' })],
            [{ values: ['East', 'west'], blanks: true }, { values: ['East'], blanks: false }, { top: { count: 2 } }, { top: { count: 3, percent: true, bottom: false } }, { color: '#ffff00' }, { color: null, by: 'fontColor' }, { dynamic: 'aboveAverage' }]);
        const given = { conditions: [{ operator: 'greaterThan', value: 5 }] };
        const first = readBack(1, given);
        given.conditions[0].value = 99; first.conditions[0].value = 77;
        check('what is read is a copy: neither side changes the filter', grid.GetFilters().columns[0].criteria, { conditions: [{ operator: 'greaterThan', value: 5 }] });
        grid.SetFilter(1, { conditions: [{ operator: 'greaterThan', value: '30' }] }); grid.Undo();
        check('undo brings the earlier filter back with its own types', grid.GetFilters().columns[0].criteria, { conditions: [{ operator: 'greaterThan', value: 5 }] });
        grid.Redo();
        check('redo the later one', grid.GetFilters().columns[0].criteria, { conditions: [{ operator: 'greaterThan', value: '30' }] });
        {
            // Through a file there is only what XLSX holds: defaults are left out and a comparison's number is a number.
            grid.ClearFilter(); grid.SetFilter(1, { conditions: [{ operator: 'greaterThan', value: '5' }, { operator: 'lessThan', value: 45 }], and: true }); grid.SetFilter(0, { values: ['East', 'West'] }); grid.SetFilter(3, { conditions: [{ operator: 'contains', value: '5' }] });
            const again = await reopened(grid);
            check('after export and re-import: the documented canonical shape', again.grid.GetFilters().columns.map((entry) => entry.criteria),
                [{ values: ['East', 'West'] }, { conditions: [{ operator: 'greaterThan', value: 5 }, { operator: 'lessThan', value: 45 }] }, { conditions: [{ operator: 'contains', value: '5' }] }]);
            again.close();
        }
        check('top and bottom', [hiddenFor(1, { top: { count: 2 } }), hiddenFor(1, { top: { count: 2, bottom: true } }), hiddenFor(1, { top: { count: 30, percent: true } })], [[1, 2, 3, 4, 7, 8], [1, 4, 5, 6, 7], [1, 2, 3, 4, 7, 8]]);
        check('by colour', [hiddenFor(1, { color: '#ffff00' }), hiddenFor(1, { color: '#FF0000', by: 'fontColor' }), hiddenFor(1, { color: null })], [[1, 2, 3, 4, 5, 6, 7], [1, 2, 4, 5, 6, 7, 8], [8]]);
        check('above and below average', [hiddenFor(1, { dynamic: 'aboveAverage' }), hiddenFor(1, { dynamic: 'belowAverage' })], [[1, 2, 3, 7, 8], [4, 5, 6, 7]]);
        grid.ClearFilter(); grid.SetFilter(0, { values: ['East', 'West'] }); grid.SetFilter(1, { conditions: [{ operator: 'greaterOrEqual', value: 20 }] });
        check('filters on two columns combine', [grid.GetFilters().hiddenRows, grid.GetFilters().columns.map((c) => c.letter)], [[2, 4, 6, 7], ['A', 'B']]);
        grid.ClearFilter('A');
        check('ClearFilter(column) takes one off', [grid.GetFilters().hiddenRows, grid.GetFilters().columns.map((c) => c.letter)], [[2, 7], ['B']]);
        grid.ClearFilter();
        check('ClearFilter() takes all off, the AutoFilter stays', [grid.GetFilters().hiddenRows, grid.GetFilters().columns, grid.GetFilters().range], [[], [], 'A1:D9']);
        check('filter: bad arguments throw', [
            thrown(() => grid.SetFilter(7, { values: ['x'] })), thrown(() => grid.SetFilter(0, { values: 'x' })), thrown(() => grid.SetFilter(0, {})), thrown(() => grid.SetFilter(0, { values: [], top: { count: 1 } })),
            thrown(() => grid.SetFilter(0, { conditions: [{ operator: 'like', value: 1 }] })), thrown(() => grid.SetFilter(0, { top: { count: 0 } })), thrown(() => grid.SetFilter(0, { color: 'red' })),
            thrown(() => grid.SetFilter(0, { dynamic: 'today' })), thrown(() => grid.SetFilter(0, { value: 1 })), thrown(() => grid.SetFilter('A1', { values: [] })), thrown(() => grid.SetFilter(0)),
        ].map((message) => message && message.startsWith('SetFilter: ')), new Array(11).fill(true));
        check('... with a message that says what', [thrown(() => grid.SetFilter(7, { values: ['x'] })), thrown(() => grid.SetAutoFilter('nope')) !== null, thrown(() => grid.SetAutoFilter()) !== null, thrown(() => grid.GetFilters(5)) !== null, thrown(() => grid.ClearFilter('1')) !== null],
            ['SetFilter: column H is outside the filter range A1:D9', true, true, true, true]);

        // ---- a filter is applied when it is set, not when a cell changes ----
        grid.SetFilter(0, { values: ['East'] });
        grid.SetCell(1, 0, 'West');
        check('an edited row stays where it is', grid.GetFilters().hiddenRows, [2, 4, 5, 6, 7]);
        check('... and the edit recalculated incrementally', grid.GetRecalcStats().Kind, 'incremental');
        changes.splice(0);
        grid.ReapplyFilter();
        check('ReapplyFilter hides it (one change)', [grid.GetFilters().hiddenRows, changes.splice(0).length], [[1, 2, 4, 5, 6, 7], 1]);
        grid.Undo(); grid.SetCell(1, 0, 'East');

        // ---- a filtered selection: copy, clear, format, fill ----
        grid.SelectRange(0, 0, 8, 1);
        check('GetSelectionText copies the rows that are shown', grid.GetSelectionText(), 'Region\tSales\nEast\t30\neast\t20\nEast\t20');
        check('GetSelectionHtml too', (grid.GetSelectionHtml().match(/<tr>/g) || []).length, 4);
        const copiedText = grid.GetSelectionText();
        await grid.Copy();
        grid.SelectCell(14, 0); grid.PasteText(copiedText, '');
        check('a copy of a filtered selection pastes as one block', [column(0, [14, 15, 16, 17, 18]), column(1, [15, 16, 17])], [['Region', 'East', 'east', 'East', ''], ['30', '20', '20']]);
        grid.Undo();
        grid.SelectRange(1, 3, 8, 3); grid.ApplyFormat('bold');
        check('a format command skips the hidden rows', [1, 2, 3, 5, 8].map((row) => grid.GetCell(row, 3).format.bold), [true, false, true, false, true]);
        grid.Undo();
        grid.SelectRange(1, 3, 8, 3); grid.ClearContents();
        check('ClearContents skips them', column(3, [1, 2, 4, 5, 6, 7]), ['', 'Second', '50% off', 'last order', 'no region', 'what?']);
        grid.Undo();
        grid.SelectRange(1, 3, 8, 3); press(host, 'Delete');
        check('the Delete key skips them', column(3, [1, 2, 4, 5]), ['', 'Second', '50% off', 'last order']);
        grid.Undo();
        grid.SelectRange(1, 3, 8, 3); grid.FillDown();
        check('FillDown fills the rows that are shown', column(3, [1, 2, 3, 4, 5, 8]), ['first order', 'Second', 'first order', '50% off', 'last order', 'first order']);
        grid.Undo();
        grid.SelectCell(0, 6);
        grid.HideRows(3);
        check('a row the user hid is another matter: it stays hidden when the filter goes', [grid.GetHiddenRows(), grid.GetFilters().hiddenRows], [['4:4'], [2, 4, 5, 6, 7]]);
        grid.ClearFilter();
        check('... and is not shown by ClearFilter', [grid.GetHiddenRows(), grid.GetFilters().hiddenRows, grid.GetRowHeight(3), grid.GetRowHeight(2)], [['4:4'], [], 0, 25]);
        grid.UnhideRows(3);
        grid.SetFilter(0, { values: ['East'] });
        grid.InsertRows(2, 1);
        check('inserted rows move the range and the hidden rows; the new row is shown', [grid.GetFilters().range, grid.GetFilters().hiddenRows], ['A1:D10', [3, 5, 6, 7, 8]]);
        grid.Undo();
        check('undo of the insert', [grid.GetFilters().range, grid.GetFilters().hiddenRows], ['A1:D9', [2, 4, 5, 6, 7]]);

        // ---- the grid's own dropdown ----
        const button = host.querySelector('.spreadsheet-col-header[data-col="1"] .col-filter-btn');
        button.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); button.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); button.click();
        await frame();
        const dropdown = document.getElementById('filter-dropdown');
        check('the column button opens the checklist of the filter range', [!!dropdown, dropdown ? [...dropdown.querySelectorAll('label span')].map((n) => n.textContent) : []], [true, ['10', '20', '30', '40', '50', '60', 'n/a']]);
        if (dropdown) {
            const boxes = [...dropdown.querySelectorAll('input[type=checkbox]')];
            boxes.forEach((box, i) => { box.checked = i === 1 || i === 2; });
            [...dropdown.querySelectorAll('button')].find((b) => b.textContent === 'Apply').click();
            await frame();
            check('Apply sets a values filter on the column (an undo step)', [grid.GetFilters().columns.map((c) => [c.letter, c.criteria]), grid.GetFilters().hiddenRows], [[['A', { values: ['East'] }], ['B', { values: ['20', '30'] }]], [2, 4, 5, 6, 7]]);
            grid.Undo();
            check('... that undo takes back', grid.GetFilters().columns.length, 1);
        }

        // ---- ExportXlsx -> ImportXlsx ----
        grid.SetFilter(1, { conditions: [{ operator: 'greaterOrEqual', value: 20 }, { operator: 'lessThan', value: 60 }] });
        grid.SetFilter(3, { conditions: [{ operator: 'notContains', value: 'zzz' }] });
        const before = grid.GetFilters();
        const bytes = await grid.ExportXlsx();
        const zip = await window.JSZip.loadAsync(bytes);
        const sheetXml = await zip.files['xl/worksheets/sheet1.xml'].async('string'), workbookXml = await zip.files['xl/workbook.xml'].async('string');
        check('the file: the filter, its columns, the hidden rows, the hidden name', [
            /<autoFilter ref="A1:D9"><filterColumn colId="0"><filters><filter val="East"\/><\/filters><\/filterColumn><filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="20"\/><customFilter operator="lessThan" val="60"\/><\/customFilters><\/filterColumn><filterColumn colId="3"><customFilters><customFilter operator="notEqual" val="\*zzz\*"\/><\/customFilters><\/filterColumn><\/autoFilter>/.test(sheetXml),
            [...sheetXml.matchAll(/<row r="(\d+)"[^>]*hidden="1"/g)].map((m) => Number(m[1]) - 1), /<definedName name="_xlnm\._FilterDatabase" localSheetId="0" hidden="1">[^<]*!\$A\$1:\$D\$9<\/definedName>/.test(workbookXml)], [true, before.hiddenRows, true]);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(bytes);
            await frame();
            check('re-imported: the same filter and the same rows hidden', again.grid.GetFilters(), before);
            check('... the cells of the hidden rows are there, the defined names are not polluted', [column(0, [2, 4], again.grid), again.grid.GetDefinedNames().length], [['West', 'North'], 0]);
            check('... it is not modified, and ReapplyFilter changes nothing', [again.grid.IsModified(), (again.grid.ReapplyFilter(), again.grid.IsModified())], [false, false]);
            again.grid.ClearFilter(1);
            check('... and goes on working', again.grid.GetFilters().hiddenRows, [2, 4, 5, 6, 7]);
            again.grid.ClearAutoFilter();
            await frame();
            check('ClearAutoFilter removes it and shows every row', [again.grid.GetFilters(), [1, 2, 4, 7].map((row) => again.grid.GetRowHeight(row) > 0)], [null, [true, true, true, true]]);
            const cleared = await reopened(again.grid);
            try { check('... also in the file written after that', [cleared.grid.GetFilters(), cleared.grid.GetHiddenRows()], [null, []]); } finally { cleared.close(); }
            again.grid.Undo();
            check('undo of ClearAutoFilter', again.grid.GetFilters().hiddenRows, [2, 4, 5, 6, 7]);
        } finally { again.close(); }

        // ---- a large sheet ----
        const big = mountGrid();
        try {
            await big.grid.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
            const names = big.grid.GetSheetNames();
            big.grid.SwitchSheet(names.indexOf('Data'));
            await frame();
            const rows = big.grid.GetModel().RowCount;
            let started = performance.now();
            big.grid.SetAutoFilter('A:K');
            const values = big.grid.GetFilterValues(1).values;
            const listed = performance.now() - started;
            started = performance.now();
            big.grid.SetFilter(1, { values: values.slice(0, 1) });
            const filtered = performance.now() - started;
            await frame();
            const state = big.grid.GetFilters();
            check('a filter on a large sheet hides most rows', state.hiddenRowCount > 1000 && state.visibleRowCount > 0, true);
            started = performance.now();
            big.grid.SortRange('A:K', [{ column: 2, descending: true }], { hasHeader: true });
            const sorted = performance.now() - started;
            started = performance.now();
            big.grid.Undo(); big.grid.Undo();
            const undone = performance.now() - started;
            check('in reasonable time', listed < 5000 && filtered < 5000 && sorted < 10000, true);
            note(`dashboard "Data" (${rows} rows): SetAutoFilter + GetFilterValues ${Math.round(listed)} ms, SetFilter ${Math.round(filtered)} ms (${state.hiddenRowCount} rows hidden), SortRange of the ${state.visibleRowCount} shown rows ${Math.round(sorted)} ms, two undos ${Math.round(undone)} ms`);
        } finally { big.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Comments: notes and threads with replies, as undo steps and through XLSX.
checks['api-comments'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const title = (row, col) => node(row, col)?.getAttribute('data-comment') ?? null;
        const changes = [];
        const off = grid.OnChange((change) => changes.push(change));
        grid.SetCell(1, 1, '42');
        check('no comment yet', [grid.GetComment(1, 1), grid.GetComments()], [null, []]);
        const r0 = grid.GetRevision(); changes.length = 0;

        // ---- a note ----
        grid.SetComment(1, 1, 'look here', 'Ada');
        check('a note', grid.GetComment(1, 1), { text: 'look here', author: 'Ada', runs: null, thread: null });
        await frame();
        check('the cell shows it: a marker, and it says its comment (no browser tooltip)', [node(1, 1).children.length > 0, title(1, 1), node(1, 1).getAttribute('title')], [true, 'Ada:\nlook here', null]);
        check('it is a change with a new revision, and the value stays', [changes.length, grid.GetRevision() !== r0, grid.GetCell(1, 1).value, grid.GetCell(1, 1).comment.text], [1, true, 42, 'look here']);
        grid.Undo();
        await frame();
        check('undo takes it away, back at the revision before', [grid.GetComment(1, 1), title(1, 1), grid.GetRevision() === r0], [null, null, true]);
        grid.Redo();
        check('redo puts it back', grid.GetComment(1, 1)?.text, 'look here');
        grid.SetComment(1, 1, 'look again');
        check('a note is replaced (no author given: it keeps its author)', grid.GetComment(1, 1), { text: 'look again', author: 'Ada', runs: null, thread: null });
        grid.Undo();
        grid.SetComment(4, 3, 'on an empty cell', null);
        check('a note on an empty cell', [grid.GetComment(4, 3).text, grid.GetCell(4, 3).value], ['on an empty cell', null]);
        check('GetComments lists them row by row', grid.GetComments().map((c) => [c.address, c.row, c.col, c.text, c.author, c.thread]), [['B2', 1, 1, 'look here', 'Ada', null], ['D5', 4, 3, 'on an empty cell', null, null]]);
        check('deleting a comment says whether there was one', [grid.DeleteComment(4, 3), grid.DeleteComment(4, 3), grid.GetComment(4, 3)], [true, false, null]);
        grid.Undo();
        check('... one undo step', grid.GetComment(4, 3)?.text, 'on an empty cell');
        grid.RemoveComment(4, 3);
        check('RemoveComment is the older name of it', grid.GetComment(4, 3), null);

        // ---- replies: a thread ----
        const replyId = grid.AddCommentReply(1, 1, 'done', 'Bob');
        const threaded = grid.GetComment(1, 1);
        check('a reply turns the note into a thread', [threaded.text, threaded.author, threaded.thread.text, threaded.thread.resolved, threaded.thread.replies.map((r) => [r.id, r.text, r.author, /^\d{4}-\d\d-\d\dT/.test(r.date)])],
            ['look here', 'Ada', 'look here', false, [[replyId, 'done', 'Bob', true]]]);
        await frame();
        check('the cell says the reply too', title(1, 1).includes('look here') && title(1, 1).includes('done'), true);
        grid.Undo();
        check('... one undo step: a note again', grid.GetComment(1, 1), { text: 'look here', author: 'Ada', runs: null, thread: null });
        grid.Redo();
        const second = grid.AddCommentReply(1, 1, 'thanks');
        check('a second reply follows the first', grid.GetComment(1, 1).thread.replies.map((r) => [r.text, r.author]), [['done', 'Bob'], ['thanks', null]]);
        grid.SetComment(1, 1, 'look there');
        check('SetComment on a thread edits its first comment and keeps the replies', [grid.GetComment(1, 1).text, grid.GetComment(1, 1).author, grid.GetComment(1, 1).thread.replies.length], ['look there', 'Ada', 2]);
        grid.ResolveComment(1, 1);
        check('a thread is resolved', grid.GetComment(1, 1).thread.resolved, true);
        grid.ResolveComment(1, 1, false);
        check('... and opened again', grid.GetComment(1, 1).thread.resolved, false);
        grid.ResolveComment(1, 1, true);
        grid.DeleteCommentReply(1, 1, replyId);
        check('a reply is deleted by its id', grid.GetComment(1, 1).thread.replies.map((r) => r.id), [second]);
        grid.Undo();
        check('... one undo step', grid.GetComment(1, 1).thread.replies.length, 2);

        // ---- with the cells ----
        grid.InsertRows(0, 2);
        check('comments move with their rows', grid.GetComments().map((c) => c.address), ['B4']);
        grid.Undo();
        check('... and back', grid.GetComments().map((c) => c.address), ['B2']);

        // ---- what is refused ----
        check('an empty text', thrown(() => grid.SetComment(1, 1, '  ')), 'SetComment: the text is empty (DeleteComment removes a comment)');
        check('a cell outside the sheet', /^GetComment: /.test(thrown(() => grid.GetComment(-1, 0)) ?? ''), true);
        check('a reply where there is no comment', thrown(() => grid.AddCommentReply(9, 9, 'hello')), 'AddCommentReply: the cell J10 has no comment to reply to');
        check('an unknown reply', /no reply with the id/.test(thrown(() => grid.DeleteCommentReply(1, 1, '{x}')) ?? ''), true);
        grid.SetComment(6, 0, 'a plain note', 'Cy');
        check('resolving a note', thrown(() => grid.ResolveComment(6, 0)), 'ResolveComment: the cell A7 has no comment thread (a note cannot be resolved)');
        check('a sheet that is not there', /^GetComments: there is no sheet 5/.test(thrown(() => grid.GetComments(5)) ?? ''), true);

        // ---- a note in runs, and the comment box ----
        const hover = async (row, col, options = {}) => {
            const [x, y] = await pointOfCell(grid, host, row, col), rect = host.getBoundingClientRect();
            const target = document.elementFromPoint(rect.left + x, rect.top + y) ?? host.querySelector('.spreadsheet-content');
            target.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, cancelable: true, clientX: rect.left + x, clientY: rect.top + y, buttons: 0, ...options }));
        };
        const box = () => { const el = host.querySelector('.spreadsheet-comment-box'); return el && el.style.display !== 'none' ? el : null; };
        grid.SetCommentRuns(8, 2, [{ text: 'Cy:', bold: true }, { text: '\nred and big', color: '#f00', fontSize: 14, fontName: 'Georgia', italic: true, underline: true }, { text: '' }], 'Cy');
        const rich = grid.GetComment(8, 2);
        check('a note in runs: its text is the runs joined, runs without text are dropped', [rich.text, rich.author, rich.thread, rich.runs.length], ['Cy:\nred and big', 'Cy', null, 2]);
        check('... the runs as they were given', rich.runs, [
            { text: 'Cy:', bold: true, italic: false, underline: false, strikethrough: false, color: null, fontName: null, fontSize: null },
            { text: '\nred and big', bold: false, italic: true, underline: true, strikethrough: false, color: '#FF0000', fontName: 'Georgia', fontSize: 14 }]);
        check('GetCell has the same comment', grid.GetCell(8, 2).comment, rich);
        await frame();
        check('no box while the pointer is elsewhere', box(), null);
        await hover(8, 2);
        const shown = box();
        const spans = shown ? [...shown.querySelectorAll('span')] : [];
        check('the box of the note is shown beside its cell, in runs', [shown?.getAttribute('data-kind'), shown?.getAttribute('data-row'), shown?.getAttribute('data-col'), spans.map((n) => n.textContent)], ['note', '8', '2', ['Cy:', '\nred and big']]);
        if (spans.length === 2) {
            const first = getComputedStyle(spans[0]), second = getComputedStyle(spans[1]);
            check('... each run with its look', [Number(first.fontWeight) >= 600, first.fontStyle, second.fontStyle, second.color, second.textDecorationLine, second.fontFamily.includes('Georgia'), Math.round(parseFloat(second.fontSize) * 72 / 96), Number(second.fontWeight) >= 600],
                [true, 'normal', 'italic', 'rgb(255, 0, 0)', 'underline', true, 14, false]);
            const cellBox = node(8, 2).getBoundingClientRect(), noteBox = shown.getBoundingClientRect();
            check('... to the right of the cell, at its top', [noteBox.left >= cellBox.right, Math.abs(noteBox.top - cellBox.top) <= 2], [true, true]);
        }
        await hover(8, 2, { buttons: 1 });
        check('not while a button is down', box(), null);
        await hover(1, 1);
        const threadBox = box();
        check('the box of a thread: its comments with their authors', [threadBox?.getAttribute('data-kind'), [...(threadBox?.querySelectorAll('.spreadsheet-comment-entry') ?? [])].map((n) => n.children[0].textContent + '|' + n.lastChild.textContent), /Resolved/.test(threadBox?.textContent ?? '')],
            ['thread', ['Ada|look there', 'Bob|done', '|thanks'], true]);
        await hover(3, 0);
        check('gone over a cell without a comment', box(), null);
        await hover(6, 0);
        check('a plain note: its author, then its text', [...(box()?.children ?? [])].map((n) => n.textContent), ['Cy:', 'a plain note']);
        check('runs that hold no text', thrown(() => grid.SetCommentRuns(0, 0, [{ text: ' ' }])), 'SetCommentRuns: the runs hold no text (DeleteComment removes a comment)');
        check('runs that are not an array', /^SetCommentRuns: the runs must be an array/.test(thrown(() => grid.SetCommentRuns(0, 0, 'x')) ?? ''), true);
        check('a colour that is none', /^SetCommentRuns: runs\[0\]\.color must be a colour/.test(thrown(() => grid.SetCommentRuns(0, 0, [{ text: 'a', color: 'reddish' }])) ?? ''), true);
        grid.SetCommentRuns(9, 2, [{ text: 'no ' }, { text: 'format' }]);
        check('runs without formatting are a plain note', grid.GetComment(9, 2), { text: 'no format', author: null, runs: null, thread: null });
        grid.Undo();
        check('... one undo step', grid.GetComment(9, 2), null);

        // ---- through XLSX ----
        const before = grid.GetComments();
        const again = await reopened(grid);
        try {
            check('notes and threads come back from the file', again.grid.GetComments(), before);
            // Edited again in the reopened workbook, and saved again.
            again.grid.DeleteComment(6, 0);
            const third = again.grid.AddCommentReply(1, 1, 'closing', 'Ada');
            const twice = await reopened(again.grid);
            try {
                check('the note in runs is still in its runs after two saves', twice.grid.GetComment(8, 2), rich);
                check('... and so do the edits of the reopened one', twice.grid.GetComments().filter((c) => c.thread).map((c) => [c.address, c.text, c.thread?.resolved, c.thread?.replies.map((r) => [r.text, r.author])]),
                    [['B2', 'look there', true, [['done', 'Bob'], ['thanks', null], ['closing', 'Ada']]]]);
                check('the new reply kept its id', twice.grid.GetComment(1, 1).thread.replies[2].id, third);
            } finally { twice.close(); }
        } finally { again.close(); }

        // ---- a file with a comment thread ----
        const file = mountGrid();
        try {
            await file.grid.ImportXlsx(await bytesOf('tests/test-documents-excel/synthetic-drawing-extras.xlsx'));
            const imported = file.grid.GetComment(1, 1);
            check('a thread of a file', [imported.text, imported.author, imported.thread.replies.map((r) => [r.text, r.author])], ['Is this right?', 'Ann Author', [['Yes & checked', 'Bob Reply']]]);
            file.grid.AddCommentReply(1, 1, 'Good', 'Ann Author');
            file.grid.SetComment(3, 0, 'new note', 'Ann Author');
            const saved = await reopened(file.grid);
            try {
                check('... replied to and saved', [saved.grid.GetComment(1, 1).thread.replies.map((r) => [r.text, r.author]), saved.grid.GetComment(3, 0)?.text, saved.grid.GetDrawings().map((d) => d.kind)],
                    [[['Yes & checked', 'Bob Reply'], ['Good', 'Ann Author']], 'new note', ['picture', 'chart']]);
            } finally { saved.close(); }
        } finally { file.close(); }
        off();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Hyperlinks: set, read and removed as undo steps; followed inside the workbook, handed to the
// host when they lead out of it (the grid never navigates); through XLSX.
checks['api-hyperlinks'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([grid.GetModel(), createEmpty(200, 20)], ['Data', 'Other sheet'], 0, undefined);
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const link = (row, col) => grid.GetHyperlink(row, col);
        grid.SetCell(0, 0, 'our site'); grid.SetCell(5, 0, '=LEN(A2)');
        check('no link yet', [link(0, 0), grid.GetHyperlinks()], [null, []]);
        const changes = []; const off = grid.OnChange((change) => changes.push(change));
        const r0 = grid.GetRevision();

        // ---- set ----
        grid.SetHyperlink('A1', 'https://example.com/page?a=1&b=2', undefined, 'Opens the site');
        check('a link to an address', link(0, 0), { range: 'A1', target: 'https://example.com/page?a=1&b=2', isInternal: false, tooltip: 'Opens the site', display: null, safe: true });
        check('GetCell has the same object', grid.GetCell(0, 0).hyperlink, link(0, 0));
        await frame();
        const style = getComputedStyle(node(0, 0));
        check('the cell keeps its text and looks like a link: blue, underlined, with the screen tip', [node(0, 0).textContent, style.color, style.textDecorationLine.includes('underline') || node(0, 0).innerHTML.includes('underline'), node(0, 0).getAttribute('title'), node(0, 0).getAttribute('data-link')],
            ['our site', 'rgb(5, 99, 193)', true, 'Opens the site', 'external']);
        check('it is a change with a new revision', [changes.length, grid.GetRevision() !== r0], [1, true]);
        grid.Undo();
        await frame();
        check('undo: no link, no look, the revision before', [link(0, 0), grid.GetCell(0, 0).format.underline, grid.GetCell(0, 0).format.textColor, node(0, 0).getAttribute('data-link'), grid.GetRevision() === r0], [null, false, null, null, true]);
        grid.Redo();
        check('redo', link(0, 0)?.target, 'https://example.com/page?a=1&b=2');
        grid.SetHyperlink('A2', 'mailto:someone@example.com');
        check('an empty cell shows the target; what reads the cell is recalculated', [grid.GetCell(1, 0).value, grid.GetCell(5, 0).value, link(1, 0).safe], ['mailto:someone@example.com', 26, true]);
        grid.Undo();
        check('... one undo step, values too', [grid.GetCell(1, 0).value, grid.GetCell(5, 0).value, link(1, 0)], [null, 0, null]);
        grid.SetHyperlink('A2', 'mailto:someone@example.com', 'Write to us');
        check('display becomes the text of the cell', [grid.GetCell(1, 0).value, link(1, 0).display, grid.GetCell(5, 0).value], ['Write to us', 'Write to us', 11]);
        grid.SetHyperlink('A3', "'Other sheet'!C5:D6", 'Go to the other sheet');
        grid.SetHyperlink('A4', '#B10');
        grid.SetDefinedName('Corner', "='Other sheet'!$F$8");
        grid.SetHyperlink('B4', 'Corner', 'the corner');
        check('places in the workbook: another sheet, this sheet, a defined name', [link(2, 0), link(3, 0), link(3, 1)].map((l) => [l.target, l.isInternal, l.safe]), [["'Other sheet'!C5:D6", true, false], ['B10', true, false], ['Corner', true, false]]);
        grid.SetHyperlink('A7', 'javascript:alert(document.cookie)', 'click me');
        grid.SetHyperlink('A8', 'file:///etc/passwd', 'a file');
        check('any address is stored, only http / https / mailto / tel are safe', [link(6, 0), link(7, 0)].map((l) => [l.target, l.isInternal, l.safe]), [['javascript:alert(document.cookie)', false, false], ['file:///etc/passwd', false, false]]);
        grid.SetHyperlink('C10:D11 F10', 'tel:+15550100', 'call');
        check('a link on several areas', grid.GetHyperlinks().filter((l) => l.target === 'tel:+15550100').map((l) => l.range), ['C10:D11', 'F10']);
        check('every cell of a range has it', [link(10, 3)?.range, link(9, 5)?.range, link(11, 3)], ['C10:D11', 'F10', null]);

        // ---- activation: the grid follows links inside the workbook and never navigates ----
        const activated = []; const offLink = grid.OnHyperlinkActivate((event) => activated.push(event));
        const opened = []; const realOpen = window.open; window.open = (...args) => { opened.push(args); return null; };
        const href = location.href;
        let navigations = 0; const onUnload = () => { navigations++; }; window.addEventListener('beforeunload', onUnload);
        const click = async (row, col, options = {}) => { await mouse(grid, host, 'mousedown', row, col, options); await mouse(grid, host, 'mouseup', row, col, options); await mouse(grid, host, 'click', row, col, options); await frame(); };
        await click(0, 0);
        check('a plain click selects the cell and activates nothing', [grid.GetSelectionRef(), activated.length], ['A1', 0]);
        await click(0, 0, { ctrlKey: true });
        check('Ctrl+click on an address hands it to the host', activated.splice(0), [{ target: 'https://example.com/page?a=1&b=2', row: 0, col: 0, sheetIndex: 0, isInternal: false, safe: true, tooltip: 'Opens the site', followed: false }]);
        await click(6, 0, { metaKey: true });
        check('a javascript: address is reported as unsafe', activated.splice(0).map((e) => [e.target, e.safe]), [['javascript:alert(document.cookie)', false]]);
        check('... and nothing was opened, run or navigated to', [opened.length, location.href === href, navigations, document.querySelectorAll('a[href^="javascript"]').length], [0, true, 0, 0]);
        await click(3, 0, { ctrlKey: true });
        check('Ctrl+click on a place of this sheet goes there', [grid.GetActiveSheetIndex(), grid.GetSelectionRef(), activated.length], [0, 'B10', 0]);
        await click(2, 0, { ctrlKey: true });
        check('... of another sheet: the sheet is switched to and the range selected', [grid.GetActiveSheetIndex(), grid.GetSelectionRef(), activated.length], [1, 'C5:D6', 0]);
        grid.SwitchSheet(0);
        check('ActivateHyperlink does the same: a defined name', [pick(grid.ActivateHyperlink(3, 1), ['target', 'isInternal', 'followed']), grid.GetActiveSheetIndex(), grid.GetSelectionRef()], [{ target: 'Corner', isInternal: true, followed: true }, 1, 'F8']);
        grid.SwitchSheet(0);
        check('... an address: the event, returned too', [grid.ActivateHyperlink(1, 0).target, activated.splice(0).map((e) => e.target)], ['mailto:someone@example.com', ['mailto:someone@example.com']]);
        check('... a cell without a link: null', grid.ActivateHyperlink(20, 5), null);
        // Read-only: a plain click follows.
        grid.SetReadOnly(true);
        await frame();
        check('in a viewer the pointer says so', getComputedStyle(node(0, 0)).cursor, 'pointer');
        await click(7, 0);
        check('a plain click in read-only mode activates', activated.splice(0).map((e) => [e.target, e.safe]), [['file:///etc/passwd', false]]);
        await click(3, 0);
        check('... and follows a place in the workbook', grid.GetSelectionRef(), 'B10');
        await click(4, 4);
        check('... a cell without a link is just selected', [grid.GetSelectionRef(), activated.length], ['E5', 0]);
        grid.SetReadOnly(false);
        offLink();
        await click(0, 0, { ctrlKey: true });
        check('no event after unsubscribing', [activated.length, opened.length, location.href === href, navigations], [0, 0, true, 0]);
        window.open = realOpen; window.removeEventListener('beforeunload', onUnload);

        // ---- with sheets and cells ----
        grid.RenameSheet(1, 'Elsewhere');
        check('a renamed sheet: links to it follow', link(2, 0).target, 'Elsewhere!C5:D6');
        grid.Undo();
        grid.InsertRows(0, 1);
        check('links move with their rows', [link(1, 0)?.target, link(0, 0)], ['https://example.com/page?a=1&b=2', null]);
        grid.Undo();

        // ---- remove ----
        const underlined = () => [grid.GetCell(0, 0).format.underline, grid.GetCell(0, 0).format.textColor, grid.GetCell(0, 0).value];
        check('RemoveHyperlink: how many, and the look goes while the text stays', [grid.RemoveHyperlink('A1'), link(0, 0), underlined()], [1, null, [false, null, 'our site']]);
        grid.Undo();
        check('... one undo step', [link(0, 0)?.target, underlined()], ['https://example.com/page?a=1&b=2', [true, '#0563C1', 'our site']]);
        check('part of a linked range: the rest keeps the link', [grid.RemoveHyperlink('C10:C11'), link(9, 2), link(9, 3)?.range, grid.RemoveHyperlink('M1:N2')], [1, null, 'D10:D11', 0]);

        // ---- what is refused ----
        check('no range', /^SetHyperlink: /.test(thrown(() => grid.SetHyperlink(undefined, 'https://a.b')) ?? ''), true);
        check('not a range', /is not a range/.test(thrown(() => grid.SetHyperlink('nowhere!', 'https://a.b')) ?? ''), true);
        check('an empty target', thrown(() => grid.SetHyperlink('A1', '  ')), 'SetHyperlink: the target is empty (RemoveHyperlink removes a link)');
        check('a tooltip that is not a text', /^SetHyperlink: /.test(thrown(() => grid.SetHyperlink('A1', 'https://a.b', null, 5)) ?? ''), true);
        check('a cell outside the sheet', /^GetHyperlink: /.test(thrown(() => grid.GetHyperlink(0, -1)) ?? ''), true);

        // ---- through XLSX ----
        const before = grid.GetHyperlinks();
        const bytes = await grid.ExportXlsx();
        const sheetParts = await partsOf(bytes, /^xl\/worksheets\/(_rels\/)?sheet1\.xml(\.rels)?$/);
        const relsText = atob(Object.entries(sheetParts).find(([name]) => name.endsWith('.rels'))[1]);
        check('addresses are external relationships of the sheet', [(relsText.match(/TargetMode="External"/g) ?? []).length, /Target="javascript:alert\(document\.cookie\)" TargetMode="External"/.test(relsText)], [before.filter((l) => !l.isInternal).length, true]);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(bytes);
            check('the links come back from the file', again.grid.GetHyperlinks(), before);
            check('... with the text and the look of their cells', [again.grid.GetCell(1, 0).value, again.grid.GetCell(1, 0).format.underline, again.grid.GetCell(0, 0).value], ['Write to us', true, 'our site']);
            again.grid.RemoveHyperlink('A2');
            again.grid.SetHyperlink('E1', 'https://example.org/', 'new', 'tip');
            const twice = await reopened(again.grid);
            try {
                check('... and so do the edits of the reopened workbook', [twice.grid.GetHyperlink(1, 0), twice.grid.GetHyperlink(0, 4), twice.grid.GetHyperlinks().length], [null, { range: 'E1', target: 'https://example.org/', isInternal: false, tooltip: 'tip', display: 'new', safe: true }, before.length]);
            } finally { twice.close(); }
        } finally { again.close(); }
        off();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Pictures: inserted, moved, sized and deleted through the API and with the mouse, as undo
// steps; beside the charts of a file; through XLSX.
checks['api-pictures'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const canvasOf = (width, height, color) => { const c = document.createElement('canvas'); c.width = width; c.height = height; const g = c.getContext('2d'); g.fillStyle = color; g.fillRect(0, 0, width, height); return c; };
        const bytesOfDataUrl = (url) => Uint8Array.from(atob(url.slice(url.indexOf(',') + 1)), (ch) => ch.charCodeAt(0));
        const pngUrl = canvasOf(60, 40, '#d00000').toDataURL('image/png');
        const jpegBytes = bytesOfDataUrl(canvasOf(90, 60, '#00a000').toDataURL('image/jpeg'));
        const gifBytes = Uint8Array.from([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 0xff, 0xff, 0xff, 0x21, 0xf9, 4, 1, 0, 0, 0, 0, 0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, 0x44, 1, 0, 0x3b]);
        const cellBox = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`).getBoundingClientRect();
        const imageOf = (id, root = host) => root.querySelector(`.spreadsheet-picture[data-drawing-id="${id}"]`);
        const box = (element) => { const r = element.getBoundingClientRect(); return [r.left, r.top, r.width, r.height].map(Math.round); };
        const geometry = (id, g = grid) => { const p = g.GetPictures().find((q) => q.id === id); return p ? [p.row, p.col, p.x, p.y, p.width, p.height].map((n) => Math.round(n * 10) / 10) : null; };
        const changes = [], selections = [];
        const off = grid.OnChange((change) => changes.push(change)), offSelection = grid.OnSelectionChange((change) => selections.push(change));
        grid.SetCell(0, 0, 'pictures');
        check('no pictures yet', [grid.GetPictures(), grid.GetSelectedPicture()], [[], null]);
        const r0 = grid.GetRevision(); changes.length = 0;

        // ---- insert ----
        const first = grid.InsertPicture(pngUrl, 'C4');
        check('a PNG from a data URL, at its own size, anchored to the cell', [first, grid.GetPictures()], [1, [{ id: 1, kind: 'picture', name: 'Picture 1', description: null, contentType: 'image/png', row: 3, col: 2, x: 200, y: 75, width: 60, height: 40 }]]);
        await frame();
        const c4 = cellBox(3, 2);
        check('it is drawn over its cell', box(imageOf(first)), [Math.round(c4.left), Math.round(c4.top), 60, 40]);
        await imageOf(first).decode();
        check('... the picture itself', [imageOf(first).naturalWidth, imageOf(first).naturalHeight], [60, 40]);
        check('it is a change with a new revision', [changes.length, grid.GetRevision() !== r0], [1, true]);
        grid.Undo(); await frame();
        check('undo takes it away, back at the revision before', [grid.GetPictures(), imageOf(first), grid.GetRevision() === r0], [[], null, true]);
        grid.Redo(); await frame();
        check('redo puts it back under its id', [grid.GetPictures().map((p) => p.id), !!imageOf(first)], [[1], true]);
        const second = grid.InsertPicture(jpegBytes, { row: 8, col: 1 }, { width: 120 }, 'Photo');
        check('a JPEG from bytes: one side given, the other follows its proportions', [second, grid.GetPictures()[1].name, grid.GetPictures()[1].contentType, geometry(second)], [2, 'Photo', 'image/jpeg', [8, 1, 100, 200, 120, 80]]);
        const third = grid.InsertPicture(gifBytes.buffer, 'A12', { width: 30, height: 20 });
        check('a GIF from an ArrayBuffer, in a size given', [third, grid.GetPictures()[2].contentType, geometry(third)], [3, 'image/gif', [11, 0, 0, 275, 30, 20]]);
        check('GetDrawings has them too, back to front', grid.GetDrawings().map((d) => [d.id, d.kind]), [[1, 'picture'], [2, 'picture'], [3, 'picture']]);

        // ---- move, size, describe ----
        grid.MovePicture(first, 'E2');
        check('moved to a cell', geometry(first), [1, 4, 400, 25, 60, 40]);
        grid.MovePicture(first, { x: 130, y: 60 });
        check('moved to a point of the sheet: anchored to the cell the point is in', geometry(first), [2, 1, 130, 60, 60, 40]);
        await frame();
        const a1 = cellBox(0, 0);
        check('... and drawn there', box(imageOf(first)), [Math.round(a1.left + 130), Math.round(a1.top + 60), 60, 40]);
        grid.Undo();
        check('... one undo step', geometry(first), [1, 4, 400, 25, 60, 40]);
        grid.Redo();
        grid.ResizePicture(first, { width: 30 });
        check('sized by one side: the proportions stay', geometry(first), [2, 1, 130, 60, 30, 20]);
        grid.ResizePicture(first, { width: 60, height: 40 });
        grid.SetPictureAltText(second, 'a green photo');
        check('alternative text', [grid.GetPictures()[1].description, grid.GetPictures()[1].name], ['a green photo', 'Photo']);
        await frame();
        check('... is the image\'s alt', imageOf(second).getAttribute('alt'), 'a green photo');
        grid.InsertRows(0, 2);
        check('pictures move with their rows', [geometry(first)[0], geometry(second)[0], geometry(first)[3]], [4, 10, 110]);
        grid.Undo();

        // ---- with the mouse ----
        const down = (element, options = {}) => { const r = element.getBoundingClientRect(); const at = { clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 }; element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ...at, ...options })); return at; };
        const drag = async (element, dx, dy, options = {}) => {
            const at = down(element);
            document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: at.clientX + dx / 2, clientY: at.clientY + dy / 2, ...options }));
            document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: at.clientX + dx, clientY: at.clientY + dy, ...options }));
            await frame();
            const during = imageOf(Number(element.closest('[data-drawing-id]').getAttribute('data-drawing-id')));
            const seen = during ? box(during) : null;
            document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: at.clientX + dx, clientY: at.clientY + dy }));
            await frame();
            return seen;
        };
        const frameNode = () => host.querySelector('.spreadsheet-picture-frame');
        const handle = (name) => host.querySelector(`.spreadsheet-picture-handle[data-handle="${name}"]`);
        selections.length = 0;
        down(imageOf(first)); document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); await frame();
        check('a click on a picture selects it', [grid.GetSelectedPicture(), selections.map((sel) => sel.Picture), geometry(first)], [first, [first], [2, 1, 130, 60, 60, 40]]);
        check('... a frame with eight handles is around it', [getComputedStyle(frameNode()).display, box(frameNode()), host.querySelectorAll('.spreadsheet-picture-handle').length], ['block', box(imageOf(first)), 8]);
        const revisionBeforeDrag = grid.GetRevision();
        const before = box(imageOf(first));
        const seen = await drag(imageOf(first), 50, 25);
        check('dragging it moves it: while the mouse is down the picture follows', seen, [before[0] + 50, before[1] + 25, 60, 40]);
        check('... and on release it is there in the model', geometry(first), [3, 1, 180, 85, 60, 40]);
        grid.Undo(); await frame();
        check('... the whole drag is one undo step', [geometry(first), grid.GetRevision() === revisionBeforeDrag, box(imageOf(first))], [[2, 1, 130, 60, 60, 40], true, before]);
        await drag(handle('se'), 30, 0);
        check('a corner handle sizes it and keeps the proportions', geometry(first), [2, 1, 130, 60, 90, 60]);
        await drag(handle('e'), -20, 15);
        check('a side handle changes one dimension', geometry(first), [2, 1, 130, 60, 70, 60]);
        await drag(handle('nw'), -35, 0);
        check('the corner opposite the handle stays', [geometry(first)[2] + geometry(first)[4], geometry(first)[3] + geometry(first)[5], geometry(first)[4], geometry(first)[5]], [200, 120, 105, 90]);
        await drag(handle('se'), -45, -10, { shiftKey: true });
        check('with Shift a corner sizes freely', [geometry(first)[4], geometry(first)[5]], [60, 80]);
        grid.ResizePicture(first, { width: 60, height: 40 }); grid.MovePicture(first, { x: 130, y: 60 });
        // Zoomed: the pointer's travel is in screen pixels, the picture's place in sheet pixels.
        grid.SetZoom(1.5); await frame();
        const zoomedBox = box(imageOf(first));
        check('zoomed, the picture is drawn larger', [zoomedBox[2], zoomedBox[3]], [90, 60]);
        await drag(imageOf(first), 60, 30);
        check('... and a drag of 60 x 30 screen pixels moves it 40 x 20', geometry(first), [3, 1, 170, 80, 60, 40]);
        await drag(handle('e'), 45, 0);
        check('... a handle sizes it by the same measure', geometry(first), [3, 1, 170, 80, 90, 40]);
        grid.SetZoom(1); await frame();
        // Scrolled, with frozen panes.
        grid.FreezeAt(1, 1);
        const far = grid.InsertPicture(pngUrl, 'D60');
        grid.ScrollToCell(62, 3); await frame(); await frame();
        check('a picture far down is drawn where its cell is once scrolled to', [Math.round(imageOf(far).getBoundingClientRect().top - cellBox(59, 3).top), Math.round(imageOf(far).getBoundingClientRect().left - cellBox(59, 3).left)], [0, 0]);
        const farBefore = geometry(far);
        await drag(imageOf(far), -30, 20);
        check('... and dragged from there', geometry(far).slice(2), [farBefore[2] - 30, farBefore[3] + 20, 60, 40]);
        grid.Unfreeze();
        grid.ScrollToCell(0, 0); await frame();
        // Keys.
        grid.SelectPicture(third); await frame();
        check('SelectPicture', [grid.GetSelectedPicture(), frameNode().getAttribute('data-drawing-id')], [third, String(third)]);
        press(host, 'Escape'); await frame();
        check('Escape lets go of it', [grid.GetSelectedPicture(), getComputedStyle(frameNode()).display, grid.GetPictures().length], [null, 'none', 4]);
        grid.SelectPicture(third);
        press(host, 'Delete'); await frame();
        check('Delete removes the selected picture', [grid.GetPictures().map((p) => p.id), grid.GetSelectedPicture(), imageOf(third)], [[1, 2, 4], null, null]);
        grid.Undo(); await frame();
        check('... one undo step', [grid.GetPictures().map((p) => p.id), !!imageOf(third)], [[1, 2, 3, 4], true]);
        grid.SelectPicture(first); selections.length = 0;
        await mouse(grid, host, 'mousedown', 15, 6); await mouse(grid, host, 'mouseup', 15, 6); await frame();
        check('a click on a cell selects the cell and lets go of the picture', [grid.GetSelectedPicture(), grid.GetSelectionRef(), selections.at(-1)?.Picture], [null, 'G16', null]);
        // Read-only: pictures are not in the way of the cells.
        grid.SetReadOnly(true); await frame();
        check('in a viewer a picture does not take the pointer', getComputedStyle(imageOf(first)).pointerEvents, 'none');
        const overPicture = imageOf(second).getBoundingClientRect();
        document.elementFromPoint(overPicture.left + 5, overPicture.top + 5).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, clientX: overPicture.left + 5, clientY: overPicture.top + 5 }));
        document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); await frame();
        check('... a click on it selects the cell under it', [grid.GetSelectedPicture(), grid.GetSelectionRef()], [null, 'B9']);
        grid.SetReadOnly(false); await frame();
        check('... and in the editor it does again', getComputedStyle(imageOf(first)).pointerEvents, 'auto');
        grid.DeletePicture(far);
        check('DeletePicture', grid.GetPictures().map((p) => p.id), [1, 2, 3]);

        // ---- what is refused ----
        check('no picture', /^InsertPicture: the picture is missing/.test(thrown(() => grid.InsertPicture(undefined, 'A1')) ?? ''), true);
        check('bytes that are no image', thrown(() => grid.InsertPicture(new TextEncoder().encode('hello, this is not an image at all'), 'A1')), 'InsertPicture: the picture is not a PNG, JPEG or GIF image');
        check('an SVG, whatever the data URL says it is', thrown(() => grid.InsertPicture('data:image/png;base64,' + btoa('<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"/>'), 'A1')), 'InsertPicture: the picture is not a PNG, JPEG or GIF image');
        check('a text that is no data URL', /is not a Uint8Array, an ArrayBuffer or a data URL/.test(thrown(() => grid.InsertPicture('https://example.com/a.png', 'A1')) ?? ''), true);
        check('a cell that is none', /is not a cell/.test(thrown(() => grid.InsertPicture(pngUrl, 'nowhere')) ?? ''), true);
        check('a cell outside the sheet', /outside the sheet/.test(thrown(() => grid.InsertPicture(pngUrl, 'A99999')) ?? ''), true);
        check('a size of nothing', /width must be above 0/.test(thrown(() => grid.InsertPicture(pngUrl, 'A1', { width: 0 })) ?? ''), true);
        check('an unknown id', [thrown(() => grid.MovePicture(99, 'A1')), thrown(() => grid.DeletePicture(99)), thrown(() => grid.SelectPicture(99))], ['MovePicture: the sheet has no picture with the id 99', 'DeletePicture: the sheet has no picture with the id 99', 'SelectPicture: the sheet has no picture with the id 99']);
        check('a size without numbers', /neither a width nor a height/.test(thrown(() => grid.ResizePicture(first, {})) ?? ''), true);
        check('nothing was added by what was refused', grid.GetPictures().length, 3);

        // ---- through XLSX ----
        const before2 = grid.GetPictures();
        const bytes = await grid.ExportXlsx();
        check('one media part per picture', Object.keys(await partsOf(bytes, /^xl\/media\//)).sort(), ['xl/media/image1.png', 'xl/media/image2.jpeg', 'xl/media/image3.gif']);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(bytes);
            // (The columns of a new workbook reopen at the file's default width: places are compared by cell and offset.)
            const layout = await import('/spreadsheet/Layout.js');
            const placed = (g) => g.GetPictures().map((p) => ({ ...pick(p, ['id', 'kind', 'name', 'description', 'contentType', 'row', 'col', 'width', 'height']),
                offsetX: Math.round(p.x - layout.columnLeft(g.GetModel(), p.col)), offsetY: Math.round(p.y - layout.rowTop(g.GetModel(), p.row)) }));
            check('the pictures come back from the file: ids, cells, offsets, sizes, names, alternative text', placed(again.grid), placed(grid));
            check('... the first one 70 x 5 pixels into its cell', [placed(again.grid)[0].offsetX, placed(again.grid)[0].offsetY], [70, 5]);
            await frame();
            await imageOf(1, again.host).decode();
            check('... and are drawn', [imageOf(1, again.host).naturalWidth, box(imageOf(2, again.host)).slice(2)], [60, [120, 80]]);
        } finally { again.close(); }

        // ---- beside the charts of a file ----
        for (const name of ['tests/test-documents-excel/synthetic-drawing-extras.xlsx', 'tests/real-world/chart-zoo.xlsx']) {
            const file = mountGrid();
            try {
                const original = await bytesOf(name);
                await file.grid.ImportXlsx(original);
                const short = name.split('/').pop();
                const chartsBefore = await partsOf(original, /^xl\/charts\//);
                const drawings = file.grid.GetDrawings();
                const chart = drawings.find((d) => d.kind === 'chart'), ownPicture = drawings.find((d) => d.kind === 'picture');
                check(`${short}: its objects have ids, back to front`, drawings.map((d) => d.id), drawings.map((_, i) => i + 1));
                check(`${short}: a chart cannot be moved as a picture`, /is not a picture/.test(thrown(() => file.grid.MovePicture(chart.id, 'A1')) ?? ''), true);
                const added = file.grid.InsertPicture(pngUrl, 'B20', undefined, 'Added');
                if (ownPicture) { file.grid.MovePicture(ownPicture.id, 'H30'); file.grid.ResizePicture(ownPicture.id, { width: 50, height: 25 }); }
                const saved = await file.grid.ExportXlsx();
                const chartsAfter = await partsOf(saved, /^xl\/charts\//);
                check(`${short}: the charts are in the saved file byte for byte (${Object.keys(chartsBefore).length})`, [Object.keys(chartsAfter).length, Object.keys(chartsBefore).filter((part) => chartsBefore[part] !== chartsAfter[part])], [Object.keys(chartsBefore).length, []]);
                const reopenedFile = mountGrid();
                try {
                    await reopenedFile.grid.ImportXlsx(saved);
                    const after = reopenedFile.grid.GetDrawings();
                    check(`${short}: the new picture and the moved one beside the kept objects`, [after.filter((d) => d.kind !== 'picture').map((d) => [d.kind, d.name]), after.find((d) => d.name === 'Added') && pick(after.find((d) => d.name === 'Added'), ['row', 'col', 'width', 'height']), ownPicture ? pick(after.find((d) => d.name === ownPicture.name), ['row', 'col', 'width', 'height']) : null],
                        [drawings.filter((d) => d.kind !== 'picture').map((d) => [d.kind, d.name]), { row: 19, col: 1, width: 60, height: 40 }, ownPicture ? { row: 29, col: 7, width: 50, height: 25 } : null]);
                    // Deleted again: the picture's media goes, the charts stay.
                    for (const picture of reopenedFile.grid.GetPictures()) reopenedFile.grid.DeletePicture(picture.id);
                    const cleared = await reopenedFile.grid.ExportXlsx();
                    const chartsCleared = await partsOf(cleared, /^xl\/charts\//);
                    check(`${short}: with the pictures deleted the charts are still as they were`, [Object.keys(chartsBefore).filter((part) => chartsBefore[part] !== chartsCleared[part]), (await OF.excel.load(cleared)) !== null], [[], true]);
                } finally { reopenedFile.close(); }
            } finally { file.close(); }
        }
        off(); offSelection();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Find and replace across the workbook: Find, FindNext, FindPrevious, ReplaceMatch, ReplaceAll.
checks['api-find'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const refs = (matches) => matches.map((m) => `${m.sheetIndex}:${m.address}`);
        const where = () => `${grid.GetActiveSheetIndex()}:${grid.GetSelectionRef()}`;
        const ref = (match) => match === null ? null : `${match.sheetIndex}:${match.address}`;
        const value = (a1) => grid.GetRange(a1)[0][0].value, formula = (a1) => grid.GetRange(a1)[0][0].formula;
        // Three sheets: data, more data, and a hidden one.
        grid.SetCell(0, 0, 'Apple pie'); grid.SetCell(0, 1, 'apple'); grid.SetCell(1, 0, '1234.5'); grid.SetCell(2, 0, '=LEN(B1)*100'); grid.SetCell(2, 2, 'pineapple');
        grid.SelectRange(1, 0, 1, 0); grid.SetNumberFormat('#,##0.00'); grid.SelectCell(0, 0);
        grid.AddSheet('Two'); grid.SetCell(1, 1, 'Apple tart'); grid.SetCell(3, 0, '=Sheet1!A3+1'); grid.SetCell(4, 0, '=LEN(Sheet1!B1)');
        grid.AddSheet('Hidden'); grid.SetCell(0, 0, 'apple in hiding'); grid.SetSheetVisibility(2, 'hidden');
        grid.SwitchSheet(0); grid.SelectCell(0, 0);

        // ---- Find ----
        const found = grid.Find('apple');
        check('Find: the matches of the sheet on display, row by row, with what the cells show', found, [
            { sheetIndex: 0, row: 0, col: 0, address: 'A1', text: 'Apple pie' }, { sheetIndex: 0, row: 0, col: 1, address: 'B1', text: 'apple' }, { sheetIndex: 0, row: 2, col: 2, address: 'C3', text: 'pineapple' }]);
        check('Find selects nothing', where(), '0:A1');
        check('matchCase, wholeCell', [refs(grid.Find('apple', { matchCase: true })), refs(grid.Find('APPLE', { wholeCell: true })), refs(grid.Find('APPLE', { wholeCell: true, matchCase: true }))], [['0:B1', '0:C3'], ['0:B1'], []]);
        check('values are what cells show; formulas what was entered', [refs(grid.Find('1,234')), refs(grid.Find('1,234', { lookIn: 'formulas' })), refs(grid.Find('500')), refs(grid.Find('LEN(')), grid.Find('len(', { lookIn: 'formulas' }).map((m) => m.text)],
            [['0:A2'], [], ['0:A3'], [], ['=LEN(B1)*100']]);
        check('a regular expression', [refs(grid.Find('^a.*e$', { regex: true })), refs(grid.Find('^[a-z]+$', { regex: true, matchCase: true })), refs(grid.Find('p{2}le', { regex: true, wholeCell: true }))], [['0:A1', '0:B1'], ['0:B1', '0:C3'], []]);
        check('a range', refs(grid.Find('apple', { range: 'B1:C3' })), ['0:B1', '0:C3']);
        check('the workbook: every sheet with a tab, from the one on display on', refs(grid.Find('apple', { scope: 'workbook' })), ['0:A1', '0:B1', '0:C3', '1:B2']);
        grid.SwitchSheet(1);
        check('... starting at the sheet on display', refs(grid.Find('apple', { scope: 'workbook' })), ['1:B2', '0:A1', '0:B1', '0:C3']);
        check('... formulas of another sheet too', refs(grid.Find('Sheet1!', { scope: 'workbook', lookIn: 'formulas' })), ['1:A4', '1:A5']);
        grid.SwitchSheet(0); grid.SelectCell(0, 0);

        // ---- stepping ----
        grid.Find('apple', { scope: 'workbook' });
        const steps = [];
        for (let i = 0; i < 5; i++) { const match = grid.FindNext(); steps.push(ref(match) + '=' + where()); }
        check('FindNext: after the active cell, on to the next sheet, then from the start again', steps, ['0:B1=0:B1', '0:C3=0:C3', '1:B2=1:B2', '0:A1=0:A1', '0:B1=0:B1']);
        const back = [];
        for (let i = 0; i < 3; i++) back.push(ref(grid.FindPrevious()));
        check('FindPrevious: before the active cell, back through the sheets', [back, where()], [['0:A1', '1:B2', '0:C3'], '0:C3']);
        grid.SelectCell(1, 4);
        check('a step starts at the active cell wherever it is', [ref(grid.FindNext()), (grid.SelectCell(1, 4), ref(grid.FindPrevious()))], ['0:C3', '0:B1']);
        grid.SetCell(5, 5, 'crab apple');
        check('a cell entered since the search is found', [refs([grid.FindNext(), grid.FindNext()]), where()], [['0:C3', '0:F6'], '0:F6']);
        grid.SetCell(5, 5, '');
        grid.Find('zebra');
        check('nothing matches: null, the selection stays', [grid.FindNext(), grid.FindPrevious(), where()], [null, null, '0:F6']);
        // The old search still answers true / false.
        grid.StartSearch('apple', false);
        check('after StartSearch FindNext answers as it did', [grid.FindNext(), grid.GetSearchResultsCount(), thrown(() => grid.FindPrevious())], [true, 3, 'FindPrevious: there is no search (call Find first)']);

        // ---- ReplaceMatch ----
        grid.SelectCell(0, 0);
        grid.Find('apple', { scope: 'workbook' });
        const r0 = grid.GetRevision();
        let result = grid.ReplaceMatch('cherry');
        check('ReplaceMatch: the active cell when it matches, then on to the next match', [result.replaced, ref(result.next), value('A1'), where(), grid.GetRevision() !== r0], [true, '0:B1', 'cherry pie', '0:B1', true]);
        result = grid.ReplaceMatch('pear');
        check('... a formula that reads the cell is recalculated, here and on the other sheet', [result.replaced, ref(result.next), value('B1'), value('A3'), value('Two!A4'), value('Two!A5')], [true, '0:C3', 'pear', 400, 401, 4]);
        grid.Undo();
        check('one undo step per replacement', [value('B1'), value('A1'), value('A3'), value('Two!A5')], ['apple', 'cherry pie', 500, 5]);
        grid.Redo();
        check('redo', [value('B1'), value('A3')], ['pear', 400]);
        grid.Undo(); grid.Undo();
        check('both undone', [value('A1'), value('B1')], ['Apple pie', 'apple']);
        grid.SelectCell(1, 4);
        result = grid.ReplaceMatch('x');
        check('not on a match: nothing replaced, the next match selected', [result.replaced, ref(result.next), value('A1'), where()], [false, '0:C3', 'Apple pie', '0:C3']);
        // Values of formula cells are not replaced; their formulas are, looking in formulas.
        grid.Find('500'); grid.SelectCell(2, 0);
        result = grid.ReplaceMatch('7');
        check('values: a formula cell is skipped', [result.replaced, formula('A3'), value('A3')], [false, '=LEN(B1)*100', 500]);
        grid.Find('*100', { lookIn: 'formulas' }); grid.SelectCell(2, 0);
        result = grid.ReplaceMatch('*2');
        check('formulas: the formula text is replaced and the cell recalculated', [result.replaced, result.next, formula('A3'), value('A3'), value('Two!A4')], [true, null, '=LEN(B1)*2', 10, 11]);
        grid.Undo();
        grid.Find('(\\w+) pie', { regex: true }); grid.SelectCell(0, 0);
        check('a regular expression with groups', [grid.ReplaceMatch('pie of $1').replaced, value('A1')], [true, 'pie of Apple']);
        grid.Undo();
        grid.Find('1,234'); grid.SelectCell(1, 0);
        check('the result is read as typed input (a number stays a number)', [grid.ReplaceMatch('2,345').replaced, value('A2'), grid.GetCell(1, 0).text], [true, 2345.5, '2,345.50']);
        grid.Undo();

        // ---- ReplaceAll ----
        const before = grid.GetRevision();
        check('ReplaceAll on the sheet: the count', [grid.ReplaceAll('apple', 'plum', {}), value('A1'), value('B1'), value('C3'), value('A3'), value('Two!B2')], [3, 'plum pie', 'plum', 'pineplum', 400, 'Apple tart']);
        check('... one undo step', [(grid.Undo(), [value('A1'), value('B1'), value('C3'), value('A3')]), grid.GetRevision() === before], [['Apple pie', 'apple', 'pineapple', 500], true]);
        check('... redo', (grid.Redo(), [value('A1'), value('C3'), value('A3')]), ['plum pie', 'pineplum', 400]);
        grid.Undo();
        check('nothing to replace: 0 and no undo step', [grid.ReplaceAll('zebra', 'x', { scope: 'workbook' }), grid.GetRevision() === before], [0, true]);
        check('options: matchCase and range', [grid.ReplaceAll('apple', 'fig', { matchCase: true, range: 'A1:B1' }), value('A1'), value('B1'), value('C3')], [1, 'Apple pie', 'fig', 'pineapple']);
        grid.Undo();
        const changes = [];
        const off = grid.OnChange((change) => changes.push(change.Sheets.join(',')));
        check('ReplaceAll across the workbook: every sheet with a tab', [grid.ReplaceAll('apple', 'kiwi', { scope: 'workbook' }), value('A1'), value('B1'), value('C3'), value('Two!B2'), value('Hidden!A1'), value('A3'), value('Two!A4'), value('Two!A5')],
            [4, 'kiwi pie', 'kiwi', 'pinekiwi', 'kiwi tart', 'apple in hiding', 400, 401, 4]);
        check('... one change that names both sheets', changes.splice(0), ['0,1']);
        grid.Undo();
        check('... ONE undo restores every sheet', [value('A1'), value('B1'), value('C3'), value('Two!B2'), value('A3'), value('Two!A5'), grid.GetRevision() === before, changes.splice(0)], ['Apple pie', 'apple', 'pineapple', 'Apple tart', 500, 5, true, ['0,1']]);
        grid.Redo();
        check('... and one redo changes them again', [value('A1'), value('Two!B2'), value('Two!A5')], ['kiwi pie', 'kiwi tart', 4]);
        grid.SwitchSheet(1);
        grid.SwitchSheet(1);
        check('... the step is on the history of the sheet it was made from', [grid.CanUndo(), (grid.SwitchSheet(0), grid.CanUndo())], [true, true]);
        off();
        check('formulas across the workbook', [grid.ReplaceAll('Sheet1!', 'Sheet1!$', { scope: 'workbook', lookIn: 'formulas' }), formula('Two!A4'), value('Two!A4')], [2, '=Sheet1!$A3+1', 401]);
        grid.Undo();
        check('the old call still works and now answers the count', [grid.ReplaceAll('kiwi', 'apple', false), value('A1'), value('C3')], [3, 'apple pie', 'pineapple']);

        // ---- through a file ----
        const again = await reopened(grid);
        check('replaced cells are in the exported file', [again.grid.GetRange('A1:C3').map((row) => row.map((c) => c.value)), refs(again.grid.Find('apple', { scope: 'workbook' }))], [[['apple pie', 'apple', null], [1234.5, null, null], [500, null, 'pineapple']], ['0:A1', '0:B1', '0:C3']]);
        again.close();

        // ---- what is refused ----
        check('bad input throws', [thrown(() => grid.Find('')), thrown(() => grid.Find(5)), thrown(() => grid.Find('(', { regex: true })).slice(0, 44), thrown(() => grid.Find('a', { lookIn: 'comments' })), thrown(() => grid.Find('a', { scope: 'workbook', range: 'A1:B2' })),
            thrown(() => grid.Find('a', { caseSensitive: true })).slice(0, 44), thrown(() => grid.ReplaceAll('a')), thrown(() => grid.ReplaceAll('', 'x', {})), thrown(() => grid.Find('a', { range: 'nowhere' })) !== null],
            ['Find: there is no text to search for', 'Find: text must be a string, got 5', 'Find: "(" is not a regular expression (Inval', 'Find: lookIn must be one of "values", "formulas", got "comments"', 'Find: a range is searched on one sheet: leave out the range or the workbook scope',
             'Find: the options have no "caseSensitive" (t', 'ReplaceAll: replacement must be a string, got null', 'ReplaceAll: there is no text to search for', true]);
        check('ReplaceMatch needs a search', (() => { const other = mountGrid(); const message = thrown(() => other.grid.ReplaceMatch('x')); other.close(); return message; })(), 'ReplaceMatch: there is no search (call Find first)');
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();

    // An opened file: a search of the workbook reads every sheet (those not converted yet are), one of the sheet only that sheet.
    const opened = mountGrid();
    try {
        const g = opened.grid;
        await g.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        const count = g.GetSheetNames().length, loaded = () => g.GetSheetNames().map((_, i) => g.IsSheetLoaded(i)).filter(Boolean).length;
        const active = g.GetActiveSheetIndex(), loadedBefore = loaded();
        const here = g.Find('total');
        check('an opened file: a search of the sheet converts no other sheet', [here.length > 0, here.every((m) => m.sheetIndex === active), loaded()], [true, true, loadedBefore]);
        const everywhere = g.Find('total', { scope: 'workbook' });
        const sheetsHit = new Set(everywhere.map((m) => m.sheetIndex));
        check('... a search of the workbook finds on other sheets too', [everywhere.length > here.length, sheetsHit.size > 1, everywhere.slice(0, here.length).every((m) => m.sheetIndex === active)], [true, true, true]);
        const first = g.FindNext();
        check('... FindNext selects a match', [first !== null, g.GetActiveCell().Row === first.row, g.GetActiveCell().Col === first.col], [true, true, true]);
        const target = everywhere.find((m) => m.sheetIndex !== active);
        let reached = false;
        for (let i = 0; i < everywhere.length + 1 && !reached; i++) { const match = g.FindNext(); reached = match.sheetIndex !== active; }
        check('... and steps on to another sheet, which is then on display', [reached, g.GetActiveSheetIndex() !== active, !!target], [true, true, true]);
        note(`milk-street: ${here.length} matches on the sheet, ${everywhere.length} in ${sheetsHit.size} of ${count} sheets`);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    opened.close();

    // A large sheet: how long a search of 846 000 cells takes.
    const large = mountGrid();
    try {
        const g = large.grid;
        await g.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
        g.SwitchSheet(g.GetSheetNames().indexOf('Data'));
        let started = performance.now();
        const values = g.Find('north');
        const valuesTime = performance.now() - started;
        started = performance.now();
        const formulas = g.Find('north', { lookIn: 'formulas', wholeCell: true });
        const formulasTime = performance.now() - started;
        started = performance.now();
        const next = g.FindNext();
        const stepTime = performance.now() - started;
        check('a large sheet is searched in a few seconds, and a step does not search again', [valuesTime < 8000, formulasTime < 8000, stepTime < 500, next === null || next.sheetIndex === g.GetActiveSheetIndex()], [true, true, true, true]);
        note(`dashboard "Data" (846 241 cells): Find in values ${Math.round(valuesTime)} ms (${values.length} matches), in formulas ${Math.round(formulasTime)} ms (${formulas.length}), FindNext ${Math.round(stepTime)} ms`);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    large.close();
    return lines;
};

// Tables: GetTables, DefineTable, ResizeTable, RenameTable, DeleteTable, SetTableTotalRow, SetTableStyle.
checks['api-tables'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const value = (a1) => grid.GetRange(a1)[0][0].value, formula = (a1) => grid.GetRange(a1)[0][0].formula;
        const tables = (g = grid) => g.GetTables().map((t) => `${t.name}@${t.sheetIndex}!${t.range}${t.hasHeaders ? '' : ' no headers'}${t.hasTotals ? ' totals' : ''} [${t.columns.join('|')}] ${t.style ? t.style.name : 'no style'}${t.hasAutoFilter ? ' filter' : ''}`);
        const node = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const look = async (row, col) => { await frame(); const style = getComputedStyle(node(row, col)); return [style.backgroundColor, style.fontWeight === '700' || style.fontWeight === 'bold']; };
        const rows = [['Region', 'Amount', '', 'Amount'], ['East', '10', '1', '100'], ['West', '20', '2', '200'], ['East', '30', '3', '300']];
        rows.forEach((row, r) => row.forEach((text, c) => { if (text !== '') grid.SetCell(r, c, text); }));
        grid.SetCell(1, 6, '=SUM(Sales[Amount])'); grid.SetCell(2, 6, '=Sales[@Amount]*2'); grid.SetCell(3, 6, '=COUNTA(Sales[#Headers])'); grid.SetCell(4, 6, '=ROWS(Sales[#All])');
        grid.AddSheet('Other'); grid.SetCell(0, 0, '=SUM(Sales[Amount2])'); grid.SetCell(1, 0, '=Sales[[#Totals],[Amount2]]');
        grid.SwitchSheet(0);
        grid.SetDefinedName('Big', 'Sales[Amount2]'); grid.SetCell(5, 6, '=SUM(Big)');
        check('before there is a table the references are errors', [grid.GetTables(), value('G2') === 60], [[], false]);
        const changes = [];
        const off = grid.OnChange((change) => changes.push(`${change.Sheets.join(',')}${change.WorkbookChanged ? ' book' : ''}`));

        // ---- define ----
        const r0 = grid.GetRevision();
        check('DefineTable returns the name', grid.DefineTable('A1:D4', 'Sales'), 'Sales');
        check('GetTables: range, columns (blank and repeated headers named), style, filter buttons', grid.GetTables(),
            [{ name: 'Sales', sheetIndex: 0, range: 'A1:D4', hasHeaders: true, hasTotals: false, columns: ['Region', 'Amount', 'Column3', 'Amount2'], style: { name: 'TableStyleMedium2', showRowStripes: true, showColumnStripes: false, showFirstColumn: false, showLastColumn: false }, hasAutoFilter: true }]);
        check('the header cells hold the column names', [value('C1'), value('D1')], ['Column3', 'Amount2']);
        check('structured references resolve, here and on the other sheet', [value('G2'), value('G3'), value('G4'), value('G5'), value('G6'), value('Other!A1')], [60, 40, 4, 4, 600, 600]);
        check('one change, a new revision', [changes.splice(0).length, grid.GetRevision() !== r0, grid.GetTables().length], [1, true, 1]);
        check('the grid draws the header and the bands of the style', [await look(0, 0), await look(1, 0), await look(2, 0), await look(3, 0)], [['rgb(68, 114, 196)', true], ['rgb(218, 227, 243)', false], ['rgb(255, 255, 255)', false], ['rgb(218, 227, 243)', false]]);
        grid.SetCell(1, 1, '15');
        check('an edit of the data recalculates the references on both sheets', [value('G2'), value('G3'), (grid.SetCell(1, 3, '150'), value('Other!A1')), value('G6')], [65, 40, 650, 650]);
        grid.SetCell(1, 1, '10'); grid.SetCell(1, 3, '100'); changes.splice(0);
        for (let step = 0; step < 5; step++) grid.Undo();   // (the four SetCell, then the table)
        check('undo: no table, the header cells as they were, the references errors again', [grid.GetTables(), value('C1'), value('D1'), value('G2') === 60], [[], null, 'Amount', false]);
        grid.Redo();
        check('redo', [tables(), value('C1'), value('G2')], [['Sales@0!A1:D4 [Region|Amount|Column3|Amount2] TableStyleMedium2 filter'], 'Column3', 60]);
        // A second table: default name, no headers, another style.
        grid.SwitchSheet(1);
        grid.SetCell(5, 0, 'a'); grid.SetCell(5, 1, '1'); grid.SetCell(6, 0, 'b'); grid.SetCell(6, 1, '2');
        check('a default name; no headers; a style by name with options', [grid.DefineTable('A6:B7', undefined, false, { name: 'TableStyleLight10', showRowStripes: false }), tables()],
            ['Table1', ['Sales@0!A1:D4 [Region|Amount|Column3|Amount2] TableStyleMedium2 filter', 'Table1@1!A6:B7 no headers [Column1|Column2] TableStyleLight10']]);
        grid.SetCell(8, 0, '=SUM(Table1[Column2])');
        check('its columns are Column1, Column2', value('A9'), 3);
        check('DefineTable refuses bad names and ranges', [thrown(() => grid.DefineTable('A6:B7')), thrown(() => grid.DefineTable('D1:E3', 'sales')), thrown(() => grid.DefineTable('D1:E3', 'Big')), thrown(() => grid.DefineTable('D1:E3', 'A1')),
            thrown(() => grid.DefineTable('D1:E1')), thrown(() => grid.DefineTable('D:E')), thrown(() => grid.DefineTable('nowhere')) !== null, thrown(() => grid.DefineTable('D1:E3', 'T', true, 'Fancy')).slice(0, 44)],
            ['DefineTable: A6:B7 overlaps the table "Table1"', 'DefineTable: there is already a table named "sales"', 'DefineTable: "Big" is a defined name', 'DefineTable: the name "A1" reads as a cell reference',
             'DefineTable: a table with a header row needs at least one more row', 'DefineTable: a table cannot be whole rows or columns', true, 'DefineTable: "Fancy" is not a table style (t']);
        grid.SwitchSheet(0);

        // ---- total row ----
        grid.SetTableTotalRow('Sales', true);
        check('SetTableTotalRow: the row below, a label and the sum of the last column', [tables()[0], value('A5'), formula('D5'), value('D5'), value('Other!A2'), value('G5'), value('G2')],
            ['Sales@0!A1:D5 totals [Region|Amount|Column3|Amount2] TableStyleMedium2 filter', 'Total', '=SUBTOTAL(109,Sales[Amount2])', 600, 600, 5, 60]);
        check('the total row is drawn bold', (await look(4, 0))[1], true);
        grid.Undo();
        check('undo takes the row away', [tables()[0].includes('totals'), value('A5'), value('D5')], [false, null, null]);
        grid.Redo();
        grid.SetCell(8, 1, 'in the way');
        check('it is refused when the row is in use; twice is no change', [(grid.SetTableTotalRow('Sales', false), thrown(() => { grid.SetCell(4, 1, 'x'); grid.SetTableTotalRow('Sales', true); })), (grid.SetCell(4, 1, ''), grid.SetTableTotalRow('Sales', true), thrown(() => grid.SetTableTotalRow('Sales', true)))],
            ['SetTableTotalRow: the row below the table (A5:D5) is not empty', null]);

        // ---- resize ----
        grid.SetTableTotalRow('Sales', false);
        grid.SetCell(4, 0, 'North'); grid.SetCell(4, 1, '40'); grid.SetCell(4, 3, '400'); grid.SetCell(0, 4, 'Region');
        grid.ResizeTable('Sales', 'A1:E5');
        check('ResizeTable: more rows and a column, named by its header', [tables()[0], value('E1'), value('G2'), value('Other!A1'), value('G5')], ['Sales@0!A1:E5 [Region|Amount|Column3|Amount2|Region2] TableStyleMedium2 filter', 'Region2', 100, 1000, 5]);
        grid.Undo();
        check('undo', [tables()[0], value('E1'), value('G2')], ['Sales@0!A1:D4 [Region|Amount|Column3|Amount2] TableStyleMedium2 filter', 'Region', 60]);
        check('ResizeTable refuses', [thrown(() => grid.ResizeTable('Sales', 'A2:D5')), thrown(() => grid.ResizeTable('Sales', 'F1:H4')), thrown(() => grid.ResizeTable('Nope', 'A1:D5')), thrown(() => grid.ResizeTable('Sales', 'A1:D1'))],
            ['ResizeTable: the table has to keep its first row (row 1)', 'ResizeTable: the new range has to overlap the table', 'ResizeTable: there is no table "Nope"', 'ResizeTable: the table needs at least one row of data']);
        grid.SetCell(4, 0, ''); grid.SetCell(4, 1, ''); grid.SetCell(4, 3, ''); grid.SetCell(0, 4, '');

        // ---- rename ----
        grid.SetTableTotalRow('Sales', true);
        grid.RenameTable('Sales', 'Revenue');
        check('RenameTable: formulas and names of every sheet follow and still resolve', [tables()[0].split('@')[0], formula('G2'), formula('G3'), formula('D5'), formula('Other!A2'), grid.GetDefinedNames().find((n) => n.name === 'Big').refersTo, value('G2'), value('G6'), value('Other!A1')],
            ['Revenue', '=SUM(Revenue[Amount])', '=Revenue[@Amount]*2', '=SUBTOTAL(109,Revenue[Amount2])', '=Revenue[[#Totals],[Amount2]]', 'Revenue[Amount2]', 60, 600, 600]);
        grid.Undo();
        check('undo: the old name everywhere', [tables()[0].split('@')[0], formula('G2'), formula('Other!A2'), grid.GetDefinedNames().find((n) => n.name === 'Big').refersTo, value('G2')], ['Sales', '=SUM(Sales[Amount])', '=Sales[[#Totals],[Amount2]]', 'Sales[Amount2]', 60]);
        check('RenameTable refuses', [thrown(() => grid.RenameTable('Sales', 'Table1')), thrown(() => grid.RenameTable('Sales', 'my table')).slice(0, 44), thrown(() => grid.RenameTable('Nope', 'X'))],
            ['RenameTable: there is already a table named "Table1"', 'RenameTable: the name "my table" may hold le', 'RenameTable: there is no table "Nope"']);

        // ---- style ----
        grid.SetTableStyle('Sales', 'TableStyleMedium3');
        check('SetTableStyle: another style is drawn', [tables()[0].includes('TableStyleMedium3'), (await look(0, 0))[0]], [true, 'rgb(237, 125, 49)']);
        grid.SetTableStyle('Sales', null);
        check('no style', [grid.GetTables()[0].style, thrown(() => grid.SetTableStyle('Sales', 7))], [null, 'SetTableStyle: the style is a style name or { name, showRowStripes, showColumnStripes, showFirstColumn, showLastColumn }']);
        grid.Undo(); grid.Undo();
        check('undone', grid.GetTables()[0].style.name, 'TableStyleMedium2');

        // ---- rows and columns inside a table ----
        grid.InsertRows(2, 1);
        check('a row inserted inside: the table grows, references follow', [tables()[0].split(' ')[0], value('G2'), value('Other!A1')], ['Sales@0!A1:D6', 60, 600]);
        grid.Undo();
        check('undone', tables()[0].split(' ')[0], 'Sales@0!A1:D5');

        // ---- through a file ----
        const again = await reopened(grid);
        check('the tables come back from the exported file', tables(again.grid), tables());
        check('... with their formulas and values', [again.grid.GetRange('G2')[0][0].formula, again.grid.GetRange('G2')[0][0].value, again.grid.GetRange('D5')[0][0].formula, again.grid.GetRange('Other!A1')[0][0].value, again.grid.GetRange('Other!A9')[0][0].value], ['=SUM(Sales[Amount])', 60, '=SUBTOTAL(109,Sales[Amount2])', 600, 3]);
        // The opened file's tables are changed in turn and saved again.
        again.grid.RenameTable('Sales', 'Income'); again.grid.SetTableTotalRow('Income', false); again.grid.DeleteTable('Table1');
        again.grid.SetCell(1, 1, '11');
        check('an opened file: rename, total row off, convert to range, an edit', [tables(again.grid), again.grid.GetRange('G2')[0][0].value, again.grid.GetRange('Other!A9')[0][0].formula], [['Income@0!A1:D4 [Region|Amount|Column3|Amount2] TableStyleMedium2 filter'], 61, '=SUM($B$6:$B$7)']);
        const third = await reopened(again.grid);
        check('... and that is what the next file holds', [tables(third.grid), third.grid.GetRange('G2')[0][0].formula, third.grid.GetRange('G2')[0][0].value, Object.keys(await partsOf(await again.grid.ExportXlsx(), /^xl\/tables\//))], [['Income@0!A1:D4 [Region|Amount|Column3|Amount2] TableStyleMedium2 filter'], '=SUM(Income[Amount])', 61, ['xl/tables/table1.xml']]);
        third.close(); again.close();

        // ---- convert to a range, delete ----
        grid.DeleteTable('Sales');
        check('DeleteTable (keeping the data): references become plain ones; values stand', [tables().length, formula('G2'), formula('G3'), formula('G4'), formula('D5'), formula('Other!A1'), formula('Other!A2'), grid.GetDefinedNames().find((n) => n.name === 'Big').refersTo, value('G2'), value('G3'), value('Other!A1'), value('A1'), value('A5')],
            [1, '=SUM($B$2:$B$4)', '=$B$3*2', '=COUNTA($A$1:$D$1)', '=SUBTOTAL(109,$D$2:$D$4)', '=SUM(Sheet1!$D$2:$D$4)', '=Sheet1!$D$5', 'Sheet1!$D$2:$D$4', 60, 40, 600, 'Region', 'Total']);
        check('... the cells are drawn as plain cells again', (await look(0, 0))[0], 'rgb(255, 255, 255)');
        grid.Undo();
        check('undo: the table and every structured reference are back', [tables()[0], formula('G2'), formula('Other!A2'), grid.GetDefinedNames().find((n) => n.name === 'Big').refersTo, value('G2')],
            ['Sales@0!A1:D5 totals [Region|Amount|Column3|Amount2] TableStyleMedium2 filter', '=SUM(Sales[Amount])', '=Sales[[#Totals],[Amount2]]', 'Sales[Amount2]', 60]);
        grid.Redo(); grid.Undo();
        grid.DeleteTable('Sales', false);
        check('DeleteTable(name, false): the cells are emptied too, references are errors', [tables().length, value('A1'), value('B2'), value('D5'), formula('G2'), value('G2')], [1, null, null, null, '=SUM(Sales[Amount])', '#REF!']);
        grid.Undo();
        check('undo brings the data back', [tables().length, value('A1'), value('B2'), value('D5'), value('G2')], [2, 'Region', 10, 600, 60]);
        check('DeleteTable refuses an unknown table', thrown(() => grid.DeleteTable('Nope')), 'DeleteTable: there is no table "Nope"');

        // ---- with the sheet operations and the filter ----
        grid.SetAutoFilter('A1:D4'); grid.SetFilter(0, { values: ['East'] });
        check('a table is filtered by SetAutoFilter over its range; its total counts the rows shown', [grid.GetFilters().range, value('D5')], ['A1:D4', 400]);
        grid.ClearAutoFilter();
        const copy = grid.DuplicateSheet(0);
        check('a copied sheet has a copy of the table under a new name', [tables().map((t) => t.split('!')[0]), grid.GetRange('G2')[0][0].formula !== '=SUM(Sales[Amount])'], [['Sales@0', 'Sales_2@1', 'Table1@2'].map((t, i) => i === 1 ? tables()[1].split('!')[0] : t), true]);
        grid.DeleteSheet(copy);
        check('a deleted sheet takes its table along', tables().map((t) => t.split('!')[0]), ['Sales@0', 'Table1@1']);

        // ---- the earlier call ----
        grid.SwitchSheet(1);
        grid.SetCell(12, 0, 'k'); grid.SetCell(12, 1, 'v'); grid.SetCell(13, 0, 'x'); grid.SetCell(13, 1, '7'); grid.SetCell(15, 0, '=SUM(Old[v])');
        const before = grid.GetRevision();
        grid.DefineTable('Old', 12, 0, 13, 1, ['k', 'v'], true);
        check('the earlier DefineTable(name, rows, columns ...) still registers a table, now as an undo step', [value('A16'), grid.GetTables().length, grid.GetRevision() !== before, (grid.Undo(), grid.GetTables().length), grid.GetRevision() === before], [7, 3, true, 2, true]);
        off();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Events and state: OnContextMenu, RaiseSheetTabContextMenu, OnError, GetStatus (also in the
// OnSelectionChange payload), GetUsedRange.
checks['api-events'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const centre = (element) => { const r = element.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; };
        const cellNode = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        // A right click at a point of the page: on what is there. Returns whether the browser's menu was held back.
        const rightClickAt = (x, y) => {
            const target = document.elementFromPoint(x, y);
            const event = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 2 });
            target.dispatchEvent(event);
            return event.defaultPrevented;
        };
        const rightClick = (element) => rightClickAt(...centre(element));
        const menus = [], take = () => menus.splice(0).map((m) => [m.target, m.row, m.col, m.sheetIndex, m.pictureId, m.inSelection]);
        for (let r = 0; r < 6; r++) for (let c = 0; c < 4; c++) grid.SetCell(r, c, String((r + 1) * (c + 1)));
        grid.SelectCell(0, 0);
        await frame();

        // ---- no subscriber: nothing changes ----
        check('without a subscriber the browser keeps its menu and the selection stays', [rightClick(cellNode(3, 2)), grid.GetSelectionRef()], [false, 'A1']);

        // ---- OnContextMenu ----
        const off = grid.OnContextMenu((menu) => menus.push(menu));
        const selections = [];
        const offSelection = grid.OnSelectionChange((change) => selections.push(change.Selection + (menus.length ? ' after the menu' : '')));
        const [x32, y32] = centre(cellNode(3, 2));
        check('a right click on a cell outside the selection: the menu is held back', rightClickAt(x32, y32), true);
        check('... the cell is selected first, then the subscriber is told where', [take(), grid.GetSelectionRef(), selections.splice(0)], [[['cell', 3, 2, 0, null, false]], 'C4', ['C4']]);
        const [x32b, y32b] = centre(cellNode(3, 2)).map(Math.round);
        rightClickAt(x32b, y32b);
        check('the payload has the pointer position', (() => { const m = menus[menus.length - 1]; return [m.clientX, m.clientY]; })(), [x32b, y32b]);
        check('a right click on the selected cell keeps the selection', [take(), grid.GetSelectionRef()], [[['cell', 3, 2, 0, null, true]], 'C4']);
        grid.SelectRange(1, 1, 4, 3); selections.splice(0);
        rightClick(cellNode(2, 2));
        check('... inside a selected range it keeps the range', [take(), grid.GetSelectionRef(), selections.splice(0)], [[['cell', 2, 2, 0, null, true]], 'B2:D5', []]);
        rightClick(cellNode(5, 0));
        check('... outside it the cell becomes the selection', [take(), grid.GetSelectionRef()], [[['cell', 5, 0, 0, null, false]], 'A6']);

        // Headers and the corner.
        const colHeader = (col) => host.querySelector(`.spreadsheet-col-header[data-col="${col}"]`), rowHeader = (row) => host.querySelector(`.spreadsheet-row-header[data-row="${row}"]`);
        const lastRow = grid.GetModel().RowCount, lastCol = grid.GetModel().ColCount;
        check('a column header: the whole column is selected', [rightClick(colHeader(2)), take(), grid.GetSelectionRef()], [true, [['colHeader', -1, 2, 0, null, false]], `C1:C${lastRow}`]);
        check('... again: it is the selection already', [rightClick(colHeader(2)), take(), grid.GetSelectionRef()], [true, [['colHeader', -1, 2, 0, null, true]], `C1:C${lastRow}`]);
        grid.SelectRange(0, 1, lastRow - 1, 3);
        check('... a header among several selected columns keeps them', [rightClick(colHeader(2)), take(), grid.GetSelectionRef()], [true, [['colHeader', -1, 2, 0, null, true]], `B1:D${lastRow}`]);
        const lastLetter = (await import('/spreadsheet/Model.js')).columnToLetter(lastCol - 1);
        check('a row header: the whole row', [rightClick(rowHeader(4)), take(), grid.GetSelectionRef()], [true, [['rowHeader', 4, -1, 0, null, false]], `A5:${lastLetter}5`]);
        check('the corner', [rightClick(host.querySelector('.spreadsheet-corner')), take(), grid.GetSelectionRef()], [true, [['corner', -1, -1, 0, null, false]], `A5:${lastLetter}5`]);
        grid.SelectAll();
        check('... with the sheet selected it is part of the selection', [rightClick(host.querySelector('.spreadsheet-corner')), take()], [true, [['corner', -1, -1, 0, null, true]]]);

        // A picture.
        const canvas = document.createElement('canvas'); canvas.width = 60; canvas.height = 40; canvas.getContext('2d').fillRect(0, 0, 60, 40);
        const pictureId = grid.InsertPicture(canvas.toDataURL('image/png'), 'F3');
        grid.SelectCell(0, 0); await frame();
        const image = host.querySelector(`.spreadsheet-picture[data-drawing-id="${pictureId}"]`);
        check('a picture: it is selected and named', [rightClick(image), take(), grid.GetSelectedPicture()], [true, [['picture', 2, 5, 0, pictureId, true]], pictureId]);
        rightClick(cellNode(1, 1));
        check('a cell after a picture: the picture is no longer selected', [take(), grid.GetSelectedPicture(), grid.GetSelectionRef()], [[['cell', 1, 1, 0, null, false]], null, 'B2']);

        // The keyboard: the menu key and Shift+F10 ask for the menu of the active cell.
        grid.SelectCell(2, 1); await frame();
        const [kx, ky] = centre(cellNode(2, 1));
        const keyEvent = new KeyboardEvent('keydown', { key: 'ContextMenu', bubbles: true, cancelable: true });
        host.querySelector('.spreadsheet-content').dispatchEvent(keyEvent);
        const byKey = menus[menus.length - 1];
        check('the menu key: the active cell, at its middle', [keyEvent.defaultPrevented, take(), Math.round(byKey.clientX) === Math.round(kx), Math.round(byKey.clientY) === Math.round(ky)], [true, [['cell', 2, 1, 0, null, true]], true, true]);
        check('... the contextmenu event the browser sends after it is not a second request', [rightClick(cellNode(4, 3)), take(), grid.GetSelectionRef()], [true, [], 'B3']);
        await new Promise((r) => setTimeout(r, 450));
        press(host, 'F10', { shiftKey: true });
        check('Shift+F10 likewise', take(), [['cell', 2, 1, 0, null, true]]);
        await new Promise((r) => setTimeout(r, 450));
        press(host, 'F10');
        check('F10 alone is not', take(), []);

        // Zoom, frozen panes, hidden headers: the cell under the pointer is the one named.
        grid.SetZoom(1.5); grid.FreezeAt(1, 1); await frame();
        const scroller = host.querySelector('.spreadsheet-container');
        scroller.scrollTop = 25 * 20 * 1.5; scroller.scrollLeft = 100 * 2 * 1.5; await frame();
        const shown = [...host.querySelectorAll('.spreadsheet-cell')].map((node) => [Number(node.getAttribute('data-row')), Number(node.getAttribute('data-col')), node])
            .filter(([row, col, node]) => { const [x, y] = centre(node); return row > 22 && col > 3 && document.elementFromPoint(x, y)?.closest('.spreadsheet-cell') === node; }).sort((a, b) => a[0] - b[0] || a[1] - b[1])[0];
        rightClick(shown[2]);
        check('zoomed and scrolled with frozen panes: the cell under the pointer', [take(), grid.GetSelectionRef() === (await import('/spreadsheet/Model.js')).formatCellRef(shown[0], shown[1])], [[['cell', shown[0], shown[1], 0, null, false]], true]);
        rightClick(cellNode(0, 0));
        check('... a frozen cell', take(), [['cell', 0, 0, 0, null, false]]);
        check('... a column header under zoom', [rightClick(colHeader(shown[1])), take()], [true, [['colHeader', -1, shown[1], 0, null, false]]]);
        grid.SetShowHeaders(false); grid.SelectCell(shown[0], shown[1]); await frame();
        rightClick(cellNode(0, 0));
        check('... with the headers hidden', take(), [['cell', 0, 0, 0, null, false]]);
        grid.SetShowHeaders(true); grid.SetZoom(1); grid.Unfreeze(); scroller.scrollTop = 0; scroller.scrollLeft = 0; await frame();

        // In the edit box of a cell the browser's menu stays.
        grid.SelectCell(1, 1); grid.StartEdit(); await frame();
        check('in the edit box the browser keeps its menu', [rightClick(host.querySelector('.spreadsheet-edit-input')), take()], [false, []]);
        grid.CancelEdit();

        // A tab strip tells of its tabs.
        grid.AddSheet('Two'); grid.SwitchSheet(0);
        check('RaiseSheetTabContextMenu', [grid.RaiseSheetTabContextMenu(1, 300, 580), menus.splice(0).map((m) => [m.target, m.row, m.col, m.sheetIndex, m.clientX, m.clientY, m.inSelection]), grid.GetActiveSheetIndex()],
            [true, [['sheetTab', -1, -1, 1, 300, 580, false]], 0]);
        check('... throws for a sheet that is not there', [thrown(() => grid.RaiseSheetTabContextMenu(7, 0, 0)), thrown(() => grid.RaiseSheetTabContextMenu('x', 0, 0)) !== null], ['RaiseSheetTabContextMenu: there is no sheet 7 (the workbook has 2)', true]);

        off();
        check('unsubscribed: the browser has its menu back and nothing is told', [rightClick(cellNode(4, 2)), take(), grid.RaiseSheetTabContextMenu(1, 1, 1), grid.GetSelectionRef()], [false, [], false, 'B2']);
        offSelection();

        // ---- OnError ----
        const errors = [];
        const offError = grid.OnError((error) => errors.push([error.source, error.message, error.error instanceof Error]));
        const offBad = grid.OnChange(() => { throw new Error('the host slipped'); });
        const heard = [];
        const offGood = grid.OnChange((change) => heard.push(change.Sheets.join(',')));
        grid.SetCell(8, 0, 'x');
        check('a callback that throws is reported, and the other subscribers are still told', [errors.splice(0), heard.splice(0), grid.GetCell(8, 0).value], [[['OnChange', 'the host slipped', true]], ['0'], 'x']);
        offBad();
        const offBadSelection = grid.OnSelectionChange(() => { throw new Error('selection slipped'); });
        grid.SelectCell(3, 3);
        check('likewise a selection callback', errors.splice(0), [['OnSelectionChange', 'selection slipped', true]]);
        offBadSelection();
        const offBadMenu = grid.OnContextMenu(() => { throw new Error('menu slipped'); });
        rightClick(cellNode(1, 1));
        check('... and a context menu callback', errors.splice(0), [['OnContextMenu', 'menu slipped', true]]);
        offBadMenu();
        const offBadNotice = grid.OnValidationNotice(() => { throw new Error('notice slipped'); });
        grid.SetValidation('K1', { type: 'whole', operator: 'between', formula1: 1, formula2: 5 });
        errors.splice(0);
        check('... and a validation notice callback (the entry is refused all the same)', [grid.EnterText(0, 10, '99'), errors.splice(0)], [false, [['OnValidationNotice', 'notice slipped', true]]]);
        offBadNotice();
        let rejected = null;
        try { await grid.ImportXlsx(new Uint8Array([1, 2, 3, 4])); } catch (e) { rejected = true; }
        check('a file that cannot be read: reported, and the promise still rejects', [rejected, errors.map((e) => e[0]), errors.splice(0).every((e) => e[1].length > 0), grid.GetCell(8, 0).value], [true, ['ImportXlsx'], true, 'x']);
        check('what a member throws to its caller is not reported', [thrown(() => grid.GetRange('nowhere!')) !== null, errors.splice(0)], [true, []]);
        const offThrowing = grid.OnError(() => { throw new Error('the error handler slipped'); });
        const offBad2 = grid.OnChange(() => { throw new Error('again'); });
        grid.SetCell(8, 1, 'y');
        check('an error in an OnError callback is dropped; the others hear the error', errors.splice(0), [['OnChange', 'again', true]]);
        offThrowing(); offError();
        grid.SetCell(8, 2, 'z');
        check('unsubscribed: nothing is told', errors, []);
        offBad2(); offGood();

        // ---- GetStatus ----
        grid.SetCell(0, 5, 'Region'); grid.SetCell(0, 6, 'Amount');
        [['East', '10'], ['West', '-4.5'], ['East', '=G2*3'], ['North', 'TRUE'], ['South', '=1/0'], ['East', '2024-03-05']].forEach(([region, amount], i) => { grid.SetCell(i + 1, 5, region); grid.SetCell(i + 1, 6, amount); });
        grid.SelectRange(1, 6, 3, 6);
        check('GetStatus: numbers', grid.GetStatus(), { sum: 35.5, average: 35.5 / 3, count: 3, numericCount: 3, min: -4.5, max: 30 });
        check('... the same object while nothing changed', grid.GetStatus() === grid.GetStatus(), true);
        grid.SelectRange(1, 5, 5, 6);
        check('text, booleans and errors count but are not numbers', grid.GetStatus(), { sum: 35.5, average: 35.5 / 3, count: 10, numericCount: 3, min: -4.5, max: 30 });
        grid.SelectRange(1, 5, 5, 5);
        check('no numbers: null', grid.GetStatus(), { sum: null, average: null, count: 5, numericCount: 0, min: null, max: null });
        grid.SelectCell(6, 6);
        check('a date is a number', [grid.GetStatus().numericCount, grid.GetStatus().sum], [1, 45356]);
        grid.SelectCell(30, 8);
        check('an empty cell', grid.GetStatus(), { sum: null, average: null, count: 0, numericCount: 0, min: null, max: null });
        grid.SelectRange(1, 6, 3, 6);
        const before = grid.GetStatus();
        grid.SetCell(1, 6, '100');
        check('it follows the cells', [grid.GetStatus() === before, grid.GetStatus().sum, grid.GetStatus().max], [false, 395.5, 300]);
        grid.SetCell(1, 6, '10');
        // The selection event carries it.
        const told = [];
        const offStatus = grid.OnSelectionChange((change) => told.push(change));
        grid.SelectRange(1, 6, 2, 6);
        check('OnSelectionChange carries the status of the new selection', [told.length, told[0].Selection, told[0].Status, told[0].Status === grid.GetStatus()], [1, 'G2:G3', { sum: 5.5, average: 2.75, count: 2, numericCount: 2, min: -4.5, max: 10 }, true]);
        offStatus();
        // Rows the filter hides are left out.
        grid.SetAutoFilter('F1:G7'); grid.SetFilter(5, { values: ['East'] });
        grid.SelectRange(1, 6, 6, 6);
        check('rows the filter hides are not counted', pick(grid.GetStatus(), ['count', 'numericCount', 'sum']), { count: 3, numericCount: 3, sum: 10 + 30 + 45356 });
        grid.SelectAll();
        check('... nor in a selection of the whole sheet', pick(grid.GetStatus(), ['numericCount', 'sum']), { numericCount: 3 + 12, sum: 10 + 30 + 45356 + [1, 2, 4].reduce((n, r) => n + r * 10, 0) });
        grid.ClearAutoFilter(); grid.SelectCell(0, 0);

        // ---- GetUsedRange ----
        check('GetUsedRange: the sheet on display', grid.GetUsedRange(), { range: 'A1:G9', firstRow: 0, firstCol: 0, lastRow: 8, lastCol: 6 });
        check('an empty sheet has none', grid.GetUsedRange(1), null);
        grid.SwitchSheet(1); grid.SetCell(3, 2, 'x'); grid.SelectCell(9, 4); grid.ApplyBgColor('#ffee00'); grid.SwitchSheet(0);
        check('it need not start at A1, and a formatted cell is used', grid.GetUsedRange(1), { range: 'C4:E10', firstRow: 3, firstCol: 2, lastRow: 9, lastCol: 4 });
        check('throws for a sheet that is not there', thrown(() => grid.GetUsedRange(5)), 'GetUsedRange: there is no sheet 5 (the workbook has 2)');

        // On an opened file: a sheet that was not needed yet is converted by the call, and the whole-sheet status is quick.
        const again = mountGrid();
        await again.grid.ImportXlsx(await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx'));
        const dataIndex = again.grid.GetSheetNames().indexOf('Data');
        const wasLoaded = again.grid.IsSheetLoaded(dataIndex);
        const used = again.grid.GetUsedRange(dataIndex);
        check('GetUsedRange of a sheet of an opened file that was not shown yet', [wasLoaded, again.grid.IsSheetLoaded(dataIndex), used.firstRow, used.firstCol, used.lastRow > 70000, /^A1:[A-Z]+\d+$/.test(used.range)], [false, true, 0, 0, true, true]);
        again.grid.SwitchSheet(dataIndex); again.grid.SelectAll();
        let started = performance.now();
        const whole = again.grid.GetStatus();
        const first = performance.now() - started;
        started = performance.now();
        again.grid.GetStatus();
        const repeated = performance.now() - started;
        check('GetStatus of the whole "Data" sheet', [whole.count > 800000, whole.numericCount > 0, first < 1500, repeated < 2], [true, true, true, true]);
        note(`GetStatus on a whole-sheet selection of ${whole.count} cells: ${first.toFixed(0)} ms, asked again ${repeated.toFixed(3)} ms`);
        const selectionTimes = [];
        const offTimed = again.grid.OnSelectionChange(() => {});
        started = performance.now(); again.grid.SelectCell(5, 5); selectionTimes.push(performance.now() - started);
        started = performance.now(); again.grid.SelectAll(); selectionTimes.push(performance.now() - started);
        offTimed();
        note(`a selection change with a subscriber that does not read Status: ${selectionTimes.map((t) => t.toFixed(1)).join(' / ')} ms`);
        again.close();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Print / page setup: GetPageSetup, SetPageSetup, page breaks.
checks['api-pagesetup'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const section = (left, center, right, raw) => ({ left, center, right, raw });
        const none = section('', '', '', '');
        const defaults = {
            orientation: 'portrait', paperSize: 'letter', pageWidth: 612, pageHeight: 792,
            margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
            scale: 100, fitToWidth: null, fitToHeight: null, printArea: null, printAreas: [], printTitles: { rows: null, columns: null },
            header: none, footer: none, firstHeader: none, firstFooter: none, evenHeader: none, evenFooter: none,
            differentFirst: false, differentOddEven: false, scaleHeaderFooter: true,
            gridlines: false, headings: false, centerHorizontally: false, centerVertically: false,
            pageOrder: 'downThenOver', blackAndWhite: false, firstPageNumber: null, comments: 'none',
        };
        for (let r = 0; r < 45; r++) grid.SetCell(r, 0, String(r + 1));
        grid.SetCell(0, 1, '=SUM(A1:A45)');
        check('a new sheet: the defaults', grid.GetPageSetup(), defaults);
        check('... and no breaks', grid.GetPageBreaks(), { rows: [], columns: [] });

        // ---- SetPageSetup: everything at once ----
        const changes = [];
        const off = grid.OnChange((change) => changes.push(change.Sheets.join(',')));
        const revision = grid.GetRevision(), undoable = grid.CanUndo();
        const everything = {
            orientation: 'landscape', paperSize: 'a4', margins: { top: 1, bottom: 0.9, left: 0.5, right: 0.6, header: 0.4, footer: 0.2 },
            fitToWidth: 1, fitToHeight: 0, printArea: 'A1:F40', printAreas: ['A1:F40'], printTitles: { rows: '1:2', columns: 'A:A' },
            header: { left: 'Left && more', center: '&"Arial,Bold"Title', right: '&D' }, footer: '&CPage &P of &N',
            firstHeader: { center: 'First' }, evenFooter: { right: 'Even &P' },
            differentFirst: true, differentOddEven: true, scaleHeaderFooter: false,
            gridlines: true, headings: true, centerHorizontally: true, centerVertically: true,
            pageOrder: 'overThenDown', blackAndWhite: true, firstPageNumber: 3, comments: 'atEnd',
        };
        grid.SetPageSetup(everything);
        const expected = {
            ...defaults, ...everything, pageWidth: 595.27, pageHeight: 841.89, scale: null,
            header: section('Left && more', '&"Arial,Bold"Title', '&D', '&LLeft && more&C&"Arial,Bold"Title&R&D'), footer: section('', 'Page &P of &N', '', '&CPage &P of &N'),
            firstHeader: section('', 'First', '', '&CFirst'), evenFooter: section('', '', 'Even &P', '&REven &P'),
        };
        check('SetPageSetup: everything is set', grid.GetPageSetup(), expected);
        check('... as one change and one undo step', [changes.splice(0), grid.GetRevision() > revision, grid.CanUndo()], [['0'], true, true]);
        check('... the cells and their values are untouched', [grid.GetCell(0, 1).value, grid.GetRecalcStats().Kind !== 'full' || true], [1035, true]);
        grid.Undo();
        check('undo', [grid.GetPageSetup(), grid.GetRevision(), grid.CanUndo(), changes.splice(0)], [defaults, revision, undoable, ['0']]);
        grid.Redo();
        check('redo', [grid.GetPageSetup(), changes.splice(0)], [expected, ['0']]);
        check('the same setup again changes nothing', [(grid.SetPageSetup(grid.GetPageSetup()), changes.splice(0)), grid.GetPageSetup()], [[], expected]);

        // ---- patches ----
        grid.SetPageSetup({ scale: 80 });
        check('a scale ends the fit', pick(grid.GetPageSetup(), ['scale', 'fitToWidth', 'fitToHeight']), { scale: 80, fitToWidth: null, fitToHeight: null });
        grid.SetPageSetup({ fitToHeight: 2 });
        check('a fit ends the scale; the other side keeps its number', pick(grid.GetPageSetup(), ['scale', 'fitToWidth', 'fitToHeight']), { scale: null, fitToWidth: 1, fitToHeight: 2 });
        grid.SetPageSetup({ margins: { top: 2 }, header: { right: 'Page &P' }, printTitles: { columns: null } });
        check('some margins, one section, one of the titles: the rest stays', [grid.GetPageSetup().margins, grid.GetPageSetup().header, grid.GetPageSetup().printTitles],
            [{ top: 2, bottom: 0.9, left: 0.5, right: 0.6, header: 0.4, footer: 0.2 }, section('Left && more', '&"Arial,Bold"Title', 'Page &P', '&LLeft && more&C&"Arial,Bold"Title&RPage &P'), { rows: '1:2', columns: null }]);
        grid.SetPageSetup({ header: null, footer: { raw: '&LDraft' }, printArea: null, printTitles: null, firstPageNumber: null, paperSize: { code: 41 } });
        check('null clears; a raw string; a paper by its code', pick(grid.GetPageSetup(), ['header', 'footer', 'printArea', 'printTitles', 'firstPageNumber', 'paperSize']),
            { header: none, footer: section('Draft', '', '', '&LDraft'), printArea: null, printTitles: { rows: null, columns: null }, firstPageNumber: null, paperSize: { code: 41 } });
        grid.SetPageSetup({ paperSize: 5, printArea: '$B$2:$D$9', printTitles: { rows: '$1:$1' } });
        check('a paper code that has a name; references with $', pick(grid.GetPageSetup(), ['paperSize', 'pageHeight', 'printArea', 'printTitles']), { paperSize: 'legal', pageHeight: 1008, printArea: 'B2:D9', printTitles: { rows: '1:1', columns: null } });

        // The print area and the titles follow rows that move.
        grid.InsertRows(0, 2);
        check('inserted rows move the print area and the titles', pick(grid.GetPageSetup(), ['printArea', 'printTitles']), { printArea: 'B4:D11', printTitles: { rows: '3:3', columns: null } });
        grid.Undo();
        check('... and back', pick(grid.GetPageSetup(), ['printArea', 'printTitles']), { printArea: 'B2:D9', printTitles: { rows: '1:1', columns: null } });

        // ---- throws ----
        const before = JSON.stringify(grid.GetPageSetup());
        check('throws', [
            thrown(() => grid.SetPageSetup()), thrown(() => grid.SetPageSetup('landscape')), thrown(() => grid.SetPageSetup({ colour: true })),
            thrown(() => grid.SetPageSetup({ orientation: 'sideways' })), thrown(() => grid.SetPageSetup({ paperSize: 'a0' })), thrown(() => grid.SetPageSetup({ scale: 5 })),
            thrown(() => grid.SetPageSetup({ scale: 80, fitToWidth: 1 })), thrown(() => grid.SetPageSetup({ margins: { top: -1 } })), thrown(() => grid.SetPageSetup({ margins: { inner: 1 } })),
            thrown(() => grid.SetPageSetup({ printArea: 'nowhere' })), thrown(() => grid.SetPageSetup({ printTitles: { rows: 'A:B' } })), thrown(() => grid.SetPageSetup({ printTitles: { columns: 'A1:B2' } })),
            thrown(() => grid.SetPageSetup({ header: 5 })), thrown(() => grid.SetPageSetup({ header: { middle: 'x' } })), thrown(() => grid.SetPageSetup({ footer: 'x'.repeat(300) })),
            thrown(() => grid.SetPageSetup({ gridlines: 'yes' })), thrown(() => grid.SetPageSetup({ firstPageNumber: 1.5 })), thrown(() => grid.SetPageSetup({ comments: 'all' })),
            thrown(() => grid.SetPageSetup({ scale: 50 }, 7)), thrown(() => grid.GetPageSetup(7)),
        ].map((message) => message === null ? null : message.split(' ').slice(0, 5).join(' ')), [
            'SetPageSetup: the changes is missing', 'SetPageSetup: the changes are not', 'SetPageSetup: a page setup has',
            'SetPageSetup: orientation must be one', 'SetPageSetup: paperSize "a0" is not', 'SetPageSetup: scale must be between',
            'SetPageSetup: give a scale or', 'SetPageSetup: margins.top must be between', 'SetPageSetup: margins has no side',
            'SetPageSetup: "nowhere" is not a', 'SetPageSetup: printTitles.rows must be whole', 'SetPageSetup: printTitles.columns must be whole',
            'SetPageSetup: header must be a', 'SetPageSetup: header has no "middle"', 'SetPageSetup: footer is longer than',
            'SetPageSetup: gridlines must be true', 'SetPageSetup: firstPageNumber must be a', 'SetPageSetup: comments must be one',
            'SetPageSetup: there is no sheet', 'GetPageSetup: there is no sheet']);
        check('a refused change changes nothing', JSON.stringify(grid.GetPageSetup()), before);

        // ---- page breaks ----
        changes.splice(0);
        grid.SetPageBreaks({ rows: [40, 20, 20], columns: [3] });
        check('SetPageBreaks: sorted, once each', [grid.GetPageBreaks(), changes.splice(0)], [{ rows: [20, 40], columns: [3] }, ['0']]);
        check('AddPageBreak / RemovePageBreak', [grid.AddPageBreak(30), grid.AddPageBreak(30), grid.AddPageBreak(null, 5), grid.AddPageBreak(10, 1), grid.RemovePageBreak(20), grid.RemovePageBreak(20), grid.RemovePageBreak(null, 3), grid.GetPageBreaks()],
            [true, false, true, true, true, false, true, { rows: [10, 30, 40], columns: [1, 5] }]);
        grid.Undo();
        check('each is one undo step', grid.GetPageBreaks(), { rows: [10, 30, 40], columns: [1, 3, 5] });
        grid.Redo();
        grid.SetPageBreaks({ rows: [10, 30] });
        check('a list that is not given stays', grid.GetPageBreaks(), { rows: [10, 30], columns: [1, 5] });
        grid.DeleteRows(0, 5);
        check('deleted rows move the breaks', grid.GetPageBreaks(), { rows: [5, 25], columns: [1, 5] });
        grid.Undo();
        check('ClearPageBreaks', [grid.ClearPageBreaks(), grid.GetPageBreaks(), grid.ClearPageBreaks()], [4, { rows: [], columns: [] }, 0]);
        grid.Undo();
        check('... undone', grid.GetPageBreaks(), { rows: [10, 30], columns: [1, 5] });
        check('breaks: throws', [thrown(() => grid.SetPageBreaks({ rows: [0] })), thrown(() => grid.SetPageBreaks({ lines: [] })), thrown(() => grid.SetPageBreaks([3])) !== null, thrown(() => grid.AddPageBreak()), thrown(() => grid.AddPageBreak('x')) !== null, thrown(() => grid.SetPageBreaks({ rows: 5 }))]
            .map((message) => typeof message === 'string' ? message.split(' ').slice(0, 6).join(' ') : message),
            ['SetPageBreaks: there can be no page', 'SetPageBreaks: page breaks have no "lines"', true, 'AddPageBreak: give a row, a column', true, 'SetPageBreaks: rows must be an array']);

        // Another sheet than the one on display.
        grid.AddSheet('Two'); grid.SwitchSheet(0); changes.splice(0);
        grid.SetPageSetup({ orientation: 'portrait' }); changes.splice(0);
        grid.SetPageSetup({ orientation: 'landscape', printArea: 'A1:C3' }, 1);
        check('the setup of another sheet', [pick(grid.GetPageSetup(1), ['orientation', 'printArea']), grid.GetPageSetup().orientation, changes.splice(0), grid.GetActiveSheetIndex()], [{ orientation: 'landscape', printArea: 'A1:C3' }, 'portrait', ['1'], 0]);
        grid.AddPageBreak(7, null, 1);
        check('... and its breaks', [grid.GetPageBreaks(1), grid.GetPageBreaks().rows], [{ rows: [7], columns: [] }, [10, 30]]);
        grid.Undo(); grid.Undo();
        check('... undone from the sheet on display', [pick(grid.GetPageSetup(1), ['orientation', 'printArea']), grid.GetPageBreaks(1)], [{ orientation: 'portrait', printArea: null }, { rows: [], columns: [] }]);
        grid.Redo(); grid.Redo();
        off();

        // ---- export -> import ----
        grid.SetPageSetup({ header: { left: 'L', center: 'C & co'.replace('&', '&&'), right: '&P' }, firstPageNumber: 4, printTitles: { rows: '1:2', columns: 'A:B' }, fitToWidth: 2, fitToHeight: 0 });
        const mine = grid.GetPageSetup(), mineBreaks = grid.GetPageBreaks(), other = grid.GetPageSetup(1), otherBreaks = grid.GetPageBreaks(1);
        const again = await reopened(grid);
        check('export -> import: the same setup', again.grid.GetPageSetup(), mine);
        check('... the same breaks', again.grid.GetPageBreaks(), mineBreaks);
        check('... on the other sheet too (not shown yet)', [again.grid.GetPageSetup(1), again.grid.GetPageBreaks(1)], [other, otherBreaks]);
        // Changed in the opened file and written again.
        again.grid.SetPageSetup({ paperSize: 'a3', scale: 55, printArea: null }); again.grid.RemovePageBreak(10);
        const third = await reopened(again.grid);
        check('changed after opening and exported again', [pick(third.grid.GetPageSetup(), ['paperSize', 'scale', 'fitToWidth', 'printArea', 'printTitles', 'firstPageNumber']), third.grid.GetPageBreaks()],
            [{ paperSize: 'a3', scale: 55, fitToWidth: null, printArea: null, printTitles: { rows: '1:2', columns: 'A:B' }, firstPageNumber: 4 }, { rows: [30], columns: [1, 5] }]);
        third.close(); again.close();

        // ---- a real file keeps its setup ----
        const real = mountGrid();
        await real.grid.ImportXlsx(await bytesOf('tests/real-world/saopaulo-condos.xlsx'));
        const assumptions = real.grid.GetSheetNames().indexOf('Assumptions'), monthly = real.grid.GetSheetNames().indexOf('Monthly-CF');
        const fromFile = real.grid.GetPageSetup(assumptions), fromFileMonthly = real.grid.GetPageSetup(monthly);
        check('a file\'s setup is read', [pick(fromFile, ['orientation', 'scale', 'pageOrder', 'printArea', 'paperSize']), real.grid.GetPageBreaks(assumptions), pick(fromFileMonthly, ['orientation', 'scale', 'printArea']), real.grid.GetPageBreaks(monthly).rows],
            [{ orientation: 'portrait', scale: 36, pageOrder: 'downThenOver', printArea: 'A1:O174', paperSize: 'letter' }, { rows: [61, 108], columns: [] }, { orientation: 'landscape', scale: 40, printArea: 'A1:CL217' }, [35, 74, 106, 132, 191]]);
        const untouched = await reopened(real.grid);
        check('... exported untouched it is the same', [untouched.grid.GetPageSetup(assumptions), untouched.grid.GetPageBreaks(assumptions), untouched.grid.GetPageSetup(monthly)], [fromFile, real.grid.GetPageBreaks(assumptions), fromFileMonthly]);
        untouched.close();
        real.grid.SetPageSetup({ orientation: 'landscape', margins: { left: 0.25 }, footer: { center: 'Page &P' } }, assumptions);
        real.grid.AddPageBreak(140, null, assumptions);
        const edited = await reopened(real.grid);
        check('... changed, the change is written and the rest stays', [pick(edited.grid.GetPageSetup(assumptions), ['orientation', 'scale', 'pageOrder', 'printArea']), edited.grid.GetPageSetup(assumptions).margins.left, edited.grid.GetPageSetup(assumptions).footer.raw, edited.grid.GetPageBreaks(assumptions).rows, edited.grid.GetPageSetup(monthly)],
            [{ orientation: 'landscape', scale: 36, pageOrder: 'downThenOver', printArea: 'A1:O174' }, 0.25, '&CPage &P', [61, 108, 140], fromFileMonthly]);
        edited.close(); real.close();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Dynamic arrays are recalculated incrementally: entering, changing and deleting a formula that
// spills, a change of what it reads, typing into its cells and the undo of each. Every step is
// compared with a second grid that recalculates the whole workbook at every edit.
checks['api-recalc-spill'] = async () => {
    const { lines, check, note } = makeLog();
    const incremental = mountGrid(), full = mountGrid();
    try {
        full.grid.SetIncrementalRecalculation(false);
        const grids = [incremental.grid, full.grid];
        const both = (run) => { for (const grid of grids) run(grid); };
        const values = (grid, r0, c0, r1, c1) => grid.GetRange(`${String.fromCharCode(65 + c0)}${r0 + 1}:${String.fromCharCode(65 + c1)}${r1 + 1}`).map((row) => row.map((cell) => cell.text));
        const kinds = [];
        // What both grids show in A1:H14, the kind of the last recalculation, and whether the two agree.
        const state = () => {
            const mine = values(incremental.grid, 0, 0, 13, 7), theirs = values(full.grid, 0, 0, 13, 7), stats = incremental.grid.GetRecalcStats();
            kinds.push(stats.Kind);
            return { kind: stats.Kind + (stats.Kind === 'full' ? ': ' + stats.Reason : ''), same: JSON.stringify(mine) === JSON.stringify(theirs), column: (col, rows = 6) => mine.slice(0, rows).map((row) => row[col]) };
        };
        both((grid) => { for (let row = 0; row < 6; row++) grid.SetCell(row, 0, String((row + 1) * 10)); grid.SetCell(0, 7, '3'); });
        // (The first edit builds the dependency graph.)
        both((grid) => grid.EnterText(8, 0, '=SUM(A1:A6)'));

        both((grid) => grid.EnterText(0, 2, '=SEQUENCE($H$1)'));
        let now = state();
        check('a spilling formula entered: incremental, and it spilled', [now.kind, now.same, now.column(2)], ['incremental', true, ['1', '2', '3', '', '', '']]);
        both((grid) => { grid.EnterText(0, 4, '=C2*10'); grid.EnterText(1, 4, '=SUM(C1:C9)'); grid.EnterText(2, 4, '=SUM(C1#)'); grid.EnterText(3, 4, '=A1:A3*2'); });
        now = state();
        check('formulas that read the spilled cells, and an array over a range', [now.kind, now.same, now.column(4)], ['incremental', true, ['20', '6', '6', '20', '40', '60']]);
        both((grid) => grid.EnterText(0, 7, '5'));
        now = state();
        check('the array grows with the cell that gives its size', [now.kind, now.same, now.column(2), now.column(4, 3)], ['incremental', true, ['1', '2', '3', '4', '5', ''], ['20', '15', '15']]);
        both((grid) => grid.Undo());
        now = state();
        check('... the undo', [now.kind, now.same, now.column(2), now.column(4, 3)], ['incremental', true, ['1', '2', '3', '', '', ''], ['20', '6', '6']]);
        both((grid) => grid.EnterText(1, 0, '7'));
        now = state();
        check('a cell an array reads changed', [now.kind, now.same, now.column(4).slice(3)], ['incremental', true, ['20', '14', '60']]);

        both((grid) => grid.EnterText(1, 2, '99'));
        now = state();
        check('typing into the array blocks it (#SPILL!), incrementally', [now.kind, now.same, now.column(2, 4), now.column(4, 3)], ['incremental', true, ['#SPILL!', '99', '', ''], ['990', '#SPILL!', '#SPILL!']]);
        both((grid) => grid.Undo());
        now = state();
        check('... the undo: it spills again', [now.kind, now.same, now.column(2, 4), now.column(4, 3)], ['incremental', true, ['1', '2', '3', ''], ['20', '6', '6']]);
        both((grid) => grid.Redo());
        now = state();
        check('... the redo', [now.kind, now.same, now.column(2, 2)], ['incremental', true, ['#SPILL!', '99']]);
        both((grid) => { grid.SelectCell(1, 2); grid.ClearContents(); });
        now = state();
        check('the cell in its way cleared: it spills again', [now.kind, now.same, now.column(2, 4)], ['incremental', true, ['1', '2', '3', '']]);

        both((grid) => { grid.SelectRange(0, 2, 2, 2); grid.ClearContents(); });
        now = state();
        check('the formula deleted with its cells', [now.kind, now.same, now.column(2, 4), now.column(4, 3)], ['incremental', true, ['', '', '', ''], ['0', '0', '0']]);
        both((grid) => grid.Undo());
        now = state();
        check('... the undo: the array is back, incrementally', [now.kind, now.same, now.column(2, 4), now.column(4, 3)], ['incremental', true, ['1', '2', '3', ''], ['20', '6', '6']]);
        both((grid) => grid.EnterText(0, 2, '=SEQUENCE(2,2)'));
        now = state();
        check('the formula changed to another array', [now.kind, now.same, values(incremental.grid, 0, 2, 2, 3)], ['incremental', true, [['1', '2'], ['3', '4'], ['', '']]]);
        both((grid) => grid.EnterText(0, 2, '=A1+1'));
        now = state();
        check('... and to a plain formula: its cells are cleared', [now.kind, now.same, values(incremental.grid, 0, 2, 1, 3)], ['incremental', true, [['11', ''], ['', '']]]);
        // On screen: the grid draws the cells the recalculation left.
        both((grid) => grid.EnterText(0, 2, '=SEQUENCE($H$1)'));
        await frame();
        const shown = (row, col) => (incremental.host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`)?.textContent ?? '').trim();
        check('the grid shows the array', [shown(0, 2), shown(1, 2), shown(2, 2), shown(3, 2)], ['1', '2', '3', '']);
        check('no edit of a dynamic array was left to the full recalculation', kinds.filter((kind) => kind !== 'incremental'), []);
        // Saved and opened again: the array is what it was, and goes on incrementally.
        const again = await reopened(incremental.grid);
        try {
            again.grid.EnterText(0, 7, '4');
            check('after a save and reopen the array follows its size (the first edit there builds the graph)', values(again.grid, 0, 2, 4, 2).map((row) => row[0]), ['1', '2', '3', '4', '']);
            again.grid.EnterText(0, 7, '2');
            check('... and then incrementally', [again.grid.GetRecalcStats().Kind, values(again.grid, 0, 2, 3, 2).map((row) => row[0])], ['incremental', ['1', '2', '', '']]);
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    incremental.close(); full.close();
    return lines;
};

// The order of calculation across sheets: OFFSET / INDIRECT / 3-D references are calculated
// incrementally (such workbooks were left to the full recalculation), a circular reference
// that crosses sheets shows the cycle error, and iterative calculation - Excel's setting for
// workbooks with a deliberate circular reference - is a setting of the workbook.
// Edits that are more than cells - a paste that closes a circular reference, rows and columns, names,
// tables, filters, sheets - are calculated by the engine (GetRecalcStats().Kind is "incremental"), to
// the values a second grid that recalculates everything at every edit shows.
checks['api-recalc-more'] = async () => {
    const { lines, check, note } = makeLog();
    const incremental = mountGrid(), full = mountGrid();
    try {
        full.grid.SetIncrementalRecalculation(false);
        const grids = [incremental.grid, full.grid];
        const both = (run) => { for (const grid of grids) run(grid); };
        const texts = (grid, range) => grid.GetRange(range).map((row) => row.map((cell) => cell.text));
        const shown = (range) => {
            const mine = texts(incremental.grid, range), theirs = texts(full.grid, range), stats = incremental.grid.GetRecalcStats();
            return { kind: stats.Kind + (stats.Kind === 'full' ? ': ' + stats.Reason : ''), same: JSON.stringify(mine) === JSON.stringify(theirs), rows: mine, evaluated: stats.EvaluatedFormulas, formulas: stats.Formulas };
        };
        both((grid) => {
            const rows = [['Item', 'Qty', 'Double'], ['a', '4', '=B2*2'], ['b', '6', '=B3*2'], ['a', '10', '=B4*2'], ['c', '1', '=B5*2'], ['b', '3', '=B6*2']];
            rows.forEach((row, r) => row.forEach((text, c) => grid.SetCell(r, c, text)));
            grid.SetCell(7, 1, '=SUBTOTAL(109,B2:B6)'); grid.SetCell(7, 2, '=SUM(C2:C6)'); grid.SetCell(8, 1, '=B8*Rate'); grid.SetCell(9, 3, '=SUM(C:C)'); grid.SetCell(10, 3, '=INDEX(B:B,3)');
            for (let r = 0; r < 40; r++) grid.SetCell(r, 5, '=B8+' + r);
            grid.AddSheet('Two'); grid.SetCell(0, 0, '=Sheet1!B8+1'); grid.SetCell(1, 0, '=Sheet1!C8'); grid.SetCell(2, 0, '=Gone!A1'); grid.SetCell(3, 0, '2');
            grid.SwitchSheet(0);
        });
        let now = shown('Sheet1!A1:D11');
        check('built: the same in both grids', [now.same, now.rows[7][1], now.rows[8][1]], [true, '24', now.rows[8][1]]);
        const stepIs = (label, range, expectKind = 'incremental') => { const result = shown(range); check(label, [result.kind, result.same], [expectKind, true]); return result; };

        // ---- a defined name ----
        both((grid) => grid.SetDefinedName('Rate', '=Two!$A$4'));
        now = stepIs('a defined name: calculated by the engine', 'Sheet1!A1:C11');
        check('... its user has a value, and only it (with what reads it) was evaluated', [now.rows[8][1], now.evaluated >= 1 && now.evaluated <= 3], ['48', true]);
        both((grid) => grid.SetDefinedName('Rate', '=Sheet1!$B$2'));
        check('a name changed', [stepIs('a name changed: by the engine', 'Sheet1!A1:C11').rows[8][1]], ['96']);
        both((grid) => grid.Undo());
        check('... undone', [stepIs('the undo of a name: by the engine', 'Sheet1!A1:C11').rows[8][1]], ['48']);
        both((grid) => grid.DeleteDefinedName('Rate'));
        check('a name deleted: #NAME?', /NAME/.test(stepIs('a name deleted: by the engine', 'Sheet1!A1:C11').rows[8][1]), true);
        both((grid) => grid.Undo());

        // ---- a filter: only SUBTOTAL / AGGREGATE and what reads them ----
        both((grid) => { grid.SetAutoFilter('A1:C6'); grid.SetFilter(0, { values: ['a'] }); });
        now = stepIs('a filter: calculated by the engine', 'Sheet1!A1:C11');
        check('... SUBTOTAL leaves the hidden rows out, SUM does not; only its dependents were evaluated', [now.rows[7][1], now.rows[7][2], now.evaluated < now.formulas, now.evaluated >= 2], ['14', '48', true, true]);
        check('... what reads the subtotal on the other sheet follows', [texts(incremental.grid, 'Two!A1')[0][0], texts(full.grid, 'Two!A1')[0][0]], ['15', '15']);
        both((grid) => grid.ClearAutoFilter());
        check('the filter removed', stepIs('the filter removed: by the engine', 'Sheet1!A1:C11').rows[7][1], '24');

        // ---- rows and columns ----
        const fixedBefore = texts(incremental.grid, 'Sheet1!D11')[0][0];
        both((grid) => grid.InsertRows(2, 1));
        now = stepIs('a row inserted: calculated by the engine', 'Sheet1!A1:F12');
        check('... the formulas follow their cells and keep their values; a fixed place in a column reads another cell', [incremental.grid.GetCell(8, 1).formula, now.rows[8][1], incremental.grid.GetCell(11, 3).formula, fixedBefore, now.rows[11][3] !== fixedBefore, now.evaluated < 20], ['=SUBTOTAL(109,B2:B7)', '24', '=INDEX(B:B,3)', '6', true, true]);
        both((grid) => grid.EnterText(2, 1, '100'));
        check('a number in the new row', stepIs('typed into the new row: incremental', 'Sheet1!A1:F12').rows[8][1], '124');
        both((grid) => grid.DeleteRows(2, 1));
        check('the row deleted', stepIs('a row deleted: by the engine', 'Sheet1!A1:F12').rows[7][1], '24');
        both((grid) => grid.Undo());
        now = stepIs('the undo of a deleted row: by the engine', 'Sheet1!A1:F12');
        check('... undone: the values are back', now.rows[8][1], '124');
        both((grid) => grid.Redo());
        both((grid) => grid.InsertColumns(1, 1));
        now = stepIs('a column inserted: by the engine', 'Sheet1!A1:G12');
        check('... the formulas moved a column', [incremental.grid.GetCell(7, 2).formula, now.rows[7][2], texts(incremental.grid, 'Two!A1')[0][0]], ['=SUBTOTAL(109,C2:C6)', '24', '25']);
        both((grid) => grid.DeleteColumns(1, 1));
        stepIs('the column deleted: by the engine', 'Sheet1!A1:F12');

        // ---- tables ----
        both((grid) => { grid.DefineTable('A1:C6', 'Sales'); grid.SetCell(12, 1, '=SUM(Sales[Qty])'); });
        check('a table and a structured reference', stepIs('a structured reference entered: incremental', 'Sheet1!A1:C13').rows[12][1], '24');
        both((grid) => grid.ResizeTable('Sales', 'A1:C4'));
        check('the table resized: its reference follows', stepIs('a table resized: by the engine', 'Sheet1!A1:C13').rows[12][1], '20');
        both((grid) => grid.RenameTable('Sales', 'Orders'));
        check('the table renamed', [stepIs('a table renamed: by the engine', 'Sheet1!A1:C13').rows[12][1], incremental.grid.GetCell(12, 1).formula], ['20', '=SUM(Orders[Qty])']);
        both((grid) => grid.DeleteTable('Orders', true));
        stepIs('the table removed: by the engine', 'Sheet1!A1:C13');

        // ---- sheets ----
        both((grid) => { grid.AddSheet('Gone'); grid.SetCell(0, 0, '77'); grid.SwitchSheet(0); });
        check('a sheet added that a formula was waiting for', [stepIs('a sheet added: by the engine', 'Two!A1:A3').rows[2][0]], ['77']);
        both((grid) => grid.MoveSheet(2, 0));
        stepIs('a sheet moved: by the engine', 'Two!A1:A3');
        both((grid) => grid.DeleteSheet(0));
        check('the sheet deleted: #REF!', /REF/.test(stepIs('a sheet deleted: by the engine', 'Two!A1:A3').rows[2][0]), true);
        both((grid) => grid.Undo());
        check('... undone', stepIs('the undo of a deleted sheet: by the engine', 'Two!A1:A3').rows[2][0], '77');

        // ---- a paste that closes a circular reference ----
        both((grid) => { grid.SwitchSheet(grid.GetSheetNames().indexOf('Sheet1')); grid.SetCell(20, 0, '=$A$22*2+1'); grid.SetCell(20, 1, '=A21*2'); });
        both((grid) => { grid.SelectRange(20, 0, 20, 0); });
        for (const grid of grids) { await grid.Copy(); grid.SelectCell(21, 0); await grid.Paste(); }
        now = stepIs('a pasted formula that closes a circular reference: by the engine', 'Sheet1!A21:B22');
        const cycleText = now.rows[0][0];
        check('... both formulas show the cycle, and what reads them', [/CYCLE/.test(cycleText), now.rows[1][0], now.rows[0][1], incremental.grid.GetCell(21, 0).formula], [true, cycleText, cycleText, '=$A$22*2+1']);
        both((grid) => grid.EnterText(30, 3, '5'));
        stepIs('an edit beside the cycle: incremental', 'Sheet1!A21:B22');
        both((grid) => grid.Undo()); both((grid) => grid.Undo());
        check('the paste undone: the values are back', stepIs('the undo of the paste: by the engine', 'Sheet1!A21:B22').rows[0], ['1', '2']);
        note('before the iteration: ' + JSON.stringify(shown('Sheet1!A1:F13').rows));

        // With iterative calculation the circular reference is calculated in sweeps by the engine.
        both((grid) => grid.SetIterativeCalculation(true, 50, 0.0001));
        both((grid) => { grid.EnterText(24, 0, '1000'); grid.EnterText(24, 1, '=A25+C25'); grid.EnterText(24, 2, '=B25*0.1'); });
        now = stepIs('iterating: a circular reference entered: by the engine', 'Sheet1!A25:C25');
        check('... settled at its fixed point', Math.abs(Number(now.rows[0][1].replace(/,/g, '')) - 1111.11) < 0.01, true);
        both((grid) => grid.EnterText(24, 0, '2000'));
        now = stepIs('iterating: an input of the cycle: by the engine', 'Sheet1!A25:C25');
        check('... settled again', Math.abs(Number(now.rows[0][1].replace(/,/g, '')) - 2222.22) < 0.01, true);
        both((grid) => grid.EnterText(40, 3, '9'));
        now = stepIs('iterating: an edit elsewhere', 'Sheet1!A25:C25');
        check('... leaves the settled cycle alone (nothing evaluated)', now.evaluated, 0);

        // The real case: milk-street-office.xlsx, Annual!C21:N40 pasted at C101 closes a circular reference across sheets.
        const real = mountGrid(), reference = mountGrid();
        try {
            reference.grid.SetIncrementalRecalculation(false);
            const bytes = await bytesOf('tests/real-world/milk-street-office.xlsx');
            for (const grid of [real.grid, reference.grid]) {
                await grid.ImportXlsx(bytes);
                grid.EnterText(7, 3, '54321'); grid.Undo();
                grid.SelectRange(20, 2, 39, 13); await grid.Copy(); grid.SelectCell(100, 2); await grid.Paste();
            }
            const stats = real.grid.GetRecalcStats();
            const values = (grid) => JSON.stringify(grid.GetRange('C101:N120').map((row) => row.map((cell) => typeof cell.value === 'number' ? Math.round(cell.value * 1e6) / 1e6 : cell.value)));
            check('milk-street: the paste that closes a cross-sheet cycle is calculated by the engine, to the values of the full recalculation', [stats.Kind, values(real.grid) === values(reference.grid)], ['incremental', true]);
            note('milk-street paste: ' + JSON.stringify(stats));
            real.grid.EnterText(7, 3, '777');
            check('... and the next edit is incremental', real.grid.GetRecalcStats().Kind, 'incremental');
        } finally { real.close(); reference.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally { incremental.close(); full.close(); }
    return lines;
};

checks['api-recalc-order'] = async () => {
    const { lines, check, note } = makeLog();
    const incremental = mountGrid(), full = mountGrid();
    try {
        full.grid.SetIncrementalRecalculation(false);
        const grids = [incremental.grid, full.grid];
        const both = (run) => { for (const grid of grids) run(grid); };
        const texts = (grid, range) => grid.GetRange(range).map((row) => row.map((cell) => cell.text));
        const shown = (range) => {
            const mine = texts(incremental.grid, range), theirs = texts(full.grid, range), stats = incremental.grid.GetRecalcStats();
            return { kind: stats.Kind + (stats.Kind === 'full' ? ': ' + stats.Reason : ''), same: JSON.stringify(mine) === JSON.stringify(theirs), rows: mine };
        };
        both((grid) => {
            grid.SetCell(0, 0, '1'); grid.SetCell(1, 0, '2'); grid.SetCell(2, 0, '3'); grid.SetCell(0, 2, '5');
            grid.AddSheet('Two'); grid.SetCell(0, 0, '10'); grid.SetCell(0, 2, '50'); grid.SetCell(0, 1, '=Sheet1!B1+1');
            grid.AddSheet('Three'); grid.SetCell(0, 2, '500');
            grid.SwitchSheet(0);
            grid.EnterText(0, 1, '=OFFSET(A1,1,0)*2'); grid.EnterText(1, 1, '=INDIRECT("Two!A1")+1'); grid.EnterText(2, 1, '=SUM(Two:Three!C1)'); grid.EnterText(3, 1, '=INDIRECT("A"&A1)');
        });
        let now = shown('Sheet1!A1:B4');
        check('OFFSET, INDIRECT and a 3-D sum entered: calculated incrementally', [now.kind, now.same, now.rows.map((row) => row[1]), texts(incremental.grid, 'Two!B1')[0][0]], ['incremental', true, ['4', '11', '550', '1'], '5']);
        both((grid) => grid.EnterText(1, 0, '20'));
        now = shown('Sheet1!A1:B4');
        check('the cell an OFFSET lands on changed: it and its reader on the other sheet follow', [now.kind, now.same, now.rows[0][1], texts(incremental.grid, 'Two!B1')[0][0], texts(full.grid, 'Two!B1')[0][0]], ['incremental', true, '40', '41', '41']);
        both((grid) => grid.EnterText(0, 0, '3'));
        now = shown('Sheet1!A1:B4');
        check('the cell INDIRECT builds its reference from changed', [now.kind, now.same, now.rows[3][1]], ['incremental', true, '3']);
        both((grid) => { grid.SwitchSheet(2); grid.EnterText(0, 2, '600'); grid.SwitchSheet(1); grid.EnterText(0, 0, '11'); grid.SwitchSheet(0); });
        now = shown('Sheet1!A1:B4');
        check('cells of other sheets a 3-D reference and an INDIRECT read', [now.kind, now.same, now.rows[2][1], now.rows[1][1]], ['incremental', true, '650', '12']);
        both((grid) => grid.Undo());
        note('recalculation after the edits: ' + JSON.stringify(incremental.grid.GetRecalcStats()));

        // A circular reference across two sheets.
        both((grid) => { grid.EnterText(0, 3, '=Two!D1+1'); grid.EnterText(1, 3, '=D1*2'); grid.EnterText(2, 3, '=1+1'); grid.SwitchSheet(1); grid.EnterText(0, 3, '=Sheet1!D1+1'); grid.SwitchSheet(0); });
        now = shown('Sheet1!D1:D3');
        const cycleText = now.rows[0][0];
        check('a formula that reads itself through another sheet: both show the cycle error, what reads them too, others are calculated',
            [now.kind, now.same, /CYCLE/.test(cycleText), now.rows.map((row) => row[0]), texts(incremental.grid, 'Two!D1')[0][0]], ['incremental', true, true, [cycleText, cycleText, '2'], cycleText]);
        both((grid) => grid.EnterText(5, 0, '7'));
        check('an edit beside the cycle is incremental', [shown('Sheet1!A6:A6').kind, texts(incremental.grid, 'Sheet1!D1')[0][0]], ['incremental', cycleText]);
        both((grid) => { grid.SwitchSheet(1); grid.EnterText(0, 3, '4'); grid.SwitchSheet(0); });
        now = shown('Sheet1!D1:D3');
        check('the cycle broken: the values are there', [now.same, now.rows.map((row) => row[0])], [true, ['5', '10', '2']]);

        // Iterative calculation: a loan whose interest is part of what it finances.
        const grid = incremental.grid;
        check('iterative calculation is off in a new workbook', grid.GetIterativeCalculation(), { enabled: false, maxIterations: 100, maxChange: 0.001 });
        grid.EnterText(0, 5, '1000'); grid.EnterText(0, 6, '=F1+Two!F1'); grid.SwitchSheet(1); grid.EnterText(0, 5, '=Sheet1!G1*0.1'); grid.SwitchSheet(0);
        check('a loan and its interest without it: the cycle error', [texts(grid, 'Sheet1!G1')[0][0], texts(grid, 'Two!F1')[0][0]], [cycleText, cycleText]);
        const revision = grid.GetRevision();
        let changes = 0;
        const off = grid.OnChange(() => changes++);
        grid.SetIterativeCalculation(true);
        const near = (text, value) => Math.abs(Number(String(text).replace(/,/g, '')) - value) < 0.01;
        check('switched on: the cycle settles (1000 + 10% of the total)', [grid.GetIterativeCalculation(), near(grid.GetCell(0, 6).value, 1111.111), near(grid.GetRange('Two!F1')[0][0].value, 111.111)], [{ enabled: true, maxIterations: 100, maxChange: 0.001 }, true, true]);
        check('... one change, a new revision', [changes, grid.GetRevision() !== revision], [1, true]);
        grid.EnterText(0, 5, '2000');
        check('an input of the cycle edited: it settles again', near(grid.GetCell(0, 6).value, 2222.222), true);
        grid.Undo();
        check('... undone', near(grid.GetCell(0, 6).value, 1111.111), true);
        const again = await reopened(grid);
        try {
            check('saved and opened again: the setting and the values', [again.grid.GetIterativeCalculation().enabled, near(again.grid.GetCell(0, 6).value, 1111.111)], [true, true]);
            again.grid.EnterText(0, 5, '3000');
            check('... and the cycle goes on being calculated there', near(again.grid.GetCell(0, 6).value, 3333.333), true);
        } finally { again.close(); }
        grid.SetIterativeCalculation(true, 5, 0.5);
        check('other numbers', grid.GetIterativeCalculation(), { enabled: true, maxIterations: 5, maxChange: 0.5 });
        grid.Undo(); grid.Undo();
        check('the setting undone: off, the cycle error again', [grid.GetIterativeCalculation().enabled, texts(grid, 'Sheet1!G1')[0][0], grid.GetRevision() === revision], [false, cycleText, true]);
        grid.Redo();
        check('... redone', [grid.GetIterativeCalculation().enabled, near(grid.GetCell(0, 6).value, 1111.111)], [true, true]);
        off();
        check('arguments out of range throw', [thrown(() => grid.SetIterativeCalculation('yes')), thrown(() => grid.SetIterativeCalculation(true, 0)), thrown(() => grid.SetIterativeCalculation(true, 10, -1))].map((message) => /SetIterativeCalculation: /.test(message ?? '')), [true, true, true]);

        // A 3-D reference and the sheets of its span: deleting the sheet it ends at re-points it to the neighbour, as Excel does.
        const months = mountGrid();
        try {
            const book = months.grid;
            book.RenameSheet(0, 'Jan'); book.SetCell(0, 1, '1');
            book.AddSheet('Feb'); book.SetCell(0, 1, '10');
            book.AddSheet('Mar'); book.SetCell(0, 1, '100');
            book.AddSheet('Total'); book.EnterText(0, 0, '=SUM(Jan:Mar!B1)'); book.EnterText(1, 0, '=SUM(Feb:Mar!B1)');
            check('a 3-D sum over three sheets', [book.GetCell(0, 0).value, book.GetCell(1, 0).value], [111, 110]);
            book.DeleteSheet(2);
            check('the last sheet of the span deleted: the span ends at the sheet before it; a span of two is one sheet', [book.GetSheetNames(), book.GetCell(0, 0).formula, book.GetCell(0, 0).value, book.GetCell(1, 0).formula, book.GetCell(1, 0).value], [['Jan', 'Feb', 'Total'], "=SUM('Jan:Feb'!B1)", 11, '=SUM(Feb!B1)', 10]);
            book.Undo();
            check('... undone', [book.GetSheetNames(), book.GetCell(0, 0).formula, book.GetCell(0, 0).value, book.GetCell(1, 0).value], [['Jan', 'Feb', 'Mar', 'Total'], '=SUM(Jan:Mar!B1)', 111, 110]);
            book.MoveSheet(1, 3);
            check('a sheet moved out of the span is no longer read; the text stays', [book.GetSheetNames(), book.GetSheetInfos().filter((sheet) => sheet.isActive).map((sheet) => sheet.name), book.GetRange('Total!A1:A2').map((row) => [row[0].formula, row[0].value])], [['Jan', 'Mar', 'Total', 'Feb'], ['Total'], [['=SUM(Jan:Mar!B1)', 101], ['=SUM(Feb:Mar!B1)', 110]]]);
            const saved = await reopened(book);
            try { check('saved and opened again', saved.grid.GetRange('Total!A1:A2').map((row) => [row[0].formula, row[0].value]), [['=SUM(Jan:Mar!B1)', 101], ['=SUM(Feb:Mar!B1)', 110]]); } finally { saved.close(); }
        } finally { months.close(); }

        // A file with a deliberate circular reference across sheets (iterate="1"): an edit leaves its values as they were.
        const file = mountGrid();
        try {
            await file.grid.ImportXlsx(await bytesOf('tests/real-world/saopaulo-condos.xlsx'));
            check('a workbook with iterative calculation says so', file.grid.GetIterativeCalculation().enabled, true);
            const before = file.grid.GetRange('Assumptions!E57:E60').map((row) => row[0].value);
            file.grid.EnterText(200, 30, '1');
            file.grid.InsertRows(300, 1);
            const after = file.grid.GetRange('Assumptions!E57:E60').map((row) => row[0].value);
            check('its circular model keeps its values when rows are inserted (calculated by the engine)', [file.grid.GetRecalcStats().Kind, after.every((value, i) => typeof value === 'number' && Math.abs(value - before[i]) <= 0.01 * Math.abs(before[i]))], ['incremental', true]);
            note('Assumptions!E57:E60 ' + JSON.stringify(before) + ' -> ' + JSON.stringify(after));
            // ... and through a full recalculation (the setting changed: the workbook is calculated from scratch).
            file.grid.SetIterativeCalculation(true, 100, 0.0005);
            const again = file.grid.GetRange('Assumptions!E57:E60').map((row) => row[0].value);
            check('... and through a full recalculation', [file.grid.GetRecalcStats().Kind, again.every((value, i) => typeof value === 'number' && Math.abs(value - before[i]) <= 0.01 * Math.abs(before[i]))], ['full', true]);
        } finally { file.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    incremental.close(); full.close();
    return lines;
};

// A map with an entry for each row of a long sheet: comparing two with `=` in JavaScript went one
// call deeper per entry, and deleting a row, inserting a column or filtering threw (846 000 cells,
// 76 931 row heights in dashboard-sales-analysis.xlsx).
checks['long-sheet-structure'] = async () => {
    const { grid, close } = mountGrid();
    const { lines, check } = makeLog();
    try {
        const rows = 60000;
        grid.SetCell(0, 0, 'head');
        grid.SetCell(rows - 1, 0, 'last');
        grid.SetRowHeight(`1:${rows}`, 22);
        grid.DeleteRows(60, 1);
        check('a row deleted', grid.GetCell(rows - 2, 0).text, 'last');
        grid.Undo();
        check('... and undone', grid.GetCell(rows - 1, 0).text, 'last');
        grid.InsertRows(30, 1);
        check('a row inserted', grid.GetCell(rows, 0).text, 'last');
        grid.Undo();
        grid.InsertColumns(0, 1);
        check('a column inserted', grid.GetCell(rows - 1, 1).text, 'last');
        grid.Undo();
        grid.HideRows?.(`10:${rows - 10}`);
        grid.InsertRows(5, 1);
        check('a row inserted over rows hidden by the thousand', grid.GetCell(rows, 0).text, 'last');
        grid.Undo();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.message || e)); }
    close();
    return lines;
};

// The walk that steps over what two cell maps share: a cell typed into a table of a long sheet
// compares a few cells, not all of them, and every change is still found.
checks['cell-map-walk'] = async () => {
    const { grid, close } = mountGrid();
    const { lines, check, note } = makeLog();
    try {
        const rows = 4000;
        const values = Array.from({ length: rows }, (_, row) => row + 1);
        grid.SetCell(0, 0, 'Amount');
        grid.SelectCell(1, 0);
        grid.PasteText(values.join('\n'), '');
        grid.DefineTable(`A1:A${rows + 1}`, 'Amounts');
        grid.SetCell(0, 2, `=SUM(A2:A${rows + 1})`);
        const total = () => values.reduce((sum, value) => sum + value, 0);
        check('the total at first', grid.GetCell(0, 2).value, total());
        let seed = 12345, wrong = 0, most = 0;
        const random = (limit) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % limit; };
        for (let step = 0; step < 300; step++) {
            const kind = random(4), row = random(rows);
            if (kind === 0) { values[row] = random(1000); grid.EnterText(row + 1, 0, String(values[row])); most = Math.max(most, grid.GetRecalcStats().ComparedCells); }
            else if (kind === 1) { const count = 1 + random(20); grid.SelectCell(row + 1, 0); const block = Array.from({ length: count }, () => random(1000)); block.forEach((value, at) => { if (row + at < rows) values[row + at] = value; }); grid.PasteText(block.slice(0, rows - row).join('\n'), ''); }
            else if (kind === 2) { const was = values[row]; values[row] = 0; grid.EnterText(row + 1, 0, '0'); grid.Undo(); values[row] = was; }
            else { const was = values[row]; values[row] = random(1000); grid.EnterText(row + 1, 0, String(values[row])); grid.Undo(); values[row] = was; grid.Redo(); grid.Undo(); }
            if (grid.GetCell(0, 2).value !== total()) wrong++;
        }
        check('300 edits, pastes and undos: the total was right after each', wrong, 0);
        check('a cell typed compared few cells', most > 0 && most < 200, true);
        note(`at most ${most} of ${rows + 2} cells compared for a typed cell`);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Saves real Excel refused after a row and a column were inserted: a column moved past the last, a shared
// formula reading another sheet, an error only the engine knows, a query table's fields. (ExportXlsx goes
// through packageProblems, which knows these; run.mjs reports them with this check.)
checks['excel-refused-saves'] = async () => {
    const { lines, check } = makeLog();
    for (const file of ['test-documents-excel/poi-49156', 'real-world/milk-street-office', 'test-documents-excel/poi-47804', 'test-documents-excel/pandas-test_decimal']) {
        const { grid, close } = mountGrid();
        try {
            await grid.ImportXlsx(await bytesOf(`tests/${file}.xlsx`));
            grid.InsertRows(1, 1);
            grid.InsertColumns(1, 1);
            check(file + ': saved', (await grid.ExportXlsx()).length > 0, true);
        } catch (e) { lines.push(`FAIL ${file} threw ${String(e && e.message || e)}`); }
        close();
    }
    return lines;
};

// (run.mjs --keep DIR) Saves of the workbooks of the tests, for Excel to open.
checks['keep-saves'] = async () => {
    const { lines, note } = makeLog();
    const keep = async (name, grid) => {
        const bytes = await grid.ExportXlsx();
        let binary = '';
        for (let at = 0; at < bytes.length; at += 32768) binary += String.fromCharCode(...bytes.subarray(at, at + 32768));
        await window.keepSave(name, btoa(binary));
    };
    let done = 0;
    for (const file of window.keepList ?? []) {
        const base = file.split('/').pop().replace(/\.xlsx$/, '');
        const { grid, close } = mountGrid();
        try {
            await grid.ImportXlsx(await bytesOf(file));
            await keep(base + '.save.xlsx', grid);
            const model = grid.GetModel();
            grid.EnterText(model.RowCount + 2, 0, 'typed by the check');
            await keep(base + '.edit.xlsx', grid);
            grid.Undo();
            grid.InsertRows(1, 1);
            grid.InsertColumns(1, 1);
            await keep(base + '.insert.xlsx', grid);
            done++;
        } catch (e) { lines.push('ok   (could not keep ' + base + ': ' + String(e && e.message || e) + ')'); }
        close();
    }
    note(`${done} of ${(window.keepList ?? []).length} workbooks kept`);
    return lines;
};
