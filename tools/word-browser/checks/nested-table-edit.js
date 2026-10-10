// Editing in a table that is inside a table cell, with real mouse and key
// events. A paragraph of a nested table was laid out, painted and found by a
// click, but no editing function could find it: what was typed there was
// lost, and getText / find did not report its text. Here: a click into a
// nested cell and typing; Enter (a new paragraph in the nested cell) and
// Backspace (joined again); Tab from cell to cell of the nested table, a new
// row from its last cell; undo back to the document as opened; Enter and
// Backspace in a cell of the outer table (above the nested table, which
// stays where it is); and a save: the nested table is written with the edit,
// the paragraphs outside it as the text they were read from.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
    const cell = (width, content) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const nested = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`
        + `<w:tr>${cell(2000, p('inner11'))}${cell(2000, p('inner12'))}</w:tr><w:tr>${cell(2000, p('inner21'))}${cell(2000, p('inner22'))}</w:tr></w:tbl>`;
    const outer = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>`
        + `<w:tr>${cell(4680, p('outerleft') + nested + p('belownested'))}${cell(4680, p('outerright'))}</w:tr></w:tbl>`;
    const body = p('before the table') + outer + p('after the table');
    zip.file('word/document.xml', head + body + sect + '</w:body></w:document>');
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
    const lines = () => Word.getLayoutLines(editor);
    const texts = () => lines().map((line) => line.Text ?? '');
    const line = (text) => lines().find((l) => (l.Text ?? '') === text);
    const context = () => Word.getCaretContext(editor);
    const cellOf = () => { const c = context(); return [c.InTable, c.Row, c.Column, c.RowCount, c.ColumnCount]; };
    const opened = texts();
    const openedText = Word.getText(editor);
    check('the lines: the nested table between the two paragraphs of its cell', opened, ['before the table', 'outerleft', 'inner11', 'inner12', 'inner21', 'inner22', 'belownested', 'outerright', 'after the table']);
    check('getText has the nested table\'s text, in document order', openedText.replace(/\s+/g, ' ').trim(), 'before the table outerleft inner11 inner12 inner21 inner22 belownested outerright after the table');
    const target = line('inner11');
    const steps = {};
    window.step = (name) => { try { steps[name](); } catch (e) { fail(e); } };
    steps.clicked = () => {
        const selection = Word.getSelection(editor);
        check('a click on a nested cell: the caret is in its paragraph', selection && selection.FocusBlockId, target.BlockId);
        check('the caret context: a cell of the 2 x 2 nested table (not of the outer 1 x 2 one)', cellOf(), [true, 0, 0, 2, 2]);
    };
    steps.typed = () => {
        check('typed in the nested cell', texts().includes('inner11Z'), true);
        check('getText reports it', Word.getText(editor).includes('inner11Z'), true);
    };
    steps.entered = () => {
        const all = texts();
        check('Enter and a letter: a second paragraph in the nested cell', [all.indexOf('Q') === all.indexOf('inner11Z') + 1, line('Q') && line('Q').Y > line('inner11Z').Y + 5, line('Q') && Math.abs(line('Q').X - line('inner11Z').X) < 1], [true, true, true]);
        check('still in the first cell of the nested table', cellOf(), [true, 0, 0, 2, 2]);
        check('the rows below moved down', line('inner21').Y > line('Q').Y + 5, true);
    };
    steps.joined = () => {
        check('Backspace twice: one paragraph again', [texts().includes('Q'), texts().includes('inner11Z'), texts().length], [false, true, opened.length]);
        check('the caret is at the end of it', [Word.getSelection(editor).FocusBlockId, Word.getSelection(editor).FocusOffset], [target.BlockId, 8]);
    };
    steps.tabbed = () => check('Tab: the next cell of the nested table', cellOf(), [true, 0, 1, 2, 2]);
    steps.tabbedTyped = () => check('typing goes into that cell', texts().some((text) => text.includes('T') && text.replace('T', '') === 'inner12'), true);
    steps.lastCell = () => check('two more Tabs: the last cell of the nested table', cellOf(), [true, 1, 1, 2, 2]);
    steps.newRow = () => {
        check('Tab in the last cell: a new row in the nested table', cellOf(), [true, 2, 0, 3, 2]);
        check('the outer table has its one row', (() => { Word.setSelection(editor, line('outerright').BlockId, 0, line('outerright').BlockId, 0); const c = cellOf(); return c; })(), [true, 0, 1, 1, 2]);
    };
    steps.undone = () => {
        check('undo, step by step: the document as opened', [texts(), Word.getText(editor) === openedText], [opened, true]);
    };
    steps.outerEntered = () => {
        const all = texts();
        check('Enter in the outer cell above the nested table: a new paragraph, the nested table after it', [all.indexOf('left'), all.indexOf('inner11')], [all.indexOf('outer') + 1, all.indexOf('outer') + 2]);
        check('the caret context: the outer table', cellOf(), [true, 0, 0, 1, 2]);
    };
    steps.outerJoined = () => check('Backspace: joined; the nested table where it was', texts(), opened);
    steps.belowJoin = () => check('Backspace at the start of the paragraph below the nested table: nothing (a table is above it)', texts(), opened);
    const at = (text, dx = 3) => { const l = line(text); return [Math.round(l.X + dx), Math.round(l.Y + 8)]; };
    window.drive = [
        ['click', ...at('inner11')], ['eval', 'window.step("clicked")'],
        ['press', 'End'], ['type', 'Z'], ['eval', 'window.step("typed")'],
        ['press', 'Enter'], ['type', 'Q'], ['eval', 'window.step("entered")'],
        ['press', 'Backspace'], ['press', 'Backspace'], ['eval', 'window.step("joined")'],
        ['press', 'Tab'], ['eval', 'window.step("tabbed")'], ['type', 'T'], ['eval', 'window.step("tabbedTyped")'],
        ['press', 'Tab'], ['press', 'Tab'], ['eval', 'window.step("lastCell")'],
        ['press', 'Tab'], ['eval', 'window.step("newRow")'],
        ...Array.from({ length: 12 }, () => ['press', 'Control+z']), ['eval', 'window.step("undone")'],
        ['then', 'window.outerClick()'],
    ];
    window.outerClick = () => {
        const l = line('outerleft');
        Word.setSelection(editor, l.BlockId, 5, l.BlockId, 5);
        Word.getInputElement(editor).focus();
        return [['press', 'Enter'], ['eval', 'window.step("outerEntered")'], ['press', 'Backspace'], ['eval', 'window.step("outerJoined")'], ['then', 'window.belowClick()']];
    };
    window.belowClick = () => {
        const l = line('belownested');
        Word.setSelection(editor, l.BlockId, 0, l.BlockId, 0);
        return [['press', 'Backspace'], ['eval', 'window.step("belowJoin")'], ['then', 'window.lastEdit()']];
    };
    window.lastEdit = () => [['click', ...at('inner22')], ['press', 'Home'], ['type', 'K']];
    window.finish = async () => {
        try {
            check('before the save: K in the last nested cell', texts().includes('Kinner22'), true);
            const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
            const saved = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
            const inner = saved.slice(saved.indexOf('<w:tbl>', saved.indexOf('<w:tbl>') + 1), saved.indexOf('</w:tbl>'));
            check('saved: the character is in the nested table', [/Kinner22/.test(inner.replace(/<[^>]+>/g, '')), (saved.match(/<w:tbl>/g) ?? []).length], [true, 2]);
            for (const text of ['before the table', 'after the table']) check(`saved: "${text}" is the text it was read from`, saved.includes(p(text)), true);
            for (const text of ['outerleft', 'inner11', 'inner12', 'inner21', 'belownested', 'outerright']) check(`saved: "${text}" is there`, saved.includes(`>${text}</w:t>`), true);
            const second = Word.initializeCanvas(document.createElement('canvas'));
            Word.importFromDocxPackage(second, await OF.processors.forEditorPackage(bytes));
            check('read again: the lines', Word.getLayoutLines(second).map((l) => l.Text ?? ''), opened.map((text) => text === 'inner22' ? 'Kinner22' : text));
            check('read again: the text the editor had', Word.getText(second), Word.getText(editor));
            Word.dispose(second);
        } catch (e) { fail(e); }
        Word.dispose(editor);
        finish();
    };
} catch (e) { fail(e); finish(); }
