// What a host's table and paragraph tools ask for: a table that was just
// inserted is visible (borders) and the caret is in its first cell with a
// caret rectangle; Tab / Shift+Tab move between cells and Tab in the last
// cell adds a row; getCaretContext says what the caret is in without a
// layout pass; getStyles and the selection format never show the internal
// "__docDefaults" entry; the selection format has line spacing and the space
// before / after; markSaved takes the revision that was written.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initialize('c');
    const context = () => Word.getCaretContext(editor);
    const cellOf = () => { const c = context(); return [c.InTable, c.Row, c.Column, c.RowCount, c.ColumnCount]; };
    const steps = {};
    window.step = (name) => { try { steps[name](); } catch (e) { fail(e); } };

    Word.insertText(editor, 'Before the table');
    check('in a paragraph: not in a table', [context().InTable, context().Row, context().InList, context().InHyperlink, context().InHeaderFooter, context().ImageSelected], [false, -1, false, false, false, false]);
    Word.insertTable(editor, 3, 3);
    const rect = Word.getCaretRect(editor);
    check('after insertTable: a caret rectangle', [rect !== null, rect && rect.Height > 5, rect && rect.Y > 90], [true, true, true]);
    const information = Word.getSelectionInformation(editor);
    check('getSelectionInformation answers, in the table', [information !== null, information && information.wdWithInTable], [true, true]);
    check('the caret is in the first cell of a 3 x 3 table', cellOf(), [true, 0, 0, 3, 3]);
    check('nothing to merge or split', [context().CanMergeCells, context().CanSplitCell], [false, false]);
    // The table is visible: its borders are painted (dark pixels along a line through the first row, one per vertical border).
    const pixels = canvas.getContext('2d').getImageData(0, Math.round(rect.Y + rect.Height / 2), canvas.width, 1).data;
    let borders = 0, inside = false;
    for (let x = 0; x < canvas.width; x++) {
        const dark = pixels[x * 4 + 3] > 0 && pixels[x * 4] < 190 && pixels[x * 4 + 1] < 190 && pixels[x * 4 + 2] < 190;
        if (dark && !inside) borders++;
        inside = dark;
    }
    check('a new table has borders: four vertical lines cross its first row', borders, 4);
    const frame = Word.getPageFrames(editor)[0];
    // (From just above the table's first row down: the text above the table is not in it.)
    const column = canvas.getContext('2d').getImageData(Math.round(rect.X + 30), Math.round(rect.Y - 8), 1, 400).data;
    let rules = 0; inside = false;
    for (let y = 0; y < column.length / 4; y++) {
        const dark = column[y * 4 + 3] > 0 && column[y * 4] < 190 && column[y * 4 + 1] < 190 && column[y * 4 + 2] < 190;
        if (dark && !inside) rules++;
        inside = dark;
    }
    check('and four horizontal lines its first column', rules, 4);

    // Tab / Shift+Tab with the real keyboard (the caret was put in the first cell by insertTable).
    steps.afterTab = () => { check('Tab: the next cell', cellOf(), [true, 0, 1, 3, 3]); check('with a caret rectangle there', Word.getCaretRect(editor).X > rect.X + 50, true); };
    steps.afterTyping = () => { check('typing goes into that cell', Word.getText(editor).includes('two'), true); };
    steps.afterRowEnd = () => check('Tab at the end of a row: the first cell of the next', cellOf(), [true, 1, 0, 3, 3]);
    steps.afterShiftTab = () => check('Shift+Tab: back to the last cell of the row above', cellOf(), [true, 0, 2, 3, 3]);
    steps.atLast = () => check('seven more Tabs: the last cell', cellOf(), [true, 2, 2, 3, 3]);
    steps.afterLast = () => {
        check('Tab in the last cell adds a row and goes to its first cell', cellOf(), [true, 3, 0, 4, 3]);
        check('with a caret rectangle', Word.getCaretRect(editor) !== null, true);
    };
    steps.undone = () => check('one undo takes the row away again', [context().RowCount, Word.canUndo(editor)], [3, true]);
    window.finish = async () => {
        try {
            // From the host's side.
            check('moveToAdjacentTableCell(false) from the first cell: nowhere to go', (() => { while (Word.moveToAdjacentTableCell(editor, false)) { /* back to the first cell */ } return cellOf().slice(0, 3); })(), [true, 0, 0]);
            check('moveToAdjacentTableCell(true)', [Word.moveToAdjacentTableCell(editor, true), cellOf().slice(1, 3)], [true, [0, 1]]);
            // Merge: a selection over two cells.
            const t0 = performance.now();
            for (let i = 0; i < 2000; i++) Word.getCaretContext(editor);
            note(`getCaretContext: ${((performance.now() - t0) / 2000 * 1000).toFixed(1)} µs`);

            // The saved revision.
            const written = Word.getRevision(editor);
            Word.insertText(editor, 'typed while saving');
            Word.markSaved(editor, written);
            check('markSaved(revision): what was typed since stays unsaved', Word.isModified(editor), true);
            Word.undo(editor);
            check('and undoing it gives the saved state', [Word.isModified(editor), Word.getRevision(editor)], [false, written]);
            Word.insertText(editor, 'x');
            Word.markSaved(editor);
            check('markSaved() marks the current content', Word.isModified(editor), false);
            Word.insertText(editor, 'y'); Word.markSaved(editor, null);
            check('markSaved(null) too', Word.isModified(editor), false);

            // Paragraph spacing in the selection format.
            let paragraph = Word.getSelectionFormat(editor).Paragraph;
            check('the selection format has the spacing fields', ['LineSpacing', 'LineSpacingRule', 'SpaceBefore', 'SpaceAfter'].filter((name) => !(name in paragraph)), []);
            Word.setLineSpacing(editor, 1.5); Word.setSpacingBefore(editor, 6); Word.setSpacingAfter(editor, 12);
            paragraph = Word.getSelectionFormat(editor).Paragraph;
            check('they follow setLineSpacing / setSpacingBefore / setSpacingAfter', [paragraph.LineSpacing, paragraph.SpaceBefore, paragraph.SpaceAfter], [1.5, 6, 12]);
            Word.setLineSpacing(editor, 0); Word.setSpacingBefore(editor, -1); Word.setSpacingAfter(editor, -1);

            // A document from a file: no internal style entry anywhere.
            for (const file of ['/repo/tests/test-documents-wild/docx4j-getting-started.docx', '/repo/tests/test-documents/sample-formatted.docx']) {
                const short = file.split('/').pop();
                Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf(file)));
                const styles = Word.getStyles(editor);
                check(`${short}: getStyles lists no internal entry`, styles.filter((style) => String(style.Id).startsWith('__') || style.Type === 'defaults').map((style) => style.Id), []);
                check(`${short}: it lists styles`, styles.length > 3, true);
                // Every paragraph of every table cell: the style a user sees, or null.
                const editorDocument = editor.Doc;
                const seen = new Set();
                let cells = 0;
                const visit = (blocks) => {
                    for (const block of blocks) {
                        const table = block.tag === 2 || (block.fields && block.fields[0] && block.fields[0].Rows) ? block.fields[0] : null;
                        if (!table || !table.Rows) continue;
                        for (const row of table.Rows) for (const cell of row.Cells) for (const para of cell.Content) {
                            Word.setSelection(editor, para.Id, 0, para.Id, 0);
                            cells++;
                            seen.add(String(Word.getSelectionFormat(editor).Paragraph.StyleId));
                            if (cells === 1) check(`${short}: the caret in a cell is in a table`, [context().InTable, context().RowCount > 0, context().ColumnCount > 0], [true, true, true]);
                        }
                    }
                };
                visit([...editorDocument.Blocks]);
                note(`${short}: ${cells} cell paragraphs, style ids ${[...seen].join(', ')}`);
                check(`${short}: no cell paragraph reports an internal style`, [...seen].filter((id) => id.startsWith('__')), []);
            }
        } catch (e) { fail(e); }
        finish();
    };
    // (The keys go to the editor's input proxy.)
    Word.getInputElement(editor).focus();
    window.drive = [['press', 'Tab'], ['eval', 'window.step("afterTab")'], ['type', 'two'], ['eval', 'window.step("afterTyping")'], ['press', 'Tab'], ['press', 'Tab'], ['eval', 'window.step("afterRowEnd")'],
        ['press', 'Shift+Tab'], ['eval', 'window.step("afterShiftTab")'], ...Array.from({ length: 6 }, () => ['press', 'Tab']), ['eval', 'window.step("atLast")'], ['press', 'Tab'], ['eval', 'window.step("afterLast")'],
        ['press', 'Control+z'], ['eval', 'window.step("undone")']];
} catch (e) { fail(e); finish(); }
