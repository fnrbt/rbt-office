// Positions a host gives are made positions of the document: setSelection
// with an offset beyond the end of the paragraph's text puts the caret at the
// end (the next keystroke used to throw - charCodeAt of undefined), one below
// 0 at the start, no number at all at 0; a block id the document does not
// have is refused (false). The range queries take offsets the same way.
// And what cannot be done where the caret is, is refused: a page break in a
// text box, a header or a footer, a text box in a text box (false, and
// nothing changes).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initialize('c');
    Word.setParagraphs(editor, ['Hello', 'World']);
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 700 });
    const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
    const selection = () => { const s = Word.getSelection(editor); return [ids.indexOf(s.AnchorBlockId), s.AnchorOffset, ids.indexOf(s.FocusBlockId), s.FocusOffset]; };
    const text = () => Word.getText(editor);
    const state = {};
    window.steps = {
        past: () => { state.past = [Word.setSelection(editor, ids[0], 999, ids[0], 999), selection(), Word.getCaretRect(editor) !== null]; },
        typed: () => { state.typed = [text(), selection()]; },
    };
    const line = Word.getLayoutLines(editor)[0];
    const box = canvas.getBoundingClientRect();
    window.drive = [
        ['click', Math.round(box.left + line.X + 5), Math.round(box.top + line.Y + line.Height / 2)],
        ['eval', 'window.steps.past()'],
        ['type', 'x'], ['press', 'ArrowRight'], ['press', 'Backspace'], ['type', 'yz'], ['press', 'ArrowLeft'], ['press', 'Delete'],
        ['eval', 'window.steps.typed()'],
    ];
    window.finish = async () => {
        try {
            check('an offset beyond the end of the text: the end of the paragraph, with a caret', state.past, [true, [0, 5, 0, 5], true]);
            // ("x" at the end, Right into the next paragraph, Backspace joins them, "yz", Left, Delete takes the "z".)
            check('the keys after it work on the text', state.typed, ['HelloxyWorld', [0, 7, 0, 7]]);
            Word.setParagraphs(editor, ['Hello', 'World']);
            const again = [...new Set(Word.getLayoutLines(editor).map((item) => item.BlockId))];
            ids.splice(0, ids.length, ...again);
            check('offsets below 0 and beyond the end, across paragraphs', [Word.setSelection(editor, ids[0], -5, ids[1], 77), selection()], [true, [0, 0, 1, 5]]);
            check('no number: the start; a fraction: cut off', [Word.setSelection(editor, ids[0], undefined, ids[0], NaN), selection(), Word.setSelection(editor, ids[0], 2.7, ids[0], '3'), selection()], [true, [0, 0, 0, 0], true, [0, 2, 0, 0]]);
            Word.setSelection(editor, ids[1], 1, ids[1], 2);
            check('a paragraph the document does not have: refused, the selection stays', [Word.setSelection(editor, 'no-such-block', 0, 'no-such-block', 0), Word.setSelection(editor, ids[0], 0, null, 0), Word.setSelection(editor, undefined, 0, ids[0], 0), selection()], [false, false, false, [1, 1, 1, 2]]);
            const bounds = Word.getRangeBounds(editor, ids[0], 3, 999), caretBounds = Word.getRangeBounds(editor, ids[0], 999, 999), whole = Word.getRangeBounds(editor, ids[0], 0, 5);
            check('getRangeBounds: offsets within the text', [bounds.StartOffset, bounds.EndOffset, caretBounds.StartOffset, Math.round(caretBounds.X) === Math.round(whole.X + whole.Width), Word.getRangeBounds(editor, 'no-such-block', 0, 1)], [3, 5, 5, true, null]);
            const information = Word.getRangeInformation(editor, ids[1], 999, 2000), statistics = Word.computeRangeStatistics(editor, ids[1], -3, 999);
            check('getRangeInformation and computeRangeStatistics answer', [information !== null && typeof information === 'object', statistics !== null && typeof statistics === 'object'], [true, true]);
            Word.selectFindMatch(editor, { BlockId: ids[1], Start: 2, End: 99, Text: 'rld' });
            check('selectFindMatch: within the text', selection(), [1, 2, 1, 5]);

            // ---- what cannot be done where the caret is ----
            const second = Word.initializeCanvas(canvas);
            Word.importFromDocxPackage(second, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-drawingml-wps.docx')));
            const shape = Word.getShapeRects(second).find((item) => item.Text.length > 0);
            const bodyLine = Word.getLayoutLines(second)[0];
            const snapshot = () => [Word.getPageCount(second), Word.getShapeRects(second).length, Word.getText(second), Word.canUndo(second)];
            Word.clearSelection(second);
            check('no selection: no page break', [Word.insertPageBreak(second), Word.canUndo(second)], [false, false]);
            check('the text of the box is entered', [Word.enterTextBox(second, shape.Id), Word.getCaretContext(second).Story], [true, 'textbox']);
            const inBox = snapshot();
            check('in a text box: no page break, no text box', [Word.insertPageBreak(second), Word.insertTextBox(second, 'inner'), snapshot()], [false, false, inBox]);
            Word.exitTextBox(second);
            check('in a header: no page break', [Word.enterHeaderFooter(second, 0, 'header'), Word.getCaretContext(second).Story, Word.insertPageBreak(second), Word.getPageCount(second)], [true, 'header', false, inBox[0]]);
            Word.exitHeaderFooter(second);
            Word.setSelection(second, bodyLine.BlockId, 4, bodyLine.BlockId, 4);
            const before = snapshot();
            check('in the body: a text box is made', [Word.insertTextBox(second, 'made'), Word.getShapeRects(second).length, Word.getShapeRects(second).some((item) => item.Text === 'made')], [true, before[1] + 1, true]);
            Word.setSelection(second, bodyLine.BlockId, 4, bodyLine.BlockId, 4);
            check('and a page break', [Word.insertPageBreak(second), Word.getPageCount(second), Word.getText(second).startsWith('Milk\n')], [true, before[0] + 1, true]);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
