// A character format chosen with nothing selected is the format of what is
// typed next at the caret: Ctrl+B (or toggleBold, a font, a size, a colour
// from a toolbar) at a caret changes nothing in the document, getSelectionFormat
// says it, the listeners hear of it (FormatChanged), the text typed next has
// it - and it is dropped when the caret moves. With text selected the
// commands change the selection, as before.
import { check, note, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initialize('c');
    Word.setParagraphs(editor, ['Hello world', 'Second paragraph']);
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 700 });
    const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
    const caret = (offset, block = 0) => Word.setSelection(editor, ids[block], offset, ids[block], offset);
    const formatOf = (from, to, block = 0) => { Word.setSelection(editor, ids[block], from, ids[block], to); const f = Word.getSelectionFormat(editor); return f; };
    const text = () => Word.getText(editor).split('\n')[0];
    const heard = [];
    Word.onChange(editor, (change) => heard.push(change));

    // ---- by the API ----
    caret(5);
    const revisionBefore = Word.getRevision ? Word.getRevision(editor) : 0;
    heard.length = 0;
    Word.toggleBold(editor);
    check('bold at a caret: the format of the caret is bold, the document is as it was', [Word.getSelectionFormat(editor).Bold, text(), Word.canUndo(editor), Word.getSelection(editor).IsCollapsed], [true, 'Hello world', false, true]);
    check('the listeners hear that the format changed, not the document', heard.map((change) => [change.FormatChanged === true, change.DocumentChanged, change.SelectionChanged]), [[true, false, true]]);
    Word.insertText(editor, 'XY');
    check('what is typed next is bold', [text(), Word.getSelectionFormat(editor).Bold], ['HelloXY world', true]);
    check('the typed text is bold, the text around it is not', [formatOf(5, 7).Bold, formatOf(0, 5).Bold, formatOf(7, 13).Bold], [true, false, false]);
    caret(7);
    Word.insertText(editor, 'Z');
    check('typing on after it goes on in bold (the text before the caret is)', [text(), formatOf(5, 8).Bold], ['HelloXYZ world', true]);
    check('undo takes the typed text back, one step for each insertion', [Word.undo(editor), text(), Word.undo(editor), text()], [true, 'HelloXY world', true, 'Hello world']);

    // ---- dropped when the caret moves ----
    caret(5);
    Word.toggleItalic(editor);
    check('italic at a caret', Word.getSelectionFormat(editor).Italic, true);
    caret(2);
    check('the caret moved: not italic there', Word.getSelectionFormat(editor).Italic, false);
    caret(5);
    check('and not when it comes back', Word.getSelectionFormat(editor).Italic, false);
    Word.insertText(editor, 'q');
    check('what is typed then is plain', [text(), formatOf(5, 6).Italic], ['Helloq world', false]);
    Word.undo(editor);

    // ---- twice: off again; several at once ----
    caret(5);
    Word.toggleBold(editor); Word.toggleBold(editor);
    check('bold chosen twice: off again', Word.getSelectionFormat(editor).Bold, false);
    Word.toggleBold(editor); Word.toggleUnderline(editor); Word.setFontSizePt(editor, 20); Word.setColor(editor, '#FF0000'); Word.setFontFamily(editor, 'Courier New');
    const pending = Word.getSelectionFormat(editor);
    check('bold, underline, 20 pt, red, a font: the caret has them all', [pending.Bold, pending.Underline, pending.FontSizePt, String(pending.Color).toUpperCase(), pending.FontFamily, text()], [true, true, 20, '#FF0000', 'Courier New', 'Hello world']);
    Word.insertText(editor, 'W');
    const typed = formatOf(5, 6), before = formatOf(0, 5);
    check('and so has the text typed next', [typed.Bold, typed.Underline, typed.FontSizePt, String(typed.Color).toUpperCase(), typed.FontFamily], [true, true, 20, '#FF0000', 'Courier New']);
    check('the text before it has not', [before.Bold, before.Underline, before.FontSizePt === 20, before.FontFamily === 'Courier New'], [false, false, false, false]);
    Word.undo(editor);

    // ---- bold where the text is bold: off for what is typed next ----
    Word.setSelection(editor, ids[1], 0, ids[1], 6);
    Word.toggleBold(editor);
    check('with text selected the selection is made bold, as before', [Word.getSelectionFormat(editor).Bold, Word.canUndo(editor)], [true, true]);
    caret(6, 1);
    check('a caret after bold text is bold', Word.getSelectionFormat(editor).Bold, true);
    Word.toggleBold(editor);
    check('bold chosen there: off', Word.getSelectionFormat(editor).Bold, false);
    Word.insertText(editor, 'n');
    check('the text typed next is not bold, the text before it still is', [formatOf(6, 7, 1).Bold, formatOf(0, 6, 1).Bold], [false, true]);

    // ---- by the keyboard ----
    const line = Word.getLayoutLines(editor).find((item) => item.BlockId === ids[0]);
    const box = canvas.getBoundingClientRect();
    const state = {};
    window.steps = {
        clicked: () => { state.offset = Word.getSelection(editor).FocusOffset; state.plain = Word.getSelectionFormat(editor).Bold; },
        chosen: () => { state.chosen = [Word.getSelectionFormat(editor).Bold, Word.getSelectionFormat(editor).Italic, text()]; },
        typed: () => { state.typed = text(); },
        moved: () => { state.moved = [Word.getSelectionFormat(editor).Bold, text()]; },
    };
    window.drive = [
        ['click', Math.round(box.left + line.X + line.Width + 30), Math.round(box.top + line.Y + line.Height / 2)], ['eval', 'window.steps.clicked()'],
        ['press', 'Control+b'], ['press', 'Control+i'], ['eval', 'window.steps.chosen()'],
        ['type', 'ab'], ['eval', 'window.steps.typed()'],
        ['press', 'Control+b'], ['press', 'ArrowLeft'], ['press', 'ArrowRight'], ['type', 'c'], ['eval', 'window.steps.moved()'],
    ];
    window.finish = () => {
        try {
            check('a click at the end of the first paragraph', [state.offset, state.plain], [11, false]);
            check('Ctrl+B, Ctrl+I: the caret is bold and italic, nothing is typed', state.chosen, [true, true, 'Hello world']);
            const typed = formatOf(11, 13);
            check('typing: bold italic text', [state.typed, typed.Bold, typed.Italic, formatOf(0, 11).Bold], ['Hello worldab', true, true, false]);
            // (Ctrl+B after bold text: off; the arrows drop that, so "c" is bold as the text before it.)
            check('Ctrl+B, then the caret moved away and back: the choice is dropped', [state.moved[1], formatOf(13, 14).Bold], ['Hello worldabc', true]);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
