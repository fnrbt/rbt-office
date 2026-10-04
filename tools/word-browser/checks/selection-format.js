// Toolbar state: getSelectionFormat reports the effective formatting of the
// selection ("mixed" across different values), canUndo / canRedo, and onChange
// fires after document and selection changes (keyboard, pointer, API) until
// it is unsubscribed.
import { check, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const editor = Word.initialize('c');
    const events = [];
    const unsubscribe = Word.onChange(editor, (change) => events.push(change));
    const take = () => events.splice(0).map((e) => (e.DocumentChanged ? 'D' : '') + (e.SelectionChanged ? 'S' : '') + (e.CanUndo ? 'u' : '') + (e.CanRedo ? 'r' : '')).join(' ');
    const select = (blockId, start, end) => Word.setSelection(editor, blockId, start, blockId, end);
    const pick = (f) => [f.Bold, f.Italic, f.Underline, f.Strike];

    Word.setParagraphs(editor, ['plain bold tail', 'second paragraph', 'third']);
    check('loading a document is reported', take().includes('D'), true);
    const ids = [...new Set(Word.getLayoutLines(editor).map((l) => l.BlockId))];
    check('three paragraphs', ids.length, 3);

    select(ids[0], 6, 10);
    check('a selection change is reported, as a selection change only', take(), 'S');
    let f = Word.getSelectionFormat(editor);
    check('plain selection', [pick(f), f.IsCollapsed, f.HasSelection, f.CanUndo, f.CanRedo], [[false, false, false, false], false, true, false, false]);
    check('plain paragraph', [f.Paragraph.Alignment, f.Paragraph.ListKind, f.Paragraph.ListLevel, f.Paragraph.HeadingLevel, f.Paragraph.IndentLeft], ['left', 'none', 0, 0, 0]);
    check('font of plain text', [typeof f.FontFamily, typeof f.FontSize, f.Color, f.Highlight], ['string', 'number', null, null]);

    Word.toggleBold(editor);
    check('a command is reported as a document change, with undo available', take(), 'Du');
    f = Word.getSelectionFormat(editor);
    check('bold after toggleBold', [f.Bold, f.CanUndo, f.CanRedo, Word.canUndo(editor)], [true, true, false, true]);
    select(ids[0], 2, 12);
    f = Word.getSelectionFormat(editor);
    check('plain + bold is mixed, italic is not', [f.Bold, f.Italic], ['mixed', false]);
    select(ids[0], 8, 8);
    check('a caret inside bold text reports bold', [Word.getSelectionFormat(editor).Bold, Word.getSelectionFormat(editor).IsCollapsed], [true, true]);
    take();

    select(ids[0], 11, 15);
    Word.toggleItalic(editor); Word.toggleUnderline(editor); Word.toggleStrike(editor);
    Word.setColor(editor, '#C00000'); Word.setHighlight(editor, 'yellow'); Word.setFontFamily(editor, 'Georgia'); Word.setFontSize(editor, 20);
    f = Word.getSelectionFormat(editor);
    check('character formatting set through the API', [pick(f), f.FontFamily, f.FontSize, f.Color, f.Highlight], [[false, true, true, true], 'Georgia', 20, '#C00000', 'yellow']);
    select(ids[0], 0, 15);
    f = Word.getSelectionFormat(editor);
    check('the whole paragraph is mixed in everything that differs', [f.Bold, f.Italic, f.FontFamily, f.FontSize, f.Color, f.Highlight], ['mixed', 'mixed', 'mixed', 'mixed', 'mixed', 'mixed']);
    take();

    select(ids[1], 3, 3);
    Word.alignCenter(editor); Word.toggleBulletList(editor);
    f = Word.getSelectionFormat(editor);
    check('paragraph formatting', [f.Paragraph.Alignment, f.Paragraph.ListKind, f.Paragraph.ListLevel >= 1], ['center', 'bullet', true]);
    Word.increaseIndent(editor);
    const indented = Word.getSelectionFormat(editor).Paragraph;
    check('indent level follows', indented.ListLevel > f.Paragraph.ListLevel || indented.IndentLeft > f.Paragraph.IndentLeft, true);
    select(ids[2], 0, 0);
    Word.setHeadingLevel(editor, 2);
    f = Word.getSelectionFormat(editor);
    check('heading', [f.Paragraph.HeadingLevel, f.Bold], [2, true]);
    Word.toggleNumberedList(editor);
    check('numbered list', Word.getSelectionFormat(editor).Paragraph.ListKind, 'numbered');
    select(ids[0], 3, 3); take();
    const third = Word.getLayoutLines(editor).filter((l) => l.BlockId === ids[2])[0];
    Word.selectAll(editor);
    f = Word.getSelectionFormat(editor);
    check('across paragraphs', [f.Paragraph.Alignment, f.Paragraph.ListKind, f.Paragraph.HeadingLevel], ['mixed', 'mixed', 'mixed']);
    take();

    Word.undo(editor);
    check('undo is reported with redo available', take(), 'DSur');
    check('canRedo after undo', [Word.canRedo(editor), Word.getSelectionFormat(editor).CanRedo], [true, true]);
    Word.redo(editor); take();

    window.afterInput = () => {
        const log = take();
        check('a click and typing are reported (selection, then document changes)', [log.startsWith('S'), log.includes('D'), events.length], [true, true, 0]);
        unsubscribe();
        window.textAtUnsubscribe = Word.getText(editor);
    };
    window.finish = () => {
        try {
            check('typing still works after unsubscribing', Word.getText(editor).length, window.textAtUnsubscribe.length + 2);
            check('no event after unsubscribing', events.length, 0);
            Word.toggleBold(editor);
            check('nor for API commands', events.length, 0);
            // A listener that unsubscribes itself while being called.
            let calls = 0;
            const off = Word.onChange(editor, () => { calls++; off(); });
            Word.selectAll(editor); Word.toggleItalic(editor);
            check('a listener can unsubscribe from inside its callback', calls, 1);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [['click', third.X + 20, third.Y + third.Height / 2], ['type', 'xy'], ['eval', 'window.afterInput()'], ['type', 'zz']];
} catch (e) { fail(e); finish(); }
