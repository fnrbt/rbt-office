// Copy and paste with the real keyboard (driven by run.mjs through Playwright):
// Ctrl+C / Ctrl+V reach the browser's own copy and paste events on the input
// proxy - no clipboard permission is asked - and what is pasted has the
// formatting it was copied with.
import { check, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const editor = Word.initialize('c');
    window.afterTyping = () => {
        // "Hello world" typed; "world" made bold, then selected for the copy.
        const id = Word.getLayoutLines(editor)[0].BlockId;
        Word.setSelection(editor, id, 6, id, 11);
        Word.toggleBold(editor);
        return [];
    };
    window.finish = () => {
        try {
            const id = Word.getLayoutLines(editor)[0].BlockId;
            check('pasted at the end', Word.getText(editor), 'Hello world world');
            Word.setSelection(editor, id, 12, id, 17);
            check('... bold, as it was copied', Word.getSelectionFormat(editor).Bold, true);
            Word.setSelection(editor, id, 0, id, 6);
            check('... the text that was not copied is not', Word.getSelectionFormat(editor).Bold, false);
            check('one undo step', [Word.undo(editor), Word.getText(editor)], [true, 'Hello world ']);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [['click', 300, 120], ['type', 'Hello world'], ['then', 'window.afterTyping()'],
        ['press', 'Control+c'], ['press', 'End'], ['type', ' '], ['press', 'Control+v']];
} catch (e) { fail(e); finish(); }
