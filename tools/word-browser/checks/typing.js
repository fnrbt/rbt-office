// Real mouse and keyboard input (driven by run.mjs through Playwright): a click
// puts the focus on the input proxy and ordinary typing, Enter, Backspace and
// shortcuts keep working through it.
import { check, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const editor = Word.initialize('c');
    window.finish = () => {
        try {
            const proxy = Word.getInputElement(editor);
            check('typed text', Word.getText(editor), 'Hello wörld\ntwo');
            check('focus is on the input proxy', document.activeElement === proxy, true);
            check('nothing is left in the proxy', proxy.value, '');
            const rect = proxy.getBoundingClientRect();
            check('proxy sits at the caret, inside the page', rect.left > 90 && rect.left < 400 && rect.top > 90 && rect.top < 400, true);
            check('undo works after typing', [Word.undo(editor), Word.getText(editor)], [true, 'Hello wörld\ntwo!']);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [['click', 300, 120], ['type', 'Hello wörld'], ['press', 'Enter'], ['type', 'two!'], ['press', 'Backspace']];
} catch (e) { fail(e); finish(); }
