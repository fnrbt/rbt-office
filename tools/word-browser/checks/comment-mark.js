// The mark of a comment ("[Comment 5]", the place of the w:commentReference in
// the text) is one thing: the caret is before it or after it. Text typed into
// its middle used to part it - two marks of one comment, two references in
// the file.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
let editor, Word, OF, anchor;
const got = {};
try {
    OF = await loadFormats();
    Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/rich-notes.docx'));
    const open = async () => {
        if (editor) Word.dispose(editor);
        editor = Word.initializeCanvas(document.getElementById('c'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.completeLayout(editor);
        anchor = Word.getCommentAnchors(editor)[0];
        Word.getInputElement(editor).focus();
    };
    const focus = () => { const s = Word.getSelection(editor); return [s.AnchorOffset - anchor.StartOffset, s.FocusOffset - anchor.StartOffset]; };
    const marks = () => Word.getCommentAnchors(editor).map((a) => [a.Id, a.EndOffset - a.StartOffset]);
    const references = async () => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        return xml.split('<w:commentReference w:id="5"/>').length - 1;
    };
    window.stage = async (name) => {
        if (name === 'api') {
            await open();
            const id = anchor.BlockId, s = anchor.StartOffset, length = anchor.EndOffset - anchor.StartOffset;
            got.mark = [anchor.Id, length, Word.getText(editor).includes('[Comment 5]')];
            got.snapped = [3, length - 3].map((at) => { Word.setSelection(editor, id, s + at, id, s + at); return focus(); });
            Word.setSelection(editor, id, s - 2, id, s + 4); got.range = focus();
            // Typed at a place in the mark.
            Word.setSelection(editor, id, s + 4, id, s + 4);
            Word.insertText(editor, 'X');
            got.typed = [Word.getText(editor).includes('X[Comment 5]'), marks(), await references()];
            // Bold chosen there, then typed.
            Word.setSelection(editor, id, s + 1 + length - 3, id, s + 1 + length - 3);
            Word.toggleBold(editor); Word.insertText(editor, 'Q'); Word.insertText(editor, 'Q');
            got.typedBold = [Word.getText(editor).includes('X[Comment 5]QQ'), marks(), await references()];
        }
        if (name === 'keys') { await open(); Word.setSelection(editor, anchor.BlockId, anchor.StartOffset, anchor.BlockId, anchor.StartOffset); }
        if (name === 'right') got.right = focus();
        if (name === 'left') got.left = focus();
        if (name === 'shift') got.shift = focus();
        if (name === 'backspace') got.backspace = [Word.getText(editor).includes('Comment 5'), marks(), Word.getText(editor).includes('and a comment')];
        if (name === 'undo') got.undo = [Word.undo(editor), Word.getText(editor).includes('[Comment 5]'), marks().length];
        Word.getInputElement(editor).focus();
    };
    window.drive = [
        ['eval', 'window.stage("api")'],
        ['eval', 'window.stage("keys")'], ['press', 'ArrowRight'], ['eval', 'window.stage("right")'], ['press', 'ArrowLeft'], ['eval', 'window.stage("left")'],
        ['press', 'Shift+ArrowRight'], ['eval', 'window.stage("shift")'], ['press', 'ArrowRight'], ['press', 'Backspace'], ['eval', 'window.stage("backspace")'], ['eval', 'window.stage("undo")'],
    ];
} catch (e) { fail(e); finish(); }
window.finish = () => {
    try {
        const length = '[Comment 5]'.length;
        check('the comment of the document has its mark in the text', got.mark, ['5', length, true]);
        check('setSelection into the mark is its nearer end', got.snapped, [[0, 0], [length, length]]);
        check('a selection that ends in the mark has all of the mark\'s nearer part: its start', got.range, [-2, 0]);
        check('text typed at a place in the mark stands before it: one mark, one reference in the file', got.typed, [true, [['5', length]], 1]);
        check('bold chosen at a place in the mark and two characters typed: after it, one mark, one reference', got.typedBold, [true, [['5', length]], 1]);
        check('ArrowRight at the start of the mark goes to its end', got.right, [length, length]);
        check('ArrowLeft at its end goes to its start', got.left, [0, 0]);
        check('Shift+ArrowRight at its start selects it', got.shift, [0, length]);
        check('Backspace after the mark removes it whole', got.backspace, [false, [], true]);
        check('one undo brings it back', got.undo, [true, true, 1]);
    } catch (e) { fail(e); }
    finish();
};
