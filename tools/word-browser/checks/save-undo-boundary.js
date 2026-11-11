// A save is an undo boundary: the host reads the revision, exports the document and uploads it while the user
// goes on typing; marked saved at that revision, undoing what was typed after the export gives the saved state
// (isModified false), though the keys typed before and after it were one run of typing. (Real keys.)
import { check, finish, fail, loadWord, loadFormats } from './common.js';
let Word, OF, editor;
const state = {};
try {
    Word = await loadWord();
    OF = await loadFormats();
    editor = Word.initialize('c');
    Word.setParagraphs(editor, ['First paragraph.', 'Second paragraph.']);
    const id = Word.getLayoutLines(editor)[0].BlockId;
    Word.setSelection(editor, id, 0, id, 0);
    Word.getInputElement(editor).focus();
    window.saveStart = () => {
        state.package = Word.exportDocumentPackage(editor);
        state.revision = Word.getRevision(editor);
        state.savedText = Word.getText(editor);
    };
    window.saveEnd = async () => {
        // Serialize the exported snapshot after later typing, as an asynchronous
        // host save may do. The file must contain only the exported revision.
        state.bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(state.package)).arrayBuffer());
        Word.markSaved(editor, state.revision);
        state.afterSave = [Word.isModified(editor), Word.getText(editor).startsWith('FIRST SECOND ')];
    };
    window.afterUndo = () => { state.undone = [Word.isModified(editor), Word.getRevision(editor) === state.revision, Word.getText(editor).startsWith('FIRST First')]; };
    window.afterUndo2 = () => { state.undone2 = [Word.isModified(editor), Word.getText(editor).startsWith('First paragraph.')]; };
    window.drive = [
        ['type', 'FIRST '], ['eval', 'window.saveStart()'],
        ['type', 'SECOND '], ['eval', 'window.saveEnd()'],
        ['press', 'Control+z'], ['eval', 'window.afterUndo()'],
        ['press', 'Control+z'], ['eval', 'window.afterUndo2()'],
    ];
} catch (e) { fail(e); finish(); }
window.finish = async () => {
    try {
        check('typed during the save: modified, both runs of text there', state.afterSave, [true, true]);
        check('one undo takes back what was typed after the export: the saved state, not modified', state.undone, [false, true, true]);
        check('another undo takes back what was typed before it', state.undone2, [true, true]);
        const reopened = Word.initializeCanvas(document.createElement('canvas'));
        try {
            Word.importFromDocxPackage(reopened, await OF.processors.forEditorPackage(state.bytes));
            check('the saved file reopens at the exported revision, without later typing', Word.getText(reopened), state.savedText);
        } finally {
            Word.dispose(reopened);
        }
    } catch (e) { fail(e); }
    finish();
};
