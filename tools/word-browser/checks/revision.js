// A document revision that undo and redo restore, and "modified since saved":
// undoing back to the saved state is not modified.
import { check, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const editor = Word.initialize('c');
    const changes = [];
    Word.onChange(editor, (change) => changes.push(change));
    Word.setParagraphs(editor, ['First paragraph.', 'Second paragraph.']);
    const loaded = Word.getRevision(editor);
    check('a loaded document is not modified', [Word.isModified(editor), typeof loaded], [false, 'number']);
    const ids = Word.getLayoutLines(editor).map((line) => line.BlockId);
    Word.setSelection(editor, ids[0], 5, ids[0], 5);
    check('moving the caret is not a revision', [Word.getRevision(editor), Word.isModified(editor)], [loaded, false]);
    Word.insertText(editor, ' typed');
    const afterTyping = Word.getRevision(editor);
    check('an edit is a new revision', [afterTyping > loaded, Word.isModified(editor)], [true, true]);
    check('the change event carries the revision and the modified flag', [changes.at(-1).Revision, changes.at(-1).Modified], [afterTyping, true]);
    Word.undo(editor);
    check('undo restores the revision: back at the loaded state, not modified', [Word.getRevision(editor), Word.isModified(editor)], [loaded, false]);
    Word.redo(editor);
    check('redo restores the later revision', [Word.getRevision(editor), Word.isModified(editor)], [afterTyping, true]);
    // Save here.
    Word.markSaved(editor);
    check('markSaved: not modified', Word.isModified(editor), false);
    Word.setSelection(editor, ids[1], 0, ids[1], 6);
    Word.toggleBold(editor);
    const bold = Word.getRevision(editor);
    Word.insertText(editor, 'X');
    const typedMore = Word.getRevision(editor);
    check('later edits are later revisions', [bold > afterTyping, typedMore > bold, Word.isModified(editor)], [true, true, true]);
    Word.undo(editor);
    check('one undo: still modified', [Word.getRevision(editor), Word.isModified(editor)], [bold, true]);
    Word.undo(editor);
    check('undone back to the saved state: not modified', [Word.getRevision(editor), Word.isModified(editor), changes.at(-1).Modified], [afterTyping, false, false]);
    Word.undo(editor);
    check('undone past the saved state: modified, at the older revision', [Word.getRevision(editor), Word.isModified(editor)], [loaded, true]);
    Word.redo(editor);
    check('redone to the saved state: not modified', [Word.getRevision(editor), Word.isModified(editor)], [afterTyping, false]);
    // A different edit from the saved state's predecessor never reuses a revision.
    Word.undo(editor);
    Word.insertText(editor, 'different');
    const branch = Word.getRevision(editor);
    check('a new branch gets a revision that was never used', [branch > typedMore, Word.isModified(editor)], [true, true]);
    // The same text typed again is still not the saved state as far as revisions go, but undo is.
    Word.undo(editor);
    check('undo from the branch', [Word.getRevision(editor), Word.isModified(editor)], [loaded, true]);
    // Loading a document starts over, unmodified.
    Word.setParagraphs(editor, ['Another document.']);
    check('a newly loaded document is not modified and has a new revision', [Word.isModified(editor), Word.getRevision(editor) > branch, Word.canUndo(editor)], [false, true, false]);
} catch (e) { fail(e); }
finish();
