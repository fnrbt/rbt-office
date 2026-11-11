// A document revision that undo and redo restore, and "modified since saved":
// undoing back to the saved state is not modified.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
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
    // ---- a document that is what was opened is saved as the file that was opened ----
    {
        const OF = await loadFormats();
        const original = await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx');
        const same = (a, b) => { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
        const written = async (pkg) => new Uint8Array(await (await OF.word.writeFromEditorPackage(pkg)).arrayBuffer());
        const pkg = await OF.processors.forEditorPackage(original);
        Word.importFromDocxPackage(editor, pkg);
        const opened = Word.getRevision(editor);
        const unedited = Word.exportDocumentPackage(editor);
        const first = await written(unedited);
        check('an unedited document says so, and is saved as the file that was opened, byte for byte', [unedited.Unchanged, same(first, original)], [true, true]);
        // (A copy each time: what the host does with the bytes does not reach the next save.)
        first.fill(0);
        check('the bytes are a copy', [same(await written(Word.exportDocumentPackage(editor)), original)], [true]);
        Word.markSaved(editor);
        check('saving it changes nothing of that', [Word.exportDocumentPackage(editor).Unchanged], [true]);
        Word.insertText(editor, 'typed');
        const edited = Word.exportDocumentPackage(editor);
        const editedBytes = await written(edited);
        const editedXml = await (await window.JSZip.loadAsync(editedBytes)).file('word/document.xml').async('string');
        check('an edited document is written by the writer', [edited.Unchanged, same(editedBytes, original), editedXml.includes('typed')], [false, false, true]);
        Word.markSaved(editor);
        check('a saved edit is still a change of what was opened', [Word.isModified(editor), Word.exportDocumentPackage(editor).Unchanged], [false, false]);
        Word.undo(editor);
        const undone = Word.exportDocumentPackage(editor);
        check('undone back to what was opened: the file that was opened again', [Word.getRevision(editor), undone.Unchanged, same(await written(undone), original)], [opened, true, true]);
        Word.redo(editor);
        check('redone: written by the writer', [Word.exportDocumentPackage(editor).Unchanged], [false]);
        Word.undo(editor);
        // Long coalesced runs must retain both the opened state and a saved
        // checkpoint, even though hundreds of intermediate states were seen.
        Word.markSaved(editor);
        const proxy = Word.getInputElement(editor);
        const typeRun = (key) => {
            for (let i = 0; i < 410; i++)
                proxy.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
        };
        typeRun('a');
        const checkpoint = Word.getRevision(editor);
        check('undo of a long coalesced run returns to the saved opened state', [Word.undo(editor), Word.getRevision(editor), Word.isModified(editor)], [true, opened, false]);
        check('redo of the long coalesced run restores its revision', [Word.redo(editor), Word.getRevision(editor), Word.isModified(editor)], [true, checkpoint, true]);
        const checkpointText = Word.getText(editor);
        Word.exportDocumentPackage(editor); // ends the first typing run
        typeRun('b');
        const afterLongRuns = Word.getRevision(editor);
        Word.markSaved(editor, checkpoint);
        check('a delayed save after long typing keeps later edits modified', [Word.isModified(editor), afterLongRuns > checkpoint], [true, true]);
        check('undo restores the long-run saved checkpoint', [Word.undo(editor), Word.getRevision(editor), Word.isModified(editor), Word.getText(editor) === checkpointText], [true, checkpoint, false, true]);
        check('undo restores the opened revision after more than 400 observations', [Word.undo(editor), Word.getRevision(editor), Word.isModified(editor)], [true, opened, true]);
        const longUndone = Word.exportDocumentPackage(editor);
        const longUndoneBytes = await written(longUndone);
        check('the restored opened revision exports the original bytes', [longUndone.Unchanged, same(longUndoneBytes, original)], [true, true]);
        const reopened = Word.initializeCanvas(document.createElement('canvas'));
        try {
            Word.importFromDocxPackage(reopened, await OF.processors.forEditorPackage(longUndoneBytes));
            const reopenedPackage = Word.exportDocumentPackage(reopened);
            check('the restored output reopens unchanged and exports byte-identically', [Word.isModified(reopened), reopenedPackage.Unchanged, same(await written(reopenedPackage), original)], [false, true, true]);
        } finally {
            Word.dispose(reopened);
        }
        check('redo returns to the saved checkpoint', [Word.redo(editor), Word.getRevision(editor), Word.isModified(editor)], [true, checkpoint, false]);
        check('redo returns to the later long run', [Word.redo(editor), Word.getRevision(editor), Word.isModified(editor)], [true, afterLongRuns, true]);
        Word.undo(editor);
        Word.undo(editor);
        // A host that changes the package it exported says so; a package that was made from the model alone is written.
        const changed = Word.exportDocumentPackage(editor);
        changed.Unchanged = false;
        const rewritten = await written(changed);
        check('a package that does not say it is unchanged is written by the writer', [same(rewritten, original), (await window.JSZip.loadAsync(rewritten)).file('word/document.xml') !== null], [false, true]);
        const regenerated = Word.exportDocumentPackage(editor);
        regenerated.RegenerateAll = true;
        check('so is one that is to be written from the model', [same(await written(regenerated), original)], [false]);
        // A document that was not opened from a file has none to give back.
        Word.setParagraphs(editor, ['A new document.']);
        const fresh = Word.exportDocumentPackage(editor);
        check('a new document is written', [fresh.Unchanged, (await written(fresh)).length > 500], [true, true]);
    }
} catch (e) { fail(e); }
finish();
