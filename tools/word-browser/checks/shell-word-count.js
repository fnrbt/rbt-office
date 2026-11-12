// Live body and selection counts, including history, stories and shell lifecycle.
import { check, finish, fail, loadWord, loadFormats } from './common.js';
let dispose, disposeOther;
try {
    const [Word, OF, Host, Shell, Dom] = await Promise.all([
        loadWord(), loadFormats(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'count-shell'; root.style.width = '390px'; document.body.appendChild(root);
    const otherRoot = document.createElement('div'); otherRoot.id = 'count-other'; document.body.appendChild(otherRoot);
    const host = Host.create('', 'word-shell');
    dispose = Dom.mountApp('#count-shell', Shell.view(host));
    const other = Host.create('independent editor', 'word-shell');
    disposeOther = Dom.mountApp('#count-other', Shell.view(other));
    const editor = host.Editor;
    const settle = () => new Promise(resolve => requestAnimationFrame(resolve));
    const count = () => root.querySelector('[data-testid="word-count"]').textContent.match(/\d+/g).map(Number);
    const otherCount = () => otherRoot.querySelector('[data-testid="word-count"]').textContent.match(/\d+/g).map(Number);
    Word.setParagraphs(editor, ['alpha\u00a0beta gamma', 'delta epsilon']);
    const ids = [...new Set(Word.getLayoutLines(editor).map(line => line.BlockId))];
    await settle();
    check('body includes Unicode whitespace and separate paragraphs', count(), [5]);
    const revision = Word.getRevision(editor);
    Word.setSelection(editor, ids[1], 5, ids[0], 2);
    await settle();
    host.Status.Value = 'Saved message'; await settle();
    check('reverse selection counts paragraph fragments', count(), [5, 4]);
    check('selection fragments agree with range statistics', [Word.computeRangeStatistics(editor, ids[0], 2, 16).Words, Word.computeRangeStatistics(editor, ids[1], 0, 5).Words], [3, 1]);
    check('count leaves action status intact', root.querySelector('[data-testid=\"word-status\"]').textContent, 'Saved message');
    check('count query does not change revision or text', [Word.getRevision(editor), Word.getText(editor)], [revision, 'alpha\u00a0beta gamma\ndelta epsilon']);
    Word.setSelection(editor, ids[1], 13, ids[1], 13);
    Word.insertText(editor, ' zeta'); await settle();
    check('edit refreshes count', count(), [6]);
    Word.undo(editor); await settle(); check('undo refreshes count', count(), [5]);
    Word.redo(editor); await settle(); check('redo refreshes count', count(), [6]);
    const blob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
    const pkg = await OF.processors.forEditorPackage(new Uint8Array(await blob.arrayBuffer()));
    Word.setContent(editor, 'replacement'); await settle();
    check('replacement updates count', count(), [1]);
    Host.importDocumentPackage(host, pkg, 'counts.docx'); await settle();
    check('import restores count', count(), [6]);
    Word.setContent(editor, 'before'); Word.insertTable(editor, 1, 2);
    Word.insertText(editor, 'one two');
    const first = Word.getSelection(editor).FocusBlockId;
    Word.moveToAdjacentTableCell(editor, true); Word.insertText(editor, 'three four');
    const second = Word.getSelection(editor).FocusBlockId;
    Word.setSelection(editor, second, 5, first, 4); await settle();
    check('selection spans table cells in reverse', [Word.getWordCounts(editor).SelectedWords, count()[1]], [2, 2]);
    const beforeBox = Word.getWordCounts(editor).BodyWords;
    Word.setSelection(editor, second, 10, second, 10);
    Word.insertTextBox(editor, 'box text'); await settle();
    check('body includes text box text once', Word.getWordCounts(editor).BodyWords, beforeBox + 2);
    Word.getInputElement(editor).readOnly = true;
    Word.setSelection(editor, first, 0, first, 3); await settle();
    check('read-only selection still reports counts', count()[1], 1);
    check('another mounted editor remains independent', otherCount(), [2]);
    check('status fits narrow shell', root.querySelector('[data-testid="word-count"]').getBoundingClientRect().right <= root.getBoundingClientRect().right, true);
    dispose(); dispose = null;
    const remounted = Host.create('fresh shell', 'word-shell');
    dispose = Dom.mountApp('#count-shell', Shell.view(remounted)); await settle();
    check('remount has its own count', count(), [2]);
    check('other editor survives disposal', otherCount(), [2]);
} catch (e) { fail(e); }
finally { if (dispose) dispose(); if (disposeOther) disposeOther(); finish(); }
