// The outline through the editor's API: getOutline / getOutlineLines / getCurrentHeading, promote /
// demote / a level for the selection, a heading moved with its subtext past its sibling, blocks
// moved, a heading deleted with its subtext, goToHeading - each one undo step - and the save read
// again with the same outline; getView / setView / setOutlineShowLevel and their ViewChanged notice.
import { check, finish, fail, loadWord, loadFormats } from './common.js';
try {
    const [Word, OF] = await Promise.all([loadWord(), loadFormats()]);
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.setParagraphs(editor, ['Intro', 'Alpha', 'alpha body', 'Beta', 'beta body', 'Gamma', 'gamma body']);
    const ids = () => Word.getOutlineLines(editor, false).map(line => line.BlockId);
    let blocks = ids();
    const caret = (i) => Word.setSelection(editor, blocks[i], 0, blocks[i], 0);
    const level = (i, n) => { caret(i); return Word.setOutlineLevel(editor, n); };
    check('a level chosen for body text makes a heading', [level(1, 1), level(3, 2), level(5, 1)], [true, true, true]);
    const outline = () => Word.getOutline(editor).map(h => [h.Text, h.Level, h.HasSubtext]);
    check('the outline: its headings, levels and subtext', outline(), [['Alpha', 1, true], ['Beta', 2, true], ['Gamma', 1, true]]);
    check('heading styles the document did not have are added', Word.getStyles(editor).filter(s => s.HeadingLevel > 0).map(s => s.Id).sort(), ['Heading1', 'Heading2']);
    const lines = Word.getOutlineLines(editor, true);
    check('outline lines: every block, level and depth', lines.map(l => [l.Text, l.Level, l.Depth]),
        [['Intro', 10, 0], ['Alpha', 1, 0], ['alpha body', 10, 1], ['Beta', 2, 1], ['beta body', 10, 2], ['Gamma', 1, 0], ['gamma body', 10, 1]]);
    check('outline lines: the runs as shown (the heading style\'s size)', lines[1].Runs.map(r => [r.Text, r.FontSize]), [['Alpha', 16]]);
    check('subtext: Alpha runs to Gamma', Word.getSubtextEnd(editor, blocks[1]), 5);
    caret(4);
    check('the heading the caret is under', Word.getCurrentHeading(editor), blocks[3]);
    caret(0);
    check('before the first heading: none', Word.getCurrentHeading(editor), null);

    const revision = Word.getRevision(editor);
    caret(3);
    check('promote: Heading 2 becomes Heading 1', [Word.outlinePromote(editor), outline()[1][1]], [true, 1]);
    check('one undo step', Word.canUndo(editor), true);
    Word.undo(editor);
    check('undone', [outline()[1][1], Word.getRevision(editor)], [2, revision]);
    caret(4);
    check('demote body text alone: one below the heading before', [Word.outlineDemote(editor), Word.getOutlineLines(editor, false)[4].Level], [true, 3]);
    Word.undo(editor);
    caret(1);
    check('Heading 1 promoted: no change, no step', [Word.outlinePromote(editor), Word.getRevision(editor)], [false, revision]);

    const texts = () => Word.getOutlineLines(editor, false).map(l => l.Text);
    check('a heading moved down with its subtext past its sibling', [Word.moveHeadingDown(editor, blocks[1]), texts()],
        [true, ['Intro', 'Gamma', 'gamma body', 'Alpha', 'alpha body', 'Beta', 'beta body']]);
    Word.undo(editor);
    check('undone in one step', texts(), ['Intro', 'Alpha', 'alpha body', 'Beta', 'beta body', 'Gamma', 'gamma body']);
    check('a first child has no previous sibling', Word.moveHeadingUp(editor, blocks[3]), false);
    check('blocks moved before another (a drag)', [Word.moveBlocks(editor, blocks[5], blocks[6], blocks[1]), texts()],
        [true, ['Intro', 'Gamma', 'gamma body', 'Alpha', 'alpha body', 'Beta', 'beta body']]);
    Word.undo(editor);
    check('blocks moved to the end', [Word.moveBlocks(editor, blocks[0], blocks[0], null), texts().at(-1)], [true, 'Intro']);
    Word.undo(editor);
    caret(2);
    check('Alt+Shift+Down semantics: the selection\'s block past the next', [Word.outlineMoveDown(editor), texts().slice(1, 4)], [true, ['Alpha', 'Beta', 'alpha body']]);
    Word.undo(editor);
    check('promote a branch: Alpha is Heading 1 already', Word.promoteHeading(editor, blocks[1]), false);
    check('demote a branch: the heading and its subheading', [Word.demoteHeading(editor, blocks[1]), outline().map(h => h[1])], [true, [2, 3, 1]]);
    Word.undo(editor);
    check('delete a heading and its subtext', [Word.deleteHeading(editor, blocks[3]), texts()], [true, ['Intro', 'Alpha', 'alpha body', 'Gamma', 'gamma body']]);
    check('the caret where it was', Word.getSelection(editor).FocusBlockId, blocks[5]);
    Word.undo(editor);
    check('go to a heading', [Word.goToHeading(editor, blocks[5]), Word.getSelection(editor).FocusBlockId, Word.getSelection(editor).FocusOffset], [true, blocks[5], 0]);
    check('select blocks', [Word.selectBlocks(editor, blocks[1], blocks[2]), Word.getSelection(editor).AnchorBlockId, Word.getSelection(editor).FocusBlockId, Word.getSelection(editor).FocusOffset], [true, blocks[1], blocks[2], 10]);
    caret(1); Word.outlineDemoteToBody(editor);
    check('demote to body text', Word.getOutlineLines(editor, false)[1].Level, 10);
    Word.outlinePromoteToHeading1(editor);
    check('promote to Heading 1', Word.getOutlineLines(editor, false)[1].Level, 1);

    // The save, read again.
    caret(3); Word.outlineDemote(editor);
    const before = Word.getOutlineLines(editor, false).map(l => [l.Text, l.Level]);
    const saved = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
    const reread = await OF.processors.forEditorPackage(new Uint8Array(await saved.arrayBuffer()));
    const other = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(other, reread);
    check('saved and read again: the same outline', Word.getOutlineLines(other, false).map(l => [l.Text, l.Level]), before);
    check('the page of each heading, once laid out', (Word.completeLayout(other), Word.getOutlinePages(other)), [1, 1, 1]);

    // The view.
    const notices = [];
    const stop = Word.onChange(editor, change => { if (change.ViewChanged) notices.push([change.View, change.OutlineShowLevel]); });
    check('print layout at first', Word.getView(editor), 3);
    Word.setView(editor, 2); Word.setOutlineShowLevel(editor, 2); Word.setOutlineShowLevel(editor, 12); Word.setView(editor, 2); Word.setView(editor, 3);
    check('setView / setOutlineShowLevel: notices, only when the view changes', notices, [[2, 0], [2, 2], [2, 0], [3, 0]]);
    stop();
} catch (e) { fail(e); }
finish();
