// Most of a long document changed at once - everything selected and made
// bold - through a viewport: the change returns with what is in view laid
// out (not the whole document: isLayoutComplete is false, the pages are those
// there are so far), the rest follows in slices, the listeners are told, and
// the layout at the end is the one a document that was bold from the start
// has. Undo is the same. Typing meanwhile works. What needs the whole layout
// (getPageCount) lays the rest out at once. Without a viewport, and for a
// change of a few paragraphs, nothing is put off.
import { check, note, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const paragraphs = Array.from({ length: 3000 }, (_, i) => `Paragraph ${i + 1}: ` + Array.from({ length: 12 + (i % 9) * 4 }, (_, w) => `word${(i * 7 + w) % 97}`).join(' '));
    const viewport = (scrollTop) => ({ scale: 1, devicePixelRatio: 1, scrollTop, scrollLeft: 0, width: 1000, height: 800 });
    const make = (withViewport) => {
        const canvas = document.createElement('canvas'); canvas.width = 1000; canvas.height = 800;
        const editor = Word.initializeCanvas(canvas);
        if (withViewport) Word.setViewport(editor, viewport(0));
        Word.setParagraphs(editor, paragraphs);
        return editor;
    };
    const linesOf = (editor) => Word.getLayoutLines(editor).map((line) => [line.BlockId.length > 0, Math.round(line.Y), line.StartOffset, line.EndOffset]);

    const editor = make(true);
    await Word.whenLayoutComplete(editor);
    const plainPages = Word.getPageCount(editor);
    const stats0 = Word.getRenderStats(editor);
    check('laid out to its end before the change', [Word.isLayoutComplete(editor), stats0.Blocks], [true, 3000]);
    const firstId = Word.getLayoutLines(editor)[0].BlockId;

    // Everything made bold.
    const changes = [];
    const off = Word.onChange(editor, (change) => changes.push(change));
    Word.selectAll(editor);
    let t = performance.now();
    Word.toggleBold(editor);
    const took = performance.now() - t;
    const after = Word.getRenderStats(editor);
    const size = Word.getDocumentSize(editor);
    note(`3000 paragraphs made bold: returned in ${took.toFixed(1)} ms with ${after.LaidOutLimit} of ${after.Blocks} blocks laid out (${size.PageCount} of ${plainPages} pages so far)`);
    check('the change returns before the document is laid out again', [Word.isLayoutComplete(editor), after.LayoutComplete, after.LaidOutLimit < 400, size.PageCount < plainPages / 4], [false, false, true, true]);
    check('what is in view is laid out, in bold', [after.LaidOutLimit >= 6, size.PageCount >= 1, Word.getSelectionFormat(editor).Bold], [true, true, true]);
    // Typing meanwhile (the selection replaced by a character would empty the document: the caret is put at the start first).
    // (The layout queries - getLayoutLines - lay the rest out at once: they are not asked here.)
    Word.setSelection(editor, firstId, 0, firstId, 0);
    Word.insertText(editor, 'Z');
    check('typing meanwhile', [Word.getText(editor).startsWith('ZParagraph 1:'), Word.isLayoutComplete(editor)], [true, false]);
    Word.undo(editor);
    check('and undoing it', [Word.getText(editor).startsWith('Paragraph 1:'), Word.isLayoutComplete(editor)], [true, false]);
    const resolved = await Word.whenLayoutComplete(editor);
    off();
    const told = changes.filter((change) => change.LayoutChanged);
    check('the rest follows in slices: the listeners are told, at the end that the layout is complete', [told.length >= 1, told[told.length - 1].LayoutComplete, Word.isLayoutComplete(editor), resolved.LayoutComplete], [true, true, true, true]);
    const boldPages = Word.getPageCount(editor);
    // The same document made bold at once (no viewport: nothing is put off).
    const whole = make(false);
    Word.selectAll(whole);
    t = performance.now();
    Word.toggleBold(whole);
    const wholeTook = performance.now() - t;
    check('without a viewport the change is laid out at once', Word.isLayoutComplete(whole), true);
    note(`without a viewport the same change takes ${wholeTook.toFixed(1)} ms`);
    check('the layout at the end is that of the document laid out at once', [boldPages, JSON.stringify(linesOf(editor)) === JSON.stringify(linesOf(whole))], [Word.getPageCount(whole), true]);
    check('bold text takes more room', boldPages >= plainPages, true);

    // Undo: put off the same way; what needs the whole layout lays the rest out at once.
    Word.undo(editor);
    check('undo returns before the document is laid out again', [Word.isLayoutComplete(editor), Word.getRenderStats(editor).LaidOutLimit < 400], [false, true]);
    check('getPageCount lays the rest out at once', [Word.getPageCount(editor), Word.isLayoutComplete(editor)], [plainPages, true]);

    // Half way down the document: what is above the view is laid out first, the rest put off.
    const frames = Word.getPageFrames(editor);
    const middle = frames[Math.floor(frames.length / 2)].Top ?? frames[Math.floor(frames.length / 2)].Y;
    Word.setViewport(editor, viewport(middle));
    Word.redo(editor);
    const midStats = Word.getRenderStats(editor);
    note(`in view half way down: ${midStats.LaidOutLimit} of ${midStats.Blocks} blocks laid out when the change returns`);
    check('in view half way down: laid out as far as is in view, not to the end', [Word.isLayoutComplete(editor), midStats.LaidOutLimit > 1000, midStats.LaidOutLimit < 2600], [false, true, true]);
    await Word.whenLayoutComplete(editor);
    check('and to the end in slices', [Word.isLayoutComplete(editor), Word.getPageCount(editor)], [true, boldPages]);

    // A change of a few paragraphs is laid out at once, as ever.
    const lines = Word.getLayoutLines(editor);
    const some = [...new Set(lines.map((line) => line.BlockId))].slice(0, 40);
    Word.setSelection(editor, some[0], 0, some[39], 3);
    Word.toggleItalic(editor);
    check('forty paragraphs changed: laid out at once', Word.isLayoutComplete(editor), true);
} catch (e) { fail(e); }
finish();
