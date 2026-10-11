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

    // Half way down the document: the blocks above the view stay as they were laid out (not one
    // of them is laid out again before the change returns), the document is laid out from the
    // block at the top of the view; the layout from the start follows in slices, and when it has
    // come as far as the view the view is moved to stay on the text it shows (the host is told
    // where to scroll: a notice with ScrollTop).
    const plainLines = Word.getLayoutLines(editor);
    const firstLines = plainLines.filter((line, i) => line.BlockId.length > 0 && line.StartOffset === 0 && (i === 0 || plainLines[i - 1].BlockId !== line.BlockId));
    const topLine = firstLines[Math.floor(firstLines.length / 2)];
    const upperLine = firstLines.filter((line) => line.Y < topLine.Y - 3000).pop();
    const middle = topLine.Y - 2;
    Word.setViewport(editor, viewport(middle));
    const blockAtTop = () => { const hit = Word.hitTestViewport(editor, 300, 10); return hit ? hit.BlockId : null; };
    check('half way down: the block at the top of the view', blockAtTop(), topLine.BlockId);
    const scrolls = [];
    const offScroll = Word.onChange(editor, (change) => { if (typeof change.ScrollTop === 'number') { scrolls.push(change.ScrollTop); Word.setViewport(editor, viewport(change.ScrollTop)); } });
    t = performance.now();
    Word.redo(editor);
    const midTook = performance.now() - t;
    const midStats = Word.getRenderStats(editor);
    note(`in view half way down: returned in ${midTook.toFixed(1)} ms; ${midStats.StaleBlocks} blocks above the view kept as they were, ${midStats.LaidOutBlocks} laid out, the limit at ${midStats.LaidOutLimit} of ${midStats.Blocks}`);
    check('in view half way down: laid out as far as is in view, not to the end', [Word.isLayoutComplete(editor), midStats.LaidOutLimit > 1000, midStats.LaidOutLimit < 2600], [false, true, true]);
    check('the blocks above the view are not laid out again before the change returns', [midStats.StaleBlocks > 1000, midStats.LaidOutBlocks < 200], [true, true]);
    check('the view is on the block it was on', [blockAtTop(), Word.getSelectionFormat(editor).Bold], [topLine.BlockId, true]);
    // Undo and redo while the blocks above are stale: the same.
    t = performance.now();
    Word.undo(editor);
    const staleUndo = Word.getRenderStats(editor);
    Word.redo(editor);
    const staleRedo = Word.getRenderStats(editor);
    note(`undo and redo meanwhile: ${(performance.now() - t).toFixed(1)} ms for both`);
    check('undo and redo meanwhile keep the blocks above the view stale', [staleUndo.StaleBlocks > 1000, staleUndo.LaidOutBlocks < 200, staleRedo.StaleBlocks > 1000, staleRedo.LaidOutBlocks < 200, Word.isLayoutComplete(editor), blockAtTop()], [true, true, true, true, false, topLine.BlockId]);
    // Scrolled up into the stale blocks: they are laid out as far as the view reaches, the view stays on its block.
    Word.setViewport(editor, viewport(upperLine.Y - 2));
    const upStats = Word.getRenderStats(editor);
    check('scrolled up into the stale blocks: fewer are stale, the view is on the block it was scrolled to', [upStats.StaleBlocks > 0 && upStats.StaleBlocks < staleRedo.StaleBlocks, blockAtTop(), Word.isLayoutComplete(editor)], [true, upperLine.BlockId, false]);
    await new Promise((resolve) => setTimeout(resolve, 0));
    await Word.whenLayoutComplete(editor);
    await new Promise((resolve) => setTimeout(resolve, 0));
    offScroll();
    const endStats = Word.getRenderStats(editor);
    check('and to the end in slices', [Word.isLayoutComplete(editor), endStats.StaleBlocks, Word.getPageCount(editor)], [true, 0, boldPages]);
    check('the host was told where to scroll; the view is on the block it was on', [scrolls.length >= 1, scrolls[scrolls.length - 1] > upperLine.Y - 2, Word.getViewport(editor).scrollTop === scrolls[scrolls.length - 1], blockAtTop()], [true, true, true, upperLine.BlockId]);
    check('the layout at the end is that of the document laid out at once', JSON.stringify(linesOf(editor)) === JSON.stringify(linesOf(whole)), true);

    // A change of a few paragraphs is laid out at once, as ever.
    const lines = Word.getLayoutLines(editor);
    const some = [...new Set(lines.map((line) => line.BlockId))].slice(0, 40);
    Word.setSelection(editor, some[0], 0, some[39], 3);
    Word.toggleItalic(editor);
    check('forty paragraphs changed: laid out at once', Word.isLayoutComplete(editor), true);
} catch (e) { fail(e); }
finish();
