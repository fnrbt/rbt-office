// Typing in a long document (docx4j-getting-started with its body repeated 8
// times: some 400 pages - 392 with these checks' fonts since layout/page-parity -, with tables, sections and paragraphs that break over
// pages). An edit lays out and paints what it changed: the layout pass starts
// at the edited block from the state the last pass had there and stops where
// its state is the last pass's again, and the paint commands (with their
// index) are kept block by block, the headers and footers page by page. A
// keystroke cost 330 ms in 0.1.5 on the machine these checks were written
// on, 125 ms in 0.1.6 and some 7 ms now; the time is reported, and bounded
// loosely so that a return to walking the whole document shows - as do the
// counts of blocks laid out and collected (getRenderStats). What is on the
// canvas after a run of edits is what a layout from nothing paints.
// tools/word-browser/perf.mjs profiles the same keystroke.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>'), sect = xml.lastIndexOf('<w:sectPr', close);
    const build = async (times) => {
        const copy = await window.JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
        copy.file('word/document.xml', xml.slice(0, open) + xml.slice(open, sect).repeat(times) + xml.slice(sect, close) + xml.slice(close));
        return OF.processors.forEditorPackage(await copy.generateAsync({ type: 'uint8array' }));
    };
    const measure = async (times) => {
        const canvas = document.createElement('canvas');
        document.body.appendChild(canvas);
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await build(times));
        const frames = Word.getPageFrames(editor), target = frames[Math.floor(frames.length / 2)];
        Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: target.Top, scrollLeft: 0, width: 1000, height: 800 });
        // The caret into a paragraph of text on that page: the first block there that takes typed text on one line.
        const box = canvas.getBoundingClientRect();
        let placed = false;
        for (let y = 120; y < 760 && !placed; y += 40) {
            canvas.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: box.left + 400, clientY: box.top + y, button: 0 }));
            canvas.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: box.left + 400, clientY: box.top + y, button: 0 }));
            const information = Word.getSelectionInformation(editor);
            placed = !!information && !information.wdWithInTable;
        }
        const before = Word.getText(editor).length;
        for (let i = 0; i < 3; i++) Word.insertText(editor, 'x');
        const builds = Word.getRenderStats(editor).FrameBuilds, times_ = [];
        for (let i = 0; i < 20; i++) { const t0 = performance.now(); Word.insertText(editor, 'x'); times_.push(performance.now() - t0); }
        times_.sort((a, b) => a - b);
        const stats = Word.getRenderStats(editor);
        const result = { pages: frames.length, median: times_[10], max: times_[19], typed: Word.getText(editor).length - before, builds: stats.FrameBuilds - builds, placed,
            laidOut: stats.LaidOutBlocks, collected: stats.CollectedBlocks, blocks: stats.Blocks };
        // Enter, typing in the new paragraph, Backspace back over it and the paragraph mark, a longer
        // insert: then the canvas and the lines are those of a layout and a paint from nothing.
        const input = Word.getInputElement(editor);
        const press = (key) => input.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
        const enterTimes = [], mergeTimes = [];
        for (let i = 0; i < 6; i++) {
            let t0 = performance.now(); press('Enter'); enterTimes.push(performance.now() - t0);
            result.enterLaidOut = Word.getRenderStats(editor).LaidOutBlocks;
            Word.insertText(editor, 'n');
            press('Backspace');
            t0 = performance.now(); press('Backspace'); mergeTimes.push(performance.now() - t0);
        }
        // Pages of text pasted in: the section gains pages, and every section after it stands that much further down.
        const pagesBefore = Word.getDocumentSize(editor).PageCount;
        const pasted = 'pasted words that fill pages '.repeat(400);
        let t1 = performance.now(); Word.insertText(editor, pasted); result.pageBreak = performance.now() - t1;
        result.pageBreakLaidOut = Word.getRenderStats(editor).LaidOutBlocks;
        result.pagesAdded = Word.getDocumentSize(editor).PageCount - pagesBefore;
        t1 = performance.now(); Word.undo(editor); result.pageBreakUndone = performance.now() - t1;
        result.pagesAfterUndo = Word.getDocumentSize(editor).PageCount - pagesBefore;
        Word.insertText(editor, pasted);    // (stays: what is compared below is the document with its later sections moved)
        press('Enter'); Word.insertText(editor, 'a new paragraph that stays, with enough words in it to run over the end of its first line and on to a second one');
        enterTimes.sort((a, b) => a - b); mergeTimes.sort((a, b) => a - b);
        result.enter = enterTimes[3]; result.merge = mergeTimes[3];
        const context = canvas.getContext('2d');
        const pixels = () => context.getImageData(0, 0, canvas.width, canvas.height).data;
        const kept = pixels(), keptLines = JSON.stringify(Word.getLayoutLines(editor)), keptPages = Word.getPageCount(editor);
        Word.clearTextMeasurer(editor);     // drops every layout and paint command that was kept
        Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: target.Top, scrollLeft: 0, width: 1000, height: 800 });
        const fresh = pixels();
        let differing = 0;
        for (let i = 0; i < kept.length; i += 4) if (kept[i] !== fresh[i] || kept[i + 1] !== fresh[i + 1] || kept[i + 2] !== fresh[i + 2]) differing++;
        result.differing = kept.length === fresh.length ? differing : -1;
        result.sameLines = keptLines === JSON.stringify(Word.getLayoutLines(editor)) && keptPages === Word.getPageCount(editor);
        Word.dispose(editor); canvas.remove();
        return result;
    };
    const short = await measure(1), long = await measure(8);
    note(`a keystroke: ${short.pages} pages median ${short.median.toFixed(1)} ms (max ${short.max.toFixed(1)}); ${long.pages} pages median ${long.median.toFixed(1)} ms (max ${long.max.toFixed(1)})`);
    check('the caret was placed in body text and the text was typed', [short.placed, long.placed, short.typed, long.typed], [true, true, 23, 23]);
    check('one frame per keystroke', [short.builds, long.builds], [20, 20]);
    note(`Enter: ${short.enter.toFixed(1)} ms / ${long.enter.toFixed(1)} ms; Backspace joining two paragraphs: ${short.merge.toFixed(1)} ms / ${long.merge.toFixed(1)} ms`);
    note(`the last keystroke laid out ${long.laidOut} and collected the commands of ${long.collected} of ${long.blocks} blocks`);
    check('the long document is long', long.pages > 350, true);
    check('a keystroke at some 600 pages stays under 45 ms (0.1.6: 125 ms, 0.1.5: 330 ms)', long.median < 45, true);
    check('and under 25 ms at 75 pages', short.median < 25, true);
    check('a keystroke lays out and collects a handful of blocks, not the document', [long.laidOut > 0 && long.laidOut < 60, long.collected > 0 && long.collected < 60, long.blocks > 5000], [true, true, true]);
    // Enter here adds a line to a section of some 500 blocks: that section is laid out from the new
    // paragraph on; the sections after it (half the document) are moved, not laid out again
    // (85 ms before; the sections are flowed from their own first page, see Layout.sectionLayoutsAt).
    note(`Enter laid out ${long.enterLaidOut} of ${long.blocks} blocks; pages of text pasted in ${long.pageBreak.toFixed(1)} ms (${long.pagesAdded} page(s) more, ${long.pageBreakLaidOut} blocks laid out), undone in ${long.pageBreakUndone.toFixed(1)} ms`);
    check('Enter in the middle of some 600 pages stays under 60 ms', long.enter < 60, true);
    check('text pasted there adds pages, which undo takes away again', [long.pagesAdded >= 1, long.pagesAfterUndo], [true, 0]);
    check('that lays out the rest of its own section, not the sections after it (half the document)', [long.pageBreakLaidOut > 0, long.pageBreakLaidOut < long.blocks / 6], [true, true]);
    check('and takes under 60 ms, as does taking it back', [long.pageBreak < 60, long.pageBreakUndone < 60], [true, true]);
    check('after Enter, Backspace and typing the lines and pages are those of a layout from nothing', [short.sameLines, long.sameLines], [true, true]);
    check('and so are the pixels', [short.differing, long.differing], [0, 0]);
} catch (e) { fail(e); }
finish();
