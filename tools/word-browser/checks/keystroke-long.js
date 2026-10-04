// Typing in a long document (docx4j-getting-started with its body repeated 8
// times: some 600 pages, with tables, sections and paragraphs that break over
// pages). What an edit did not touch is kept from the last pass - table plans,
// the rest of paragraphs broken over a page, the paint commands of
// paragraphs - so a keystroke costs a fraction of what it did (330 ms in
// 0.1.5 on the machine these checks were written on, 125 ms now); the time is
// reported, and bounded loosely so that a return to rebuilding everything
// shows. tools/word-browser/perf.mjs profiles the same keystroke.
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
        const result = { pages: frames.length, median: times_[10], max: times_[19], typed: Word.getText(editor).length - before, builds: Word.getRenderStats(editor).FrameBuilds - builds, placed };
        Word.dispose(editor); canvas.remove();
        return result;
    };
    const short = await measure(1), long = await measure(8);
    note(`a keystroke: ${short.pages} pages median ${short.median.toFixed(1)} ms (max ${short.max.toFixed(1)}); ${long.pages} pages median ${long.median.toFixed(1)} ms (max ${long.max.toFixed(1)})`);
    check('the caret was placed in body text and the text was typed', [short.placed, long.placed, short.typed, long.typed], [true, true, 23, 23]);
    check('one frame per keystroke', [short.builds, long.builds], [20, 20]);
    check('the long document is long', long.pages > 400, true);
    check('a keystroke at some 600 pages stays under 220 ms (0.1.5: 330 ms)', long.median < 220, true);
    check('and under 40 ms at 75 pages', short.median < 40, true);
} catch (e) { fail(e); }
finish();
