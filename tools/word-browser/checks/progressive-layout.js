// A long document loaded through a viewport shows its first screens at once
// and lays the rest out in slices (docx4j-getting-started with its body
// repeated 8 times: some 400 pages (since a page Word recorded as started is not started again - layout/page-parity - 49 pages repeated 8 times: 392 with these checks' fonts; Word has 45 pages of it)). The import returns with the first screen
// painted and a few pages laid out; getDocumentSize / getPageFrames then give
// the pages there are so far, the onChange listeners are told as they grow
// ({ LayoutChanged, LayoutComplete }), isLayoutComplete / whenLayoutComplete
// say when the whole document is laid out, and what needs all of it
// (getPageCount, a selection at the end, renderPage of a late page) lays the
// rest out at once. Typing while the rest is laid out is an ordinary edit.
// The whole, laid out in slices, is the document laid out at once: the same
// page frames, the same lines, the same pixels. The viewer does the same when
// asked to (load(package, { progressive: true })).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>'), sect = xml.lastIndexOf('<w:sectPr', close);
    zip.file('word/document.xml', xml.slice(0, open) + xml.slice(open, sect).repeat(8) + xml.slice(sect, close) + xml.slice(close));
    const editorPackage = await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' }));
    const viewport = (top) => ({ scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
    const editorOn = (id) => { const canvas = document.createElement('canvas'); canvas.id = id; document.body.appendChild(canvas); return [Word.initializeCanvas(canvas), canvas]; };
    const pixelsOf = (canvas) => canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    const same = (a, b) => { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
    const inkOf = (data) => { let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) ink++; return ink; };
    const linesOf = (editor) => JSON.stringify(Word.getLayoutLines(editor));

    // ---- the reference: the document laid out at once (no viewport when it is loaded) ----
    const [whole, wholeCanvas] = editorOn('whole');
    let t0 = performance.now();
    Word.importFromDocxPackage(whole, editorPackage);
    const atOnce = performance.now() - t0;
    check('without a viewport the whole document is laid out by the import', [Word.isLayoutComplete(whole), Word.getRenderStats(whole).LayoutComplete], [true, true]);
    const pageCount = Word.getPageCount(whole);
    const wholeLines = linesOf(whole);
    Word.setViewport(whole, viewport(0));
    // (Through a viewport, as the frames below: a narrower page is centred in the document's width.)
    const wholeFrames = JSON.stringify(Word.getPageFrames(whole));
    const wholeFirstScreen = pixelsOf(wholeCanvas);
    check('a long document', pageCount > 350, true);

    // ---- in slices ----
    const [editor, canvas] = editorOn('sliced');
    Word.setViewport(editor, viewport(0));
    const changes = [];
    const off = Word.onChange(editor, (change) => changes.push(change));
    t0 = performance.now();
    Word.importFromDocxPackage(editor, editorPackage);
    const first = performance.now() - t0;
    const firstSize = Word.getDocumentSize(editor);
    const firstStats = Word.getRenderStats(editor);
    check('the import returns before the whole document is laid out', [Word.isLayoutComplete(editor), firstSize.LayoutComplete, firstStats.LayoutComplete], [false, false, false]);
    check('a few pages are laid out, not the document', [firstSize.PageCount >= 2, firstSize.PageCount < 40, firstStats.LaidOutLimit < firstStats.Blocks / 10], [true, true, true]);
    check('getPageFrames gives the pages so far', Word.getPageFrames(editor).length, firstSize.PageCount);
    // (What is shown and a screen below it - with a viewport 800 high some two pages - not the forty blocks, twenty to eighty pages of some documents, it used to start with.)
    check(`the first paint lays out one or two screens (${firstSize.PageCount} pages, ${firstStats.LaidOutLimit} blocks)`, [firstSize.PageCount <= 3, firstSize.Height >= 1600], [true, true]);
    check('the first screen is painted, as in the whole layout', same(pixelsOf(canvas), wholeFirstScreen), true);
    check(`the first paint takes a fraction of the whole layout (${first.toFixed(0)} ms of ${atOnce.toFixed(0)} ms)`, first < Math.max(120, atOnce / 3), true);
    check('the load is told as a change of the document', [changes.length >= 1, changes[0].DocumentChanged, changes[0].LayoutChanged, changes[0].LayoutComplete], [true, true, false, false]);

    // The rest follows by itself; the page stays able to answer meanwhile.
    let longest = 0, last = performance.now();
    const ticker = setInterval(() => { const now = performance.now(); longest = Math.max(longest, now - last); last = now; }, 4);
    const sizes = [];
    const offSizes = Word.onChange(editor, (change) => { if (change.LayoutChanged) sizes.push(Word.getDocumentSize(editor).PageCount); });
    t0 = performance.now();
    const resolved = await Word.whenLayoutComplete(editor);
    const rest = performance.now() - t0;
    clearInterval(ticker); offSizes();
    const progress = changes.filter((change) => change.LayoutChanged);
    check('whenLayoutComplete resolves with the document size', [resolved.PageCount, resolved.LayoutComplete, Word.isLayoutComplete(editor)], [pageCount, true, true]);
    check('the listeners were told as the pages grew, and at the end', [progress.length >= 2, progress.every((change) => !change.DocumentChanged && !change.SelectionChanged), progress[progress.length - 1].LayoutComplete, progress.slice(0, -1).every((change) => !change.LayoutComplete)], [true, true, true, true]);
    check('the page count only grows, up to the whole document', [sizes.every((count, i) => i === 0 || count >= sizes[i - 1]), sizes[sizes.length - 1]], [true, pageCount]);
    check('laid out in slices, the page frames are those of the document laid out at once', JSON.stringify(Word.getPageFrames(editor)) === wholeFrames, true);
    check('and so is every line', linesOf(editor) === wholeLines, true);
    check('loading is not an edit', [Word.isModified(editor), Word.canUndo(editor)], [false, false]);
    check(`no slice held the page for long (longest ${longest.toFixed(0)} ms)`, longest < 400, true);
    check(`in slices the layout takes about what it takes at once (${(first + rest).toFixed(0)} ms, at once ${atOnce.toFixed(0)} ms)`, first + rest < atOnce * 2.5 + 400, true);
    note(`${pageCount} pages: first paint ${first.toFixed(0)} ms (${firstSize.PageCount} pages, ${firstStats.LaidOutLimit} of ${firstStats.Blocks} blocks), the rest ${rest.toFixed(0)} ms in ${Word.getRenderStats(editor).LayoutPasses} passes with ${progress.length} notices, longest slice ${longest.toFixed(0)} ms; at once ${atOnce.toFixed(0)} ms`);
    // The same pixels further down, too.
    const middle = JSON.parse(wholeFrames)[300].Top;
    Word.setViewport(editor, viewport(middle)); Word.setViewport(whole, viewport(middle));
    check('a screen in the middle is the same pixels', [inkOf(pixelsOf(canvas)) > 500, same(pixelsOf(canvas), pixelsOf(wholeCanvas))], [true, true]);
    off();
    Word.dispose(editor); canvas.remove();

    // ---- what needs the whole layout lays the rest out at once ----
    const fresh = () => { const [e, c] = editorOn('fresh'); Word.setViewport(e, viewport(0)); Word.importFromDocxPackage(e, editorPackage); return [e, c]; };
    {
        const [e, c] = fresh();
        check('getPageCount lays the rest out', [Word.isLayoutComplete(e), Word.getPageCount(e), Word.isLayoutComplete(e)], [false, pageCount, true]);
        Word.dispose(e); c.remove();
    }
    {
        const [e, c] = fresh();
        check('completeLayout does, and tells the listeners', (() => { const told = []; const stop = Word.onChange(e, (change) => told.push([change.LayoutChanged, change.LayoutComplete])); Word.completeLayout(e); stop(); return [Word.isLayoutComplete(e), Word.getDocumentSize(e).PageCount, JSON.stringify(told)]; })(), [true, pageCount, '[[true,true]]']);
        Word.dispose(e); c.remove();
    }
    {
        const [e, c] = fresh();
        const target = document.createElement('canvas');
        check('renderPage of a page that is laid out leaves the rest for later', [Word.renderPage(e, target, 1, 0.5), Word.isLayoutComplete(e)], [true, false]);
        check('renderPage of a late page lays the document out', [Word.renderPage(e, target, pageCount - 1, 0.5), Word.isLayoutComplete(e), Word.renderPage(e, target, pageCount, 0.5)], [true, true, false]);
        Word.dispose(e); c.remove();
    }
    {
        const [e, c] = fresh();
        check('the layout queries see the whole document', [linesOf(e) === wholeLines, Word.isLayoutComplete(e)], [true, true]);
        Word.dispose(e); c.remove();
    }
    {
        // A selection far down (go to, a find match, Ctrl+End): laid out as far as it reaches before it is shown.
        const [e, c] = fresh();
        const blocks = Word.getLayoutLines(whole).filter((line) => line.BlockId && !line.InTable);
        const far = blocks[Math.floor(blocks.length * 0.6)];
        Word.setSelection(e, far.BlockId, 0, far.BlockId, 0);
        const rect = Word.getCaretRect(e);
        Word.setSelection(whole, far.BlockId, 0, far.BlockId, 0);
        const wholeRect = Word.getCaretRect(whole);
        check('a caret put far down is where it is in the whole layout', [rect !== null && Math.round(rect.Y), rect !== null && Math.round(rect.X)], [Math.round(wholeRect.Y), Math.round(wholeRect.X)]);
        await Word.whenLayoutComplete(e);
        check('and the rest still follows', [Word.isLayoutComplete(e), Word.getDocumentSize(e).PageCount], [true, pageCount]);
        Word.dispose(e); c.remove();
    }

    // ---- typing while the rest is laid out (real keys, see below) ----
    const [typed, typedCanvas] = editorOn('typed');
    typedCanvas.style.position = 'absolute'; typedCanvas.style.left = '0px'; typedCanvas.style.top = '0px';
    Word.setViewport(typed, viewport(0));
    Word.importFromDocxPackage(typed, editorPackage);
    const typedState = { complete: Word.isLayoutComplete(typed) };
    window.progressive = {
        before: () => { typedState.atClick = Word.isLayoutComplete(typed); typedState.textBefore = Word.getText(typed).length; },
        after: () => { typedState.afterKeys = Word.isLayoutComplete(typed); typedState.sizeAfterKeys = Word.getDocumentSize(typed).PageCount; },
    };
    window.drive = [['eval', 'window.progressive.before()'], ['click', 500, 300], ['type', 'typed while laying out '], ['eval', 'window.progressive.after()']];
    window.finish = async () => {
        try {
            check('the keys were typed into the document while the rest was not laid out yet', [typedState.complete, typedState.atClick, Word.getText(typed).length - typedState.textBefore, Word.canUndo(typed)], [false, false, 'typed while laying out '.length, true]);
            const size = await Word.whenLayoutComplete(typed);
            // The same edit in the document laid out at once: the same lines.
            const selection = Word.getSelection(typed);
            Word.setSelection(whole, selection.FocusBlockId, selection.FocusOffset - 'typed while laying out '.length, selection.FocusBlockId, selection.FocusOffset - 'typed while laying out '.length);
            Word.insertText(whole, 'typed while laying out ');
            check('the document typed into while it was laid out is the document typed into afterwards', [size.LayoutComplete, linesOf(typed) === linesOf(whole), JSON.stringify(Word.getPageFrames(typed)) === JSON.stringify(Word.getPageFrames(whole))], [true, true, true]);
            Word.dispose(typed); typedCanvas.remove();

            // ---- the viewer ----
            const createViewer = (await import('/word/WordViewer.js')).default;
            const plain = createViewer(document.createElement('canvas'));
            plain.load(editorPackage);
            check('a viewer lays the whole document out unless asked otherwise', [plain.isLayoutComplete(), plain.getPageFrames().length, plain.getDocumentSize().LayoutComplete], [true, plain.getPageCount(), true]);
            const viewer = createViewer(document.createElement('canvas'));
            const told = [];
            viewer.load(editorPackage, { progressive: true });
            const stop = viewer.onLayoutProgress((progress) => told.push(progress));
            const t = performance.now();
            const early = viewer.getPageFrames();
            const target = document.createElement('canvas');
            const painted = viewer.renderPage(target, 0, 1);
            const firstPage = performance.now() - t;
            const firstPixels = pixelsOf(target);
            check('asked to, a viewer has its first pages before the rest', [viewer.isLayoutComplete(), early.length >= 1, early.length < 40, painted, inkOf(firstPixels) > 500], [false, true, true, true, true]);
            const frames = await viewer.whenLayoutComplete();
            stop();
            check('whenLayoutComplete resolves with all the page frames', [frames.length, viewer.isLayoutComplete(), JSON.stringify(frames) === JSON.stringify(plain.getPageFrames())], [plain.getPageCount(), true, true]);
            check('onLayoutProgress told of the pages as they were added', [told.length >= 2, told[told.length - 1].LayoutComplete, told[told.length - 1].PageCount, told.every((progress, i) => i === 0 || progress.PageCount >= told[i - 1].PageCount)], [true, true, plain.getPageCount(), true]);
            const again = document.createElement('canvas');
            plain.renderPage(again, 0, 1);
            viewer.renderPage(target, 0, 1);
            check('the first page is the same pixels before and after, and in a viewer that laid out at once', [same(pixelsOf(target), firstPixels), same(pixelsOf(again), firstPixels)], [true, true]);
            plain.renderPage(again, 350, 1); viewer.renderPage(target, 350, 1);
            check('and so is a late page', same(pixelsOf(again), pixelsOf(target)), true);
            const late = createViewer(document.createElement('canvas'));
            late.load(editorPackage, { progressive: true });
            check('renderPage of a late page lays a progressive viewer out at once', [late.renderPage(target, 350, 1), late.isLayoutComplete(), same(pixelsOf(again), pixelsOf(target))], [true, true, true]);
            note(`the viewer: first page ${firstPage.toFixed(0)} ms with ${early.length} pages laid out, ${told.length} progress notices`);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
