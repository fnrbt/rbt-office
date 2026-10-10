// A DOCX opened front first (officeFormats.word.openProgressive): the package
// is handed over when the parts around the body and the first blocks are read,
// the rest is read in slices; read to its end it is, field for field, the
// package forEditorPackage makes - for every test document, whatever the
// size of the steps.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    // (A package as text: bytes by their length and a sum, the reading state left out.)
    const textOf = (pkg) => JSON.stringify(pkg, (key, value) => {
        if (key === 'Progressive') return undefined;
        if (value instanceof Uint8Array) { let sum = 0; for (let i = 0; i < value.length; i++) sum = (sum * 31 + value[i]) >>> 0; return `bytes:${value.length}:${sum}`; }
        return value;
    });
    const firstDifference = (a, b) => { let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++; return `at ${i}: ...${a.slice(Math.max(0, i - 60), i + 80)}... against ...${b.slice(Math.max(0, i - 60), i + 80)}...`; };
    const corpus = await (await fetch('/corpus-list')).json();
    let same = 0, frontFirst = 0, differing = [];
    for (const name of corpus) {
        const bytes = await bytesOf('/repo/' + name);
        const whole = textOf(await OF.processors.forEditorPackage(bytes));
        for (const [label, read] of [
            ['in steps', async (pkg) => { while (!pkg.Progressive.IsComplete) pkg.Progressive.readMore(0); }],
            ['at once', async (pkg) => { pkg.Progressive.complete(); }],
            ['in slices', async (pkg) => { pkg.Progressive.start(); await OF.word.whenLoadComplete(pkg); }],
        ]) {
            const pkg = await OF.word.openProgressive(bytes, { autoStart: false, firstBlocks: label === 'in steps' ? 3 : 64 });
            if (label === 'in steps' && !pkg.Progressive.AtOnce) frontFirst++;
            // What is there at first is the start of what is there at the end.
            const firstItems = pkg.Paragraphs.length, firstText = JSON.stringify(pkg.Paragraphs);
            await read(pkg);
            const text = textOf(pkg);
            if (text === whole && OF.word.isLoadComplete(pkg) && JSON.stringify(pkg.Paragraphs.slice(0, firstItems)) === firstText) same++;
            else differing.push(`${name} (${label}): ${text === whole ? 'the first items changed, or the package is not complete' : firstDifference(whole, text)}`);
        }
    }
    check(`the package read front first is the package read at once (${corpus.length} documents, three ways each)`, [same, differing.slice(0, 5)], [corpus.length * 3, []]);
    check(`most documents are read front first (${frontFirst} of ${corpus.length})`, frontFirst >= corpus.length * 0.7, true);

    // ---- a long document in an editor: shown before it is read -------------------
    // docx4j-getting-started with its body repeated 8 times (some 600 pages).
    const zip = await window.JSZip.loadAsync(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>'), sect = xml.lastIndexOf('<w:sectPr', close);
    zip.file('word/document.xml', xml.slice(0, open) + xml.slice(open, sect).repeat(8) + xml.slice(sect, close) + xml.slice(close));
    const long = await zip.generateAsync({ type: 'uint8array' });
    const viewport = (top) => ({ scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
    const editorOn = (id) => { const canvas = document.createElement('canvas'); canvas.id = id; document.body.appendChild(canvas); const editor = Word.initializeCanvas(canvas); Word.setViewport(editor, viewport(0)); return [editor, canvas]; };
    const pixelsOf = (canvas) => canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    const samePixels = (a, b) => { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
    const inkOf = (data) => { let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) ink++; return ink; };

    // The reference: the document read at once.
    let t0 = performance.now();
    const wholePackage = await OF.processors.forEditorPackage(long);
    const readAtOnce = performance.now() - t0;
    const [whole, wholeCanvas] = editorOn('whole');
    Word.importFromDocxPackage(whole, wholePackage);
    const wholeFirstScreen = pixelsOf(wholeCanvas);
    const wholeText = Word.getText(whole);
    const pageCount = Word.getPageCount(whole);
    const wholeFrames = JSON.stringify(Word.getPageFrames(whole));
    const linesOf = (e) => JSON.stringify(Word.getLayoutLines(e).map((line) => [line.BlockIndex, line.LineIndex, line.Page, Math.round(line.Y), Math.round(line.Height)]));
    const wholeLines = linesOf(whole);
    const wholeFonts = Word.getRequiredFonts(wholePackage).map((face) => `${face.family}|${face.bold}|${face.italic}`).sort();
    check('a long document', [pageCount > 500, wholeText.length > 500000], [true, true]);

    // Opened front first: the package comes with the first blocks, the editor shows them.
    t0 = performance.now();
    const pkg = await OF.word.openProgressive(long, { autoStart: false });
    const opened = performance.now() - t0;
    check('the package is handed over with the start of the body', [pkg.Progressive.IsComplete, pkg.Progressive.AtOnce, pkg.Paragraphs.length >= 64, pkg.Paragraphs.length < 100, OF.word.isLoadComplete(pkg)], [false, false, true, true, false]);
    // (The document's last section is on A4, its first on Letter: the package states the page setup
    // of the section its blocks end in, so the first pages are on their paper before the paragraph
    // that ends the first section is read.)
    check('the page setup is that of the section being read', [Math.round(pkg.PageSetup.PageWidth), Math.round(wholePackage.PageSetup.PageWidth)], [816, 794]);
    check(`it is opened in a fraction of the time the whole takes to read (${opened.toFixed(0)} ms of ${readAtOnce.toFixed(0)} ms)`, opened < readAtOnce / 2, true);
    const earlyFonts = Word.getRequiredFonts(pkg).map((face) => `${face.family}|${face.bold}|${face.italic}`);
    const discovered = [];
    const stopFonts = Word.onFontsDiscovered(pkg, (faces) => { for (const face of faces) discovered.push(`${face.family}|${face.bold}|${face.italic}`); });
    check('the faces of the styles and the first blocks are known at once', [earlyFonts.length > 0, earlyFonts.every((face) => wholeFonts.includes(face))], [true, true]);
    const [editor, canvas] = editorOn('front');
    const changes = [];
    const off = Word.onChange(editor, (change) => changes.push({ ...change }));
    t0 = performance.now();
    Word.importFromDocxPackage(editor, pkg);
    const imported = performance.now() - t0;
    check('the import returns with the document still being read', [Word.isLoadComplete(editor), Word.isLayoutComplete(editor), Word.getDocumentSize(editor).LayoutComplete], [false, false, false]);
    check('the first screen is painted, as in the document read at once', [inkOf(pixelsOf(canvas)) > 2000, samePixels(pixelsOf(canvas), wholeFirstScreen)], [true, true]);
    check('the load is told as a change of the document, not complete', [changes.length >= 1, changes[0].DocumentChanged, changes[0].LoadComplete, changes[0].LayoutComplete], [true, true, false, false]);
    check('loading is not an edit', [Word.isModified(editor), Word.canUndo(editor)], [false, false]);
    note(`opened in ${opened.toFixed(0)} ms (${pkg.Paragraphs.length} of ${wholePackage.Paragraphs.length} items), imported and painted in ${imported.toFixed(0)} ms; read at once in ${readAtOnce.toFixed(0)} ms`);

    // More is read: the document grows at its end; what is on screen stays.
    const revision = Word.getRevision(editor), selection = JSON.stringify(Word.getSelection(editor));
    const before = pkg.Paragraphs.length;
    pkg.Progressive.readMore(20);
    check('more is read', [pkg.Paragraphs.length > before, Word.isLoadComplete(editor), Word.getRenderStats(editor).Blocks], [true, false, pkg.Paragraphs.length]);
    check('the revision, the selection and the screen are what they were', [Word.getRevision(editor), JSON.stringify(Word.getSelection(editor)), Word.isModified(editor), samePixels(pixelsOf(canvas), wholeFirstScreen)], [revision, selection, false, true]);

    // Typing while the rest is read: an edit like any other, with undo.
    const own = Word.getSelection(editor);
    Word.setSelection(editor, own.FocusBlockId, 0, own.FocusBlockId, 0);
    Word.insertText(editor, 'TYPED');
    const typedRevision = Word.getRevision(editor);
    check('typing while the document is read', [Word.isLoadComplete(editor), Word.canUndo(editor), Word.isModified(editor), typedRevision !== revision], [false, true, true, true]);
    pkg.Progressive.readMore(20);
    Word.insertText(editor, '!');
    pkg.Progressive.readMore(20);
    check('what was typed stays as more is read', [Word.isLoadComplete(editor), Word.getRenderStats(editor).Blocks, Word.canUndo(editor)], [false, pkg.Paragraphs.length, true]);
    // Scrolling down while it is read: what is there is laid out and shown.
    const sizeBefore = Word.getDocumentSize(editor);
    Word.setViewport(editor, viewport(Math.max(0, sizeBefore.Height - 400)));
    const sizeAfter = Word.getDocumentSize(editor);
    check('scrolling to the end of what is laid out lays out more of what is read', [sizeAfter.Height > sizeBefore.Height, sizeAfter.LayoutComplete, inkOf(pixelsOf(canvas)) > 200], [true, false, true]);
    Word.setViewport(editor, viewport(0));

    // The rest, in the package's own slices: the page can answer meanwhile.
    let longest = 0, last = performance.now();
    const ticker = setInterval(() => { const now = performance.now(); longest = Math.max(longest, now - last); last = now; }, 4);
    t0 = performance.now();
    check('whenLoadComplete resolves when the document is all there', [await Word.whenLoadComplete(editor), Word.isLoadComplete(editor), pkg.Progressive.IsComplete, OF.word.isLoadComplete(pkg)], [true, true, true, true]);
    const readRest = performance.now() - t0;
    clearInterval(ticker);
    const loadNotices = changes.filter((change) => change.LoadChanged);
    check('the listeners were told as it was read, and at the end', [loadNotices.length >= 2, loadNotices.every((change) => !change.DocumentChanged && !change.SelectionChanged && !change.LayoutChanged), loadNotices[loadNotices.length - 1].LoadComplete, loadNotices.slice(0, -1).every((change) => !change.LoadComplete)], [true, true, true, true]);
    check(`no slice held the page for long (longest ${longest.toFixed(0)} ms)`, longest < 250, true);
    check('the layout is not complete when the document is just read', [Word.isLayoutComplete(editor)], [false]);
    stopFonts();
    check('the faces found as the rest was read complete the list', [[...new Set([...earlyFonts, ...discovered])].sort().join(' '), discovered.every((face) => !earlyFonts.includes(face))], [wholeFonts.join(' '), true]);
    // The text is the whole document's with what was typed; undo takes the typing back and keeps the rest.
    const typedText = Word.getText(editor);
    check('the document is all there, with what was typed', [typedText.length, typedText.startsWith('TYPED!'), typedText.slice(6) === wholeText], [wholeText.length + 6, true, true]);
    Word.undo(editor); Word.undo(editor);
    check('undo takes the typing back and keeps what was read since', [Word.getText(editor) === wholeText, Word.getRevision(editor), Word.isModified(editor), Word.canRedo(editor)], [true, revision, false, true]);
    Word.redo(editor);
    check('redo brings it back, in the whole document', [Word.getText(editor).length, Word.getText(editor).startsWith('TYPED'), Word.getRevision(editor)], [wholeText.length + 5, true, typedRevision]);
    Word.undo(editor);
    await Word.whenLayoutComplete(editor);
    check('read front first and laid out in slices, the pages and the lines are those of the document read at once',
        [Word.getPageCount(editor), JSON.stringify(Word.getPageFrames(editor)) === wholeFrames, linesOf(editor) === wholeLines],
        [pageCount, true, true]);
    check('and the first screen is the same pixels', samePixels(pixelsOf(canvas), wholeFirstScreen), true);
    {
        // How the slices went, for a host that wants to know what held its page: the reading's and the layout's.
        const stats = Word.getRenderStats(editor), reading = pkg.Progressive;
        check('the package and the editor say how their slices went', [reading.Slices >= 2, reading.LongestSlice >= 1 && reading.LongestSlice <= reading.Slices, reading.LongestSliceMs > 0, reading.LongestSliceItems >= 0, reading.LongestSlicePassMs <= reading.LongestSliceMs, reading.LongestSliceNotifyMs <= reading.LongestSliceMs,
            stats.LayoutSlices >= 2, stats.LongestLayoutSlice >= 1 && stats.LongestLayoutSlice <= stats.LayoutSlices, stats.LongestLayoutSliceMs > 0, stats.LongestLayoutSliceLayoutMs <= stats.LongestLayoutSliceMs, stats.LongestLayoutSlicePasses >= 1, stats.LongestLayoutSliceFrom >= 1],
            [true, true, true, true, true, true, true, true, true, true, true, true]);
        note(`reading: ${reading.Slices} slices, the longest ${reading.LongestSliceMs.toFixed(1)} ms; layout: ${stats.LayoutSlices} slices, the longest ${stats.LongestLayoutSliceMs.toFixed(1)} ms`);
    }
    note(`the rest was read in ${readRest.toFixed(0)} ms (${loadNotices.length} notices, longest time without an answer ${longest.toFixed(0)} ms)`);
    off();
    Word.dispose(editor); canvas.remove();

    // ---- what goes by the whole document reads the rest at once ------------------
    const fresh = async () => { const p = await OF.word.openProgressive(long, { autoStart: false }); const [e, c] = editorOn('fresh'); Word.importFromDocxPackage(e, p); return [e, c, p]; };
    const forced = async (label, run, expected) => {
        const [e, c, p] = await fresh();
        const pending = !Word.isLoadComplete(e);
        const got = run(e, p);
        check(label, [pending, Word.isLoadComplete(e), p.Progressive.IsComplete, got], [true, true, true, expected]);
        Word.dispose(e); c.remove();
    };
    await forced('getText reads the rest', (e) => Word.getText(e) === wholeText, true);
    await forced('getPageCount reads and lays out the rest', (e) => Word.getPageCount(e), pageCount);
    await forced('find sees the whole document', (e) => Word.findText(e, 'docx4j', false).length, Word.findText(whole, 'docx4j', false).length);
    await forced('select all selects the whole document', (e) => { Word.selectAll(e); return Word.getSelection(e).FocusBlockId !== Word.getSelection(e).AnchorBlockId && Word.getRenderStats(e).Blocks; }, Word.getRenderStats(whole).Blocks);
    const pressCtrlEnd = (e) => { Word.getInputElement(e).dispatchEvent(new KeyboardEvent('keydown', { key: 'End', ctrlKey: true, bubbles: true, cancelable: true })); const rect = Word.getCaretRect(e); return rect === null ? null : [Math.round(rect.X), Math.round(rect.Y)]; };
    const wholeEnd = pressCtrlEnd(whole);
    await forced('Ctrl+End goes to the end of the whole document', (e) => pressCtrlEnd(e), wholeEnd);
    await forced('saving saves the whole document', (e) => Word.exportDocumentPackage(e).Paragraphs.length, Word.exportDocumentPackage(whole).Paragraphs.length);
    await forced('getRequiredFonts of the editor is of the whole document', (e) => Word.getRequiredFonts(e).map((face) => `${face.family}|${face.bold}|${face.italic}`).sort().join(' '), Word.getRequiredFonts(whole).map((face) => `${face.family}|${face.bold}|${face.italic}`).sort().join(' '));
    await forced('the layout fingerprint is of the whole document', (e) => JSON.stringify(Word.getLayoutFingerprint(e)) === JSON.stringify(Word.getLayoutFingerprint(whole)), true);
    await forced('completeLoad reads the rest', (e) => { Word.completeLoad(e); return Word.getRenderStats(e).Blocks; }, Word.getRenderStats(whole).Blocks);
    {
        // An arrow key or a click while it is read asks for nothing.
        const [e, c, p] = await fresh();
        Word.getInputElement(e).dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true, cancelable: true }));
        Word.getInputElement(e).dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true }));
        check('moving the caret leaves the rest for later', [Word.isLoadComplete(e), p.Progressive.IsComplete], [false, false]);
        // Another document loaded: the editor is done with the package.
        Word.setContent(e, 'another document');
        p.Progressive.complete();
        check('a document loaded meanwhile is not given the rest', [Word.isLoadComplete(e), Word.getText(e)], [true, 'another document']);
        Word.dispose(e); c.remove();
    }
    {
        // Without a viewport the canvas shows the whole document: it is read first.
        const p = await OF.word.openProgressive(long, { autoStart: false });
        const c = document.createElement('canvas'); document.body.appendChild(c);
        const e = Word.initializeCanvas(c);
        Word.importFromDocxPackage(e, p);
        check('without a viewport the whole document is read by the import', [Word.isLoadComplete(e), p.Progressive.IsComplete, Word.getText(e) === wholeText], [true, true, true]);
        Word.dispose(e); c.remove();
    }

    // ---- the viewer ---------------------------------------------------------------
    {
        const reference = Word.createViewer(document.createElement('canvas'));
        reference.load(wholePackage);
        const referenceFrames = JSON.stringify(reference.getPageFrames());
        const p = await OF.word.openProgressive(long, { autoStart: false });
        const viewer = Word.createViewer(document.createElement('canvas'));
        const progress = [];
        viewer.load(p, { progressive: true });
        const stop = viewer.onLayoutProgress((state) => progress.push({ ...state }));
        const firstPages = viewer.getPageFrames().length;
        check('the viewer shows the first pages of a document that is still being read', [viewer.isLoadComplete(), viewer.isLayoutComplete(), firstPages >= 1, firstPages < 100], [false, false, true, true]);
        const target = document.createElement('canvas');
        check('a first page is painted meanwhile', [viewer.renderPage(target, 0, 0.5), viewer.isLoadComplete()], [true, false]);
        check('whenLoadComplete resolves when the document is read', [await viewer.whenLoadComplete(), viewer.isLoadComplete(), p.Progressive.IsComplete], [true, true, true]);
        await viewer.whenLayoutComplete();
        check('the pages arrive as it is read and laid out, up to those of the document read at once',
            [progress.length >= 2, progress.every((state, i) => i === 0 || state.PageCount >= progress[i - 1].PageCount), progress[progress.length - 1].LayoutComplete, progress[progress.length - 1].LoadComplete, JSON.stringify(viewer.getPageFrames()) === referenceFrames],
            [true, true, true, true, true]);
        stop();
        // A page there is not yet, the page count: of the whole document.
        const late = await OF.word.openProgressive(long, { autoStart: false });
        const second = Word.createViewer(document.createElement('canvas'));
        second.load(late, { progressive: true });
        check('renderPage of a page that is not there yet reads and lays out the document', [second.isLoadComplete(), second.renderPage(target, pageCount - 1, 0.5), second.isLoadComplete(), second.getPageCount()], [false, true, true, reference.getPageCount()]);
        // Not asked to show it early: the document is read by the load.
        const plain = await OF.word.openProgressive(long, { autoStart: false });
        const third = Word.createViewer(document.createElement('canvas'));
        third.load(plain);
        check('without { progressive: true } the viewer reads the document first', [third.isLoadComplete(), plain.Progressive.IsComplete, JSON.stringify(third.getPageFrames()) === referenceFrames], [true, true, true]);
    }
} catch (e) { fail(e); }
finish();
