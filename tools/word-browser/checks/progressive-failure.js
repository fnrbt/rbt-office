// A DOCX opened front first whose body cannot be read to its end (its XML is
// not well formed at paragraph 7000 of 9000): the load ends in a failed state.
// What was read stays on display and is laid out to its end; the text, find
// and the page count answer for it; the package, the editor and the viewer
// say that it failed and why; it is not saved.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const base = await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx');
    // `broken(k)`: the XML of paragraph k when it is the one that is not well formed (else null).
    const documentWith = async (count, broken) => {
        const zip = await window.JSZip.loadAsync(base);
        const xml = await zip.file('word/document.xml').async('string');
        const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>'), sect = xml.lastIndexOf('<w:sectPr', close);
        const parts = [];
        for (let k = 1; k <= count; k++)
            parts.push(broken(k) ?? `<w:p><w:r><w:t>Paragraph ${k} of a long document, with words enough to say which one it is.</w:t></w:r></w:p>`);
        zip.file('word/document.xml', xml.slice(0, open) + parts.join('') + xml.slice(sect, close) + xml.slice(close));
        return zip.generateAsync({ type: 'uint8array' });
    };
    const mismatched = '<w:p><w:r><w:t>not well formed</w:r></w:p>';
    const failing = await documentWith(9000, (k) => k === 7000 ? mismatched : null);
    const good = await documentWith(6999, () => null);
    const viewport = (top) => ({ scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
    const editorOn = (id) => { const canvas = document.createElement('canvas'); canvas.id = id; document.body.appendChild(canvas); const editor = Word.initializeCanvas(canvas); Word.setViewport(editor, viewport(0)); return [editor, canvas]; };
    const within = (promise, ms) => Promise.race([promise, new Promise((resolve) => setTimeout(() => resolve('timed out'), ms))]);
    const attempt = (run) => { try { return run(); } catch (e) { return 'threw: ' + String(e && e.message || e); } };

    // The reference: the 6999 paragraphs before the broken one, read at once.
    const [whole, wholeCanvas] = editorOn('whole');
    Word.importFromDocxPackage(whole, await OF.processors.forEditorPackage(good));
    const goodText = Word.getText(whole);
    const goodPages = Word.getPageCount(whole);
    const goodFound = Word.findText(whole, 'Paragraph 69', false).length;
    check('the reference: 6999 paragraphs', [Word.getRenderStats(whole).Blocks, goodPages > 100, goodFound], [6999, true, 111]);

    // ---- the package ----------------------------------------------------------------
    {
        const pkg = await OF.word.openProgressive(failing, { autoStart: false });
        check('the package is handed over with the start of the body', [pkg.Progressive.IsComplete, pkg.Progressive.Failed, pkg.Progressive.Error, OF.word.isLoadFailed(pkg), OF.word.getLoadError(pkg)], [false, false, null, false, null]);
        const calls = [];
        pkg.Progressive.subscribe((first, count, complete) => calls.push([first, count, complete, pkg.Progressive.Failed]));
        const completed = attempt(() => pkg.Progressive.complete());
        check('reading the rest does not throw: it ends failed, with what was read', [completed === undefined, pkg.Progressive.IsComplete, pkg.Progressive.Failed, typeof pkg.Progressive.Error, pkg.Paragraphs.length, OF.word.isLoadComplete(pkg), OF.word.isLoadFailed(pkg), OF.word.getLoadError(pkg) === pkg.Progressive.Error],
            [true, false, true, 'string', 6999, false, true, true]);
        check('the listeners are told a last time, the failure set', [calls.length >= 1, calls[calls.length - 1][2], calls[calls.length - 1][3], calls.reduce((sum, call) => sum + call[1], 0)], [true, false, true, 6999 - 64]);
        check('nothing more is read', [attempt(() => pkg.Progressive.readMore(50)), attempt(() => pkg.Progressive.complete()), pkg.Paragraphs.length], [false, undefined, 6999]);
        check('the promise of the complete package rejects with the error', [await OF.word.whenLoadComplete(pkg).then(() => 'resolved', (e) => String(e && e.message || e) === pkg.Progressive.Error)], [true]);
        note(`the error: ${pkg.Progressive.Error}`);
    }

    // ---- an editor, the rest read in the package's own slices ---------------------
    {
        const pkg = await OF.word.openProgressive(failing, { autoStart: false });
        const [editor, canvas] = editorOn('front');
        const changes = [];
        Word.onChange(editor, (change) => changes.push({ ...change }));
        Word.importFromDocxPackage(editor, pkg);
        check('the import returns with the document still being read', [Word.isLoadComplete(editor), Word.isLoadFailed(editor), Word.getLoadError(editor), Word.isLayoutComplete(editor)], [false, false, null, false]);
        check('whenLoadComplete resolves to false when the document cannot be read on', [await within(Word.whenLoadComplete(editor), 60000)], [false]);
        check('the load is failed, not complete', [Word.isLoadComplete(editor), Word.isLoadFailed(editor), typeof Word.getLoadError(editor), Word.getLoadError(editor) === pkg.Progressive.Error], [false, true, 'string', true]);
        const failures = changes.filter((change) => change.LoadFailed);
        check('the listeners are told once that it failed', [failures.length, failures.length > 0 && failures[0].LoadChanged, failures.length > 0 && failures[0].LoadComplete, failures.length > 0 && failures[0].LoadError === Word.getLoadError(editor), changes.filter((change) => change.LoadChanged && !change.LoadFailed).every((change) => change.LoadError === null && !change.LoadComplete)], [1, true, false, true, true]);
        check('the layout completes over what was read', [await within(Word.whenLayoutComplete(editor).then(() => 'laid out'), 120000), Word.isLayoutComplete(editor), Word.getDocumentSize(editor).LayoutComplete, Word.getDocumentSize(editor).PageCount], ['laid out', true, true, goodPages]);
        check('the text, find and the page count are of what was read', [attempt(() => Word.getText(editor) === goodText), attempt(() => Word.findText(editor, 'Paragraph 69', false).length), attempt(() => Word.getPageCount(editor)), Word.getRenderStats(editor).Blocks], [true, goodFound, goodPages, 6999]);
        const saved = attempt(() => Word.exportDocumentPackage(editor));
        check('saving refuses, and says why', [typeof saved === 'string' && saved.startsWith('threw: ') && saved.includes('cannot be saved') && saved.includes(Word.getLoadError(editor)), String(attempt(() => Word.exportToDocx(editor))).includes('cannot be saved')], [true, true]);
        note(String(saved));
        check('it can still be edited and resolves at once from here on', [attempt(() => { Word.insertText(editor, 'X'); return Word.getText(editor).startsWith('X'); }), await Word.whenLoadComplete(editor), Word.isLoadFailed(editor)], [true, false, true]);
        // Another document loaded: the failure was the other document's.
        Word.setContent(editor, 'another document');
        check('a document loaded afterwards is not failed', [Word.isLoadFailed(editor), Word.getLoadError(editor), Word.isLoadComplete(editor), await Word.whenLoadComplete(editor), attempt(() => Word.exportDocumentPackage(editor).Paragraphs.length)], [false, null, true, true, 1]);
        Word.dispose(editor); canvas.remove();
    }

    // ---- what goes by the whole document: of what can be read, at once -------------
    const forced = async (label, run, expected) => {
        const pkg = await OF.word.openProgressive(failing, { autoStart: false });
        const [editor, canvas] = editorOn('forced');
        Word.importFromDocxPackage(editor, pkg);
        const pending = !Word.isLoadComplete(editor) && !Word.isLoadFailed(editor);
        const got = attempt(() => run(editor));
        check(label, [pending, Word.isLoadFailed(editor), Word.isLoadComplete(editor), got], [true, true, false, expected]);
        Word.dispose(editor); canvas.remove();
    };
    await forced('getText at open gives the text that was read', (e) => Word.getText(e) === goodText, true);
    await forced('getPageCount at open gives its pages', (e) => Word.getPageCount(e), goodPages);
    await forced('find at open finds in it', (e) => Word.findText(e, 'Paragraph 69', false).length, goodFound);
    await forced('completeLoad reads what there is', (e) => { Word.completeLoad(e); return Word.getRenderStats(e).Blocks; }, 6999);
    await forced('saving at open refuses', (e) => String(attempt(() => Word.exportDocumentPackage(e))).includes('cannot be saved'), true);
    {
        // A package that failed before an editor saw it, and an editor without a viewport.
        const pkg = await OF.word.openProgressive(failing, { autoStart: false });
        pkg.Progressive.complete();
        const [editor, canvas] = editorOn('late');
        Word.importFromDocxPackage(editor, pkg);
        check('a package that has failed is shown as far as it was read', [Word.isLoadFailed(editor), Word.isLoadComplete(editor), Word.getRenderStats(editor).Blocks, await Word.whenLoadComplete(editor), String(attempt(() => Word.exportDocumentPackage(editor))).includes('cannot be saved')], [true, false, 6999, false, true]);
        Word.dispose(editor); canvas.remove();
        const other = await OF.word.openProgressive(failing, { autoStart: false });
        const plainCanvas = document.createElement('canvas'); document.body.appendChild(plainCanvas);
        const plain = Word.initializeCanvas(plainCanvas);
        Word.importFromDocxPackage(plain, other);
        check('without a viewport the import reads what there is', [Word.isLoadFailed(plain), Word.getRenderStats(plain).Blocks, attempt(() => Word.getText(plain) === goodText)], [true, 6999, true]);
        Word.dispose(plain); plainCanvas.remove();
    }

    // ---- the viewer -----------------------------------------------------------------
    {
        const pkg = await OF.word.openProgressive(failing, { autoStart: false });
        const viewer = Word.createViewer(document.createElement('canvas'));
        const progress = [];
        viewer.load(pkg, { progressive: true });
        const stop = viewer.onLayoutProgress((state) => progress.push({ ...state }));
        check('the viewer shows the first pages', [viewer.isLoadComplete(), viewer.isLoadFailed(), viewer.getLoadError(), viewer.getPageFrames().length >= 1], [false, false, null, true]);
        check('its whenLoadComplete resolves to false', [await within(viewer.whenLoadComplete(), 60000), viewer.isLoadComplete(), viewer.isLoadFailed(), viewer.getLoadError() === pkg.Progressive.Error], [false, false, true, true]);
        await within(viewer.whenLayoutComplete(), 120000);
        const last = progress[progress.length - 1];
        check('the pages of what was read are laid out to the end', [viewer.isLayoutComplete(), viewer.getPageFrames().length, attempt(() => viewer.getPageCount()), last && last.LayoutComplete, last && last.LoadComplete, last && last.LoadFailed, last && last.LoadError === pkg.Progressive.Error], [true, goodPages, goodPages, true, false, true, true]);
        check('a page of it is painted', [attempt(() => viewer.renderPage(document.createElement('canvas'), goodPages - 1, 0.5))], [true]);
        stop();
        // Not asked to show it early, and a package that has failed already.
        const plain = Word.createViewer(document.createElement('canvas'));
        plain.load(await OF.word.openProgressive(failing, { autoStart: false }));
        check('without { progressive: true } the viewer reads what there is', [plain.isLoadFailed(), plain.isLoadComplete(), plain.getPageCount()], [true, false, goodPages]);
        plain.load(await OF.processors.forEditorPackage(good));
        check('a package loaded afterwards is not failed', [plain.isLoadFailed(), plain.getLoadError(), plain.isLoadComplete()], [false, null, true]);
    }

    // ---- other ways of not being well formed, and other places ---------------------
    for (const [label, count, at, xml, read] of [
        ['an element that is not closed', 300, 200, '<w:p><w:r><w:t>not closed</w:t></w:p>', 199],
        ['a character reference that is none', 300, 200, '<w:p><w:r><w:t>a & b</w:t></w:r></w:p>', 199],
        ['a tag that does not end', 300, 200, '<w:p><w:r><w:t>cut off</w:t></w:r></w:p', 199],
        ['among the first blocks', 300, 3, mismatched, 2],
        ['the first block', 300, 1, mismatched, 0],
    ]) {
        const bytes = await documentWith(count, (k) => k === at ? xml : null);
        const pkg = await attempt(() => OF.word.openProgressive(bytes, { autoStart: false }));
        if (typeof pkg === 'string') { check(label, pkg, 'a package'); continue; }
        const [editor, canvas] = editorOn('kinds');
        const got = await attempt(async () => {
            Word.importFromDocxPackage(editor, pkg);
            const complete = await within(Word.whenLoadComplete(editor), 30000);
            return [complete, Word.isLoadFailed(editor), Word.getRenderStats(editor).Blocks, Word.getText(editor).includes(`Paragraph ${at + 1} `), typeof Word.getPageCount(editor)];
        });
        check(`${label}: the load fails with the ${read} paragraphs before it`, got, [false, true, read, false, 'number']);
        Word.dispose(editor); canvas.remove();
    }
} catch (e) { fail(e); }
finish();
