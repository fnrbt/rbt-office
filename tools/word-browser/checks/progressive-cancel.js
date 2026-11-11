// The reading of a DOCX opened front first is cancelled: by the host
// (Progressive.cancel, cancelLoad of an editor or a viewer), and by an editor
// or a viewer that is done with the package (another document loaded,
// dispose) unless something else still shows it. A cancelled package reads no
// further slice; what was read stays, as of a package that failed.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    // docx4j-getting-started with its body repeated 8 times (some 600 pages).
    const zip = await window.JSZip.loadAsync(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>'), sect = xml.lastIndexOf('<w:sectPr', close);
    zip.file('word/document.xml', xml.slice(0, open) + xml.slice(open, sect).repeat(8) + xml.slice(sect, close) + xml.slice(close));
    const long = await zip.generateAsync({ type: 'uint8array' });
    const other = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const viewport = { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 800 };
    const editorOn = () => { const canvas = document.createElement('canvas'); document.body.appendChild(canvas); const editor = Word.initializeCanvas(canvas); Word.setViewport(editor, viewport); return [editor, canvas]; };
    const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const within = (promise) => Promise.race([promise, wait(5000).then(() => 'timed out')]);
    const attempt = (run) => { try { return run(); } catch (e) { return 'threw: ' + String(e && e.message || e); } };
    const cancelledMessage = 'The reading of the document was cancelled';
    // Whether a package that reads by itself has stopped: no slice, no item in a while.
    const stopped = async (pkg) => { const slices = pkg.Progressive.Slices, items = pkg.Paragraphs.length; await wait(250); return [pkg.Progressive.Slices - slices, pkg.Paragraphs.length - items]; };

    // ---- the package ----------------------------------------------------------------
    {
        const pkg = await OF.word.openProgressive(long, {});
        await wait(60);
        const calls = [];
        pkg.Progressive.subscribe((first, count, complete) => calls.push([count, complete, pkg.Progressive.Failed, pkg.Progressive.Cancelled]));
        const pending = pkg.Progressive.whenComplete().then(() => 'resolved', (e) => String(e && e.message || e));
        const slicesBefore = pkg.Progressive.Slices, itemsBefore = pkg.Paragraphs.length, progressBefore = pkg.Progressive.Progress();
        check('the package reads by itself', [slicesBefore >= 1, itemsBefore > 64, pkg.Progressive.IsComplete, pkg.Progressive.Cancelled], [true, true, false, false]);
        check('cancel says that it cancelled', [pkg.Progressive.cancel()], [true]);
        check('the package is as one that failed, and says that it was cancelled', [pkg.Progressive.Cancelled, pkg.Progressive.Failed, pkg.Progressive.IsComplete, pkg.Progressive.Error, OF.word.isLoadFailed(pkg), OF.word.isLoadComplete(pkg)], [true, true, false, cancelledMessage, true, false]);
        check('the listeners were told, and the promise of the package rejects', [calls.length, calls.length > 0 && calls[calls.length - 1].slice(1), await pending], [1, [false, true, true], cancelledMessage]);
        check('no slice reads on, no item arrives', [await stopped(pkg), pkg.Paragraphs.length >= itemsBefore], [[0, 0], true]);
        check('nothing reads it on', [attempt(() => pkg.Progressive.readMore(50)), attempt(() => pkg.Progressive.complete()), attempt(() => pkg.Progressive.start()), await stopped(pkg), pkg.Progressive.cancel()], [false, undefined, undefined, [0, 0], false]);
        check('how far it got stays what it was', [pkg.Progressive.Progress() >= progressBefore, pkg.Progressive.Progress() < 1, pkg.Progressive.BlockCount() > 64], [true, true, true]);
        // What was read can be shown.
        const [editor, canvas] = editorOn();
        Word.importFromDocxPackage(editor, pkg);
        check('an editor shows what a cancelled package had read', [Word.isLoadFailed(editor), Word.getLoadError(editor), Word.getRenderStats(editor).Blocks > 64, attempt(() => Word.getText(editor).length > 1000)], [true, cancelledMessage, true, true]);
        Word.dispose(editor); canvas.remove();
        // A package that is read is not cancelled.
        const read = await OF.word.openProgressive(long, { autoStart: false });
        read.Progressive.complete();
        check('a package that is read is not cancelled', [read.Progressive.cancel(), read.Progressive.IsComplete, read.Progressive.Failed, read.Progressive.Cancelled, read.Progressive.BlockCount() > 5000], [false, true, false, false, true]);
    }

    // ---- an editor ------------------------------------------------------------------
    {
        const pkg = await OF.word.openProgressive(long, {});
        const [editor, canvas] = editorOn();
        const changes = [];
        Word.onChange(editor, (change) => changes.push({ ...change }));
        Word.importFromDocxPackage(editor, pkg);
        await wait(60);
        const waiting = Word.whenLoadComplete(editor);
        check('cancelLoad stops the reading of the document on display', [Word.isLoadComplete(editor), Word.cancelLoad(editor), pkg.Progressive.Cancelled, await stopped(pkg)], [false, true, true, [0, 0]]);
        const failures = changes.filter((change) => change.LoadFailed);
        check('the editor keeps what was read, failed', [Word.isLoadFailed(editor), Word.getLoadError(editor), Word.isLoadComplete(editor), await waiting, failures.filter((change) => change.LoadChanged).length, failures.length > 0 && failures[0].LoadError, Word.getRenderStats(editor).Blocks === pkg.Paragraphs.length], [true, cancelledMessage, false, false, 1, cancelledMessage, true]);
        await Word.whenLayoutComplete(editor);
        check('it is laid out, read and not saved', [Word.isLayoutComplete(editor), attempt(() => Word.getPageCount(editor) > 2), attempt(() => Word.getText(editor).length > 1000), String(attempt(() => Word.exportDocumentPackage(editor))).includes('cannot be saved'), Word.cancelLoad(editor)], [true, true, true, true, false]);
        Word.dispose(editor); canvas.remove();
    }
    {
        // The host cancels the package under an editor.
        const pkg = await OF.word.openProgressive(long, {});
        const [editor, canvas] = editorOn();
        Word.importFromDocxPackage(editor, pkg);
        check('a package cancelled under an editor ends its load', [pkg.Progressive.cancel(), Word.isLoadFailed(editor), Word.getLoadError(editor), await Word.whenLoadComplete(editor)], [true, true, cancelledMessage, false]);
        Word.dispose(editor); canvas.remove();
    }
    {
        // An editor that is done with a package cancels it.
        // (The packages are opened before any of them is shown: opening one takes longer than
        // reading the rest of another.)
        const [first, second, third, fourth] = [await OF.word.openProgressive(long, { autoStart: false }), await OF.word.openProgressive(long, { autoStart: false }), await OF.word.openProgressive(long, { autoStart: false }), await OF.word.openProgressive(long, { autoStart: false })];
        const [editor, canvas] = editorOn();
        Word.importFromDocxPackage(editor, first);
        let firstSettlements = 0;
        const firstWaiters = [Word.whenLoadComplete(editor), Word.whenLoadComplete(editor)]
            .map((p) => p.then((value) => { firstSettlements++; return value; }));
        // (The same package loaded again: still shown.)
        Word.importFromDocxPackage(editor, first);
        check('reloading the same package ends the outgoing generation', [await within(Promise.all(firstWaiters)), firstSettlements], [[false, false], 2]);
        const samePackageWaiter = Word.whenLoadComplete(editor);
        check('the same package loaded again is read on', [first.Progressive.Cancelled, Word.isLoadFailed(editor), Word.isLoadComplete(editor)], [false, false, false]);
        Word.importFromDocxPackage(editor, second);
        check('the replaced generation remains cancelled while a new package loads', await within(samePackageWaiter), false);
        check('another package loaded: the first one is cancelled', [first.Progressive.Cancelled, await stopped(first), second.Progressive.Cancelled, Word.isLoadFailed(editor)], [true, [0, 0], false, false]);
        const secondWaiter = Word.whenLoadComplete(editor);
        Word.importFromDocxPackage(editor, other);
        check('replacement by a completed document settles the old waiter false and the new one true', [await within(secondWaiter), await Word.whenLoadComplete(editor)], [false, true]);
        check('a read document loaded: the one being read is cancelled', [second.Progressive.Cancelled, await stopped(second), Word.isLoadFailed(editor), Word.isLoadComplete(editor)], [true, [0, 0], false, true]);
        Word.importFromDocxPackage(editor, third);
        const thirdWaiter = Word.whenLoadComplete(editor);
        Word.setContent(editor, 'another document');
        check('setContent settles the outgoing waiter false', [await within(thirdWaiter), await Word.whenLoadComplete(editor)], [false, true]);
        check('setContent: cancelled', [third.Progressive.Cancelled, Word.isLoadFailed(editor), Word.getText(editor)], [true, false, 'another document']);
        Word.importFromDocxPackage(editor, fourth);
        const fourthWaiter = Word.whenLoadComplete(editor);
        Word.dispose(editor); canvas.remove();
        check('dispose settles the outgoing waiter false', [await within(fourthWaiter), await Word.whenLoadComplete(editor)], [false, false]);
        check('dispose: cancelled', [fourth.Progressive.Cancelled, await stopped(fourth)], [true, [0, 0]]);
        check('dispose: the editor says that its document was not read to its end', [Word.isLoadComplete(editor), Word.isLoadFailed(editor), Word.getLoadError(editor)], [false, true, cancelledMessage]);
        // A package the editor read to its end is not.
        const fifth = await OF.word.openProgressive(long, { autoStart: false });
        const [reader, readerCanvas] = editorOn();
        Word.importFromDocxPackage(reader, fifth);
        const completedWaiter = Word.whenLoadComplete(reader);
        Word.completeLoad(reader);
        check('successful completion settles its generation true', [await within(completedWaiter), await Word.whenLoadComplete(reader)], [true, true]);
        Word.dispose(reader); readerCanvas.remove();
        check('a document that was read is not cancelled by dispose', [fifth.Progressive.Cancelled, fifth.Progressive.IsComplete, fifth.Progressive.Failed], [false, true, false]);
    }
    {
        // History states extended by later slices retain the same revision.
        const pkg = await OF.word.openProgressive(long, { autoStart: false });
        const [editor, canvas] = editorOn();
        Word.importFromDocxPackage(editor, pkg);
        const openedRevision = Word.getRevision(editor);
        const proxy = Word.getInputElement(editor);
        proxy.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true, cancelable: true }));
        const typedRevision = Word.getRevision(editor);
        pkg.Progressive.readMore(50);
        Word.completeLoad(editor);
        check('undo after progressive slices restores the opened revision', [Word.undo(editor), Word.getRevision(editor), Word.isModified(editor), Word.exportDocumentPackage(editor).Unchanged], [true, openedRevision, false, true]);
        check('redo after progressive slices restores the edited revision', [Word.redo(editor), Word.getRevision(editor), Word.isModified(editor)], [true, typedRevision, true]);
        Word.dispose(editor); canvas.remove();
    }

    // ---- the viewer, and a package shown twice -----------------------------------
    {
        const pkg = await OF.word.openProgressive(long, {});
        const viewer = Word.createViewer(document.createElement('canvas'));
        viewer.load(pkg, { progressive: true });
        const progress = [];
        viewer.onLayoutProgress((state) => progress.push({ ...state }));
        await wait(60);
        check('the viewer cancels the reading of its document', [viewer.cancelLoad(), pkg.Progressive.Cancelled, viewer.isLoadFailed(), viewer.getLoadError(), viewer.isLoadComplete(), await viewer.whenLoadComplete(), await stopped(pkg), viewer.cancelLoad()], [true, true, true, cancelledMessage, false, false, [0, 0], false]);
        await viewer.whenLayoutComplete();
        check('what was read is laid out', [viewer.isLayoutComplete(), viewer.getPageCount() > 2, progress.length > 0 && progress[progress.length - 1].LoadFailed], [true, true, true]);

        const [first, second] = [await OF.word.openProgressive(long, { autoStart: false }), await OF.word.openProgressive(long, { autoStart: false })];
        viewer.load(first, { progressive: true });
        check('a package loaded after a cancelled one is not failed', [viewer.isLoadFailed(), viewer.isLoadComplete(), first.Progressive.Cancelled], [false, false, false]);
        viewer.load(first, { progressive: true });
        check('the same package loaded again is read on', [first.Progressive.Cancelled, viewer.isLoadFailed()], [false, false]);
        viewer.load(second, { progressive: true });
        check('the viewer loads another package: the first one is cancelled', [first.Progressive.Cancelled, await stopped(first), second.Progressive.Cancelled], [true, [0, 0], false]);
        viewer.dispose();
        check('the viewer is disposed of: cancelled, nothing more is laid out', [second.Progressive.Cancelled, await stopped(second)], [true, [0, 0]]);
        check('the viewer disposed of says that its document was not read to its end', [viewer.isLoadComplete(), viewer.isLoadFailed(), viewer.getLoadError(), await viewer.whenLoadComplete()], [false, true, cancelledMessage, false]);
        // Disposed of at once, before the first slice (a host that closes a document it has just opened).
        {
            const atOnce = await OF.word.openProgressive(long, {});
            const short = Word.createViewer(document.createElement('canvas'));
            short.load(atOnce, { progressive: true });
            short.dispose();
            check('a viewer disposed of right after load: not complete, failed as cancelled', [short.isLoadComplete(), short.isLoadFailed(), short.getLoadError(), atOnce.Progressive.Cancelled, atOnce.Progressive.IsComplete], [false, true, cancelledMessage, true, false]);
            // A document that was read to its end is complete after dispose as before it.
            const whole = await OF.word.openProgressive(long, { autoStart: false });
            whole.Progressive.complete();
            const done = Word.createViewer(document.createElement('canvas'));
            done.load(whole, { progressive: true });
            done.dispose();
            check('a viewer of a document that was read: complete after dispose', [done.isLoadComplete(), done.isLoadFailed()], [true, false]);
        }

        // An editor and a viewer show the same package: it is read while one of them does.
        const shared = await OF.word.openProgressive(long, {});
        const [editor, canvas] = editorOn();
        Word.importFromDocxPackage(editor, shared);
        const preview = Word.createViewer(document.createElement('canvas'));
        preview.load(shared, { progressive: true });
        const editorWaiter = Word.whenLoadComplete(editor);
        const viewerWaiter = preview.whenLoadComplete();
        Word.dispose(editor); canvas.remove();
        check('disposing one consumer settles only its waiter false', [await within(editorWaiter), shared.Progressive.Cancelled, preview.isLoadFailed()], [false, false, false]);
        check('a package a viewer still shows is not cancelled by the editor', [shared.Progressive.Cancelled, preview.isLoadFailed(), preview.isLoadComplete()], [false, false, false]);
        const slices = shared.Progressive.Slices; await wait(120);
        check('it is read on', [shared.Progressive.Slices > slices || shared.Progressive.IsComplete], [true]);
        preview.completeLoad();
        check('the shared viewer completes its own generation successfully', await within(viewerWaiter), true);
        preview.dispose();
        check('the last one to let go of it cancels it', [shared.Progressive.Cancelled || shared.Progressive.IsComplete, await stopped(shared)], [true, [0, 0]]);
    }
    {
        // Replacing the editor's generation does not cancel a shared package
        // or transfer the outgoing waiter's result to the replacement.
        const shared = await OF.word.openProgressive(long, { autoStart: false });
        const [editor, canvas] = editorOn();
        const preview = Word.createViewer(document.createElement('canvas'));
        Word.importFromDocxPackage(editor, shared);
        preview.load(shared, { progressive: true });
        const outgoing = Word.whenLoadComplete(editor);
        const stillReading = preview.whenLoadComplete();
        Word.setContent(editor, 'replacement');
        check('shared-package replacement settles the editor only', [await within(outgoing), await Word.whenLoadComplete(editor), shared.Progressive.Cancelled, preview.isLoadFailed()], [false, true, false, false]);
        preview.completeLoad();
        check('the other consumer can complete after replacement', await within(stillReading), true);
        preview.dispose();
        Word.dispose(editor); canvas.remove();
    }
} catch (e) { fail(e); }
finish();
