// The faces a document needs and the host has not got: setFontAvailability /
// getPendingFonts / FontsPending, and onFontsDiscovered called - synchronously,
// before the blocks that need the faces are given to an editor or a viewer -
// also when the rest of a document is read at once (getPageCount at open), so
// that a host with the bytes at hand installs them before the layout.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    // docx4j-getting-started with its body four times, then 400 paragraphs in bold Georgia: a face
    // the styles and the first blocks do not name.
    const zip = await window.JSZip.loadAsync(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>'), sect = xml.lastIndexOf('<w:sectPr', close);
    const late = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/><w:b/></w:rPr><w:t>Text in a face that is found late, long enough to go over more than one line of the page it is on, whatever the face.</w:t></w:r></w:p>';
    zip.file('word/document.xml', xml.slice(0, open) + xml.slice(open, sect).repeat(4) + late.repeat(400) + xml.slice(sect, close) + xml.slice(close));
    const long = await zip.generateAsync({ type: 'uint8array' });
    const keyOf = (face) => `${face.family}|${face.bold}|${face.italic}`;
    const georgia = 'Georgia|true|false';
    // A text stack: `faces` are the faces it has; text in a face it has not is set in the default one
    // (narrower here), as a shaper falls back.
    const stack = (faces) => {
        const width = (style) => (faces.has(keyOf({ family: style.FontFamily, bold: !!style.Bold, italic: !!style.Italic })) && style.FontFamily === 'Georgia' ? 1.0 : 0.5);
        return {
            faces,
            isAvailable: (family, bold, italic) => faces.has(keyOf({ family, bold, italic })),
            measureText: (text, style) => text.length * style.FontSize * width(style),
            measureRange: (text, style, start, length) => Math.max(0, Math.min(length, text.length - start)) * style.FontSize * width(style),
        };
    };
    const viewport = { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 800 };
    const editorOn = (text) => { const canvas = document.createElement('canvas'); document.body.appendChild(canvas); const editor = Word.initializeCanvas(canvas); Word.setViewport(editor, viewport); Word.setFontAvailability(editor, text.isAvailable); Word.setTextMeasurers(editor, text.measureText, text.measureRange); return [editor, canvas]; };

    const wholePackage = await OF.processors.forEditorPackage(long);
    const allFaces = Word.getRequiredFonts(wholePackage).map(keyOf).sort();
    const sorted = (faces) => faces.map(keyOf).sort().join(' ');
    check('the document needs a face that is found late', [allFaces.includes(georgia)], [true]);
    // The pages with every face (what the document is), and with all but the late one.
    const [whole, wholeCanvas] = editorOn(stack(new Set(allFaces)));
    Word.importFromDocxPackage(whole, wholePackage);
    const finalPages = Word.getPageCount(whole);
    check('with every face nothing is pending', [Word.getPendingFonts(whole), Word.getDocumentSize(whole).FontsPending], [[], false]);
    const [without, withoutCanvas] = editorOn(stack(new Set(allFaces.filter((face) => face !== georgia))));
    Word.importFromDocxPackage(without, wholePackage);
    const provisionalPages = Word.getPageCount(without);
    check('without the late face the pages are fewer, and the editor says that a face is pending', [provisionalPages < finalPages, Word.getPendingFonts(without).map(keyOf), Word.getDocumentSize(without).FontsPending], [true, [georgia], true]);
    note(`${finalPages} pages; ${provisionalPages} without the late face`);
    Word.dispose(whole); wholeCanvas.remove(); Word.dispose(without); withoutCanvas.remove();
    // A host that does not say which faces it has: nothing is pending.
    {
        const canvas = document.createElement('canvas'); document.body.appendChild(canvas);
        const editor = Word.initializeCanvas(canvas); Word.setViewport(editor, viewport);
        Word.importFromDocxPackage(editor, wholePackage);
        check('a host that does not say which faces it has: nothing is pending', [Word.getPendingFonts(editor), Word.getDocumentSize(editor).FontsPending], [[], false]);
        Word.setFontAvailability(editor, () => false);
        check('a host that has none: all of them', [Word.getPendingFonts(editor).map(keyOf).sort().join(' '), Word.getDocumentSize(editor).FontsPending], [allFaces.join(' '), true]);
        Word.setFontAvailability(editor, null);
        check('and no longer says', [Word.getPendingFonts(editor), Word.getDocumentSize(editor).FontsPending], [[], false]);
        Word.dispose(editor); canvas.remove();
    }

    // ---- getPageCount at open, the bytes of the late face at hand ------------------
    {
        const pkg = await OF.word.openProgressive(long, { autoStart: false });
        const early = Word.getRequiredFonts(pkg).map(keyOf);
        check('the late face is not among those known at open', [early.includes(georgia), early.length > 0], [false, true]);
        const text = stack(new Set(early));
        const [editor, canvas] = editorOn(text);
        const changes = [];
        Word.onChange(editor, (change) => changes.push({ ...change }));
        // (The editor is given the package first: the fonts are still told before it.)
        Word.importFromDocxPackage(editor, pkg);
        const calls = [];
        Word.onFontsDiscovered(pkg, (faces) => {
            calls.push({ faces: faces.map(keyOf), blocks: Word.getRenderStats(editor).Blocks, items: pkg.Paragraphs.length, loaded: Word.isLoadComplete(editor) });
            // The host has the bytes: the faces are in its font set before the text is laid out.
            for (const face of faces) text.faces.add(keyOf(face));
            Word.setTextMeasurers(editor, text.measureText, text.measureRange);
        });
        check('before anything is read on nothing is pending, and asking reads nothing', [Word.getPendingFonts(editor), Word.getDocumentSize(editor).FontsPending, Word.isLoadComplete(editor), pkg.Paragraphs.length < 100], [[], false, false, true]);
        const blocksAtOpen = Word.getRenderStats(editor).Blocks;
        const measurersBefore = Word.getRenderStats(editor).MeasurerVersion;
        const pages = Word.getPageCount(editor);
        check('the callback was called once, in the call, with every face the rest needs', [calls.length, calls.flatMap((call) => call.faces).sort().join(' '), Word.isLoadComplete(editor)], [1, allFaces.filter((face) => !early.includes(face)).join(' '), true]);
        check('it was called before the editor was given the blocks that need them', [calls[0].blocks, calls[0].items > blocksAtOpen, calls[0].loaded], [blocksAtOpen, true, false]);
        check('the page count is the final one the first time', [pages, Word.getPendingFonts(editor), Word.getDocumentSize(editor).FontsPending, Word.getRenderStats(editor).MeasurerVersion - measurersBefore], [finalPages, [], false, 1]);
        check('no notice says that a face is pending', [changes.length > 0, changes.some((change) => change.FontsPending)], [true, false]);
        Word.dispose(editor); canvas.remove();
    }

    // ---- getPageCount at open, the late face still to be fetched -------------------
    {
        const pkg = await OF.word.openProgressive(long, { autoStart: false });
        const early = Word.getRequiredFonts(pkg).map(keyOf);
        // (All the faces but Georgia are at hand in this host.)
        const text = stack(new Set(allFaces.filter((face) => face !== georgia)));
        const [editor, canvas] = editorOn(text);
        Word.importFromDocxPackage(editor, pkg);
        const wanted = [];
        Word.onFontsDiscovered(pkg, (faces) => wanted.push(...faces.map(keyOf)));
        const changes = [];
        Word.onChange(editor, (change) => changes.push({ ...change }));
        const pages = Word.getPageCount(editor);
        check('the page count is that of the faces there are, and the editor says so', [pages, Word.getPendingFonts(editor).map(keyOf), Word.getDocumentSize(editor).FontsPending, wanted.includes(georgia)], [provisionalPages, [georgia], true, true]);
        check('the notices of the load say that a face is pending', [changes.length > 0, changes.every((change) => change.FontsPending)], [true, true]);
        check('whenLayoutComplete says it too', [(await Word.whenLayoutComplete(editor)).FontsPending], [true]);
        // The face arrives: the host installs it; the listeners are told that the layout changed.
        changes.length = 0;
        for (const face of wanted) text.faces.add(face);
        Word.setTextMeasurers(editor, text.measureText, text.measureRange);
        check('installing the face tells the listeners that the layout changed, nothing pending', [changes.length >= 1, changes.some((change) => change.LayoutChanged && !change.FontsPending && !change.DocumentChanged), changes.some((change) => change.FontsPending)], [true, true, false]);
        check('the page count is the final one', [Word.getPageCount(editor), Word.getPendingFonts(editor), Word.getDocumentSize(editor).FontsPending], [finalPages, [], false]);
        Word.dispose(editor); canvas.remove();
    }

    // ---- read in steps: the faces before the blocks --------------------------------
    {
        const pkg = await OF.word.openProgressive(long, { autoStart: false });
        const text = stack(new Set(Word.getRequiredFonts(pkg).map(keyOf)));
        const [editor, canvas] = editorOn(text);
        Word.importFromDocxPackage(editor, pkg);
        const calls = [];
        Word.onFontsDiscovered(pkg, (faces) => calls.push({ faces: faces.map(keyOf), behind: pkg.Paragraphs.length - Word.getRenderStats(editor).Blocks, pendingThen: Word.getPendingFonts(editor).map(keyOf) }));
        let pendingSeen = false;
        while (!pkg.Progressive.IsComplete) { pkg.Progressive.readMore(5); if (!pendingSeen && Word.getPendingFonts(editor).length > 0) pendingSeen = calls.some((call) => call.faces.includes(georgia)); }
        check('in every step the faces are told before the editor has the blocks', [calls.length >= 1, calls.every((call) => call.behind > 0), calls.every((call) => !call.faces.some((face) => call.pendingThen.includes(face)))], [true, true, true]);
        check('a face is pending from the step that found it', [pendingSeen, Word.getPendingFonts(editor).map(keyOf).includes(georgia), Word.getDocumentSize(editor).FontsPending], [true, true, true]);
        Word.dispose(editor); canvas.remove();
    }

    // ---- the viewer -----------------------------------------------------------------
    {
        const pkg = await OF.word.openProgressive(long, { autoStart: false });
        const text = stack(new Set(allFaces.filter((face) => face !== georgia)));
        const viewer = Word.createViewer(document.createElement('canvas'));
        viewer.setTextMeasurers(text.measureText, text.measureRange);
        viewer.setFontAvailability(text.isAvailable);
        viewer.load(pkg, { progressive: true });
        const progress = [];
        viewer.onLayoutProgress((state) => progress.push({ ...state }));
        const calls = [];
        Word.onFontsDiscovered(pkg, (faces) => calls.push(viewer.getPageFrames().length));
        check('the viewer: nothing pending at open, and asking reads nothing', [viewer.getPendingFonts(), viewer.getDocumentSize().FontsPending, viewer.isLoadComplete()], [[], false, false]);
        const pages = viewer.getPageCount();
        check('its page count at open is that of the faces there are, and it says so', [pages, viewer.getPendingFonts().map(keyOf), viewer.getDocumentSize().FontsPending, calls.length], [provisionalPages, [georgia], true, 1]);
        check('its notices say it', [progress.length > 0 && progress[progress.length - 1].FontsPending], [true]);
        text.faces.add(georgia);
        viewer.setTextMeasurers(text.measureText, text.measureRange);
        check('with the face installed the pages are the final ones', [viewer.getPageCount(), viewer.getPendingFonts(), viewer.getDocumentSize().FontsPending], [finalPages, [], false]);
        viewer.dispose();
    }
} catch (e) { fail(e); }
finish();
