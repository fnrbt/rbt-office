// getRequiredFonts: the font faces a document's text is set in, after style
// resolution and with the faces the engine itself introduces (list markers in
// Symbol / Wingdings / Courier New, notes) - from an editor package before
// anything is shown, from an editor, and from a viewer (named export of
// word/WordViewer.js, which stays free of the editor).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const viewerModule = await import('/word/WordViewer.js');
    const key = (faces) => faces.map((face) => `${face.family}${face.bold ? ' bold' : ''}${face.italic ? ' italic' : ''}`);
    for (const [name, url] of [['docx4j-getting-started', '/repo/tests/test-documents-wild/docx4j-getting-started.docx'], ['synthetic/lists', '/fixture?name=' + encodeURIComponent('synthetic/lists.docx')], ['synthetic/run-props', '/fixture?name=' + encodeURIComponent('synthetic/run-props.docx')]]) {
        const editorPackage = await OF.processors.forEditorPackage(await bytesOf(url));
        const t0 = performance.now();
        const fromPackage = Word.getRequiredFonts(editorPackage);
        const elapsed = performance.now() - t0;
        check(`${name}: an array of { family, bold, italic }`, [Array.isArray(fromPackage), fromPackage.length > 0, fromPackage.every((face) => typeof face.family === 'string' && face.family !== '' && typeof face.bold === 'boolean' && typeof face.italic === 'boolean')], [true, true, true]);
        check(`${name}: each face once`, new Set(key(fromPackage)).size, fromPackage.length);
        // The editor, with the document shown, and the viewer: the same faces.
        const canvas = document.createElement('canvas');
        document.body.appendChild(canvas);
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, editorPackage);
        check(`${name}: the same from the editor`, key(Word.getRequiredFonts(editor)), key(fromPackage));
        const viewer = viewerModule.default(document.createElement('canvas'));
        viewer.load(editorPackage);
        check(`${name}: the same from the viewer, and from the viewer module's function`, [key(viewer.getRequiredFonts()), key(viewerModule.getRequiredFonts(editorPackage))], [key(fromPackage), key(fromPackage)]);
        // Every face the painter is asked to draw with is in the list.
        const drawn = new Set();
        Word.setTextPathPainter(editor, (text, style) => { drawn.add(`${style.FontFamily ?? 'Calibri'}${style.Bold ? ' bold' : ''}${style.Italic ? ' italic' : ''}`); return false; });
        const size = Word.getDocumentSize(editor);
        for (let top = 0; top < size.Height; top += 700) Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
        const listed = new Set(key(fromPackage));
        check(`${name}: every face text is painted in is listed (${drawn.size} painted)`, [...drawn].filter((face) => !listed.has(face)), []);
        // ... and measured in.
        const measured = new Set();
        const seen = (style) => measured.add(`${style.FontFamily ?? 'Calibri'}${style.Bold ? ' bold' : ''}${style.Italic ? ' italic' : ''}`);
        Word.setTextMeasurers(editor, (text, style) => { seen(style); return text.length * style.FontSize * 0.5; }, (text, style, start, length) => { seen(style); return Math.max(0, Math.min(length, text.length - start)) * style.FontSize * 0.5; });
        Word.getPageCount(editor);
        check(`${name}: every face text is measured in is listed (${measured.size} measured)`, [...measured].filter((face) => !listed.has(face)), []);
        note(`${name}: ${key(fromPackage).join(', ')} (${elapsed.toFixed(0)} ms from the package)`);
        Word.dispose(editor); canvas.remove();
    }
    const lists = key(Word.getRequiredFonts(await OF.processors.forEditorPackage(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/lists.docx')))));
    check('a bulleted list brings the face of its markers', lists.some((face) => /^(Symbol|Wingdings|Courier New)/.test(face)), true);
    check('nothing to list for nothing', [Word.getRequiredFonts(null).length], [0]);
} catch (e) { fail(e); }
finish();
