// The fonts a document embeds (word/fonts/*.odttf, de-obfuscated by Office.Formats): getEmbeddedFonts
// of a package, an editor and a viewer; loadEmbeddedFonts makes them canvas fonts - for the families
// the page has no font of, as Word uses them only when the machine has none (fixtures embedded-fonts,
// docs/layout-parity-flow.md) - and lays their text out again; an embedded face loses to the page's
// own font of its family; a face of restricted licence is not given; the PDF export embeds them; a
// save after an edit keeps the font parts and the font table as they were.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const viewerModule = await import('/word/WordViewer.js');
    const fixture = (name) => '/repo/tests/word/fixtures/embedded-fonts/' + name + '.docx';
    const key = (faces) => faces.map((face) => `${face.family}${face.bold ? ' bold' : ''}${face.italic ? ' italic' : ''}`);
    const familyFaces = (family) => [...document.fonts].filter((face) => face.family.replace(/"/g, '') === family);

    const calibreBytes = await bytesOf('/repo/tests/test-documents/calibre-demo.docx');
    const calibre = await OF.processors.forEditorPackage(calibreBytes);
    const fromPackage = Word.getEmbeddedFonts(calibre);
    check('calibre-demo: the faces it embeds, from the package', key(fromPackage), ['Ubuntu', 'Ubuntu bold', 'Ubuntu italic', 'Ubuntu bold italic', 'Tahoma', 'Ubuntu Mono']);
    check('calibre-demo: each a font file', fromPackage.map((face) => face.data instanceof Uint8Array && face.data.length > 1000 && face.data[0] === 0 && face.data[1] === 1), fromPackage.map(() => true));

    // The editor: the faces become canvas fonts, and the text in them is laid out again.
    const canvas = document.createElement('canvas');
    document.body.appendChild(canvas);
    const editor = Word.initializeCanvas(canvas);
    let layoutChanged = 0;
    Word.onChange(editor, (change) => { if (change.LayoutChanged) layoutChanged++; });
    Word.importFromDocxPackage(editor, calibre);
    check('the editor: the same faces', key(Word.getEmbeddedFonts(editor)), key(fromPackage));
    const loaded = await Word.loadEmbeddedFonts(editor);
    check('loadEmbeddedFonts: resolves to the faces', key(loaded), key(fromPackage));
    check('Ubuntu: four faces in document.fonts, loaded', familyFaces('Ubuntu').map((face) => face.status), ['loaded', 'loaded', 'loaded', 'loaded']);
    check('Ubuntu is a canvas font now', document.fonts.check('16px "Ubuntu"') && (() => {
        const context = document.createElement('canvas').getContext('2d');
        context.font = '72px monospace'; const alone = context.measureText('mmmmmlli').width;
        context.font = '72px "Ubuntu", monospace'; return context.measureText('mmmmmlli').width !== alone;
    })(), true);
    check('the text in it laid out again (LayoutChanged)', layoutChanged > 0, true);
    check('asked again: the same, nothing laid out again', [key(await Word.loadEmbeddedFonts(editor)), layoutChanged], [key(loaded), layoutChanged]);

    // A family the page has a font of keeps it: the "Calibri" a fixture embeds (Liberation Mono renamed),
    // once the page has a Calibri of its own.
    const own = new FontFace('Calibri', await bytesOf('/fonts/Carlito-Regular.ttf'));
    await own.load(); document.fonts.add(own);
    const calibriEditor = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(calibriEditor, await OF.processors.forEditorPackage(await bytesOf(fixture('calibri'))));
    check('calibri: the embedded faces are given', Word.getEmbeddedFonts(calibriEditor).length, 4);
    await Word.loadEmbeddedFonts(calibriEditor);
    check('calibri: the page\'s own Calibri stays the only one', familyFaces('Calibri').length, 1);
    Word.dispose(calibriEditor);
    // A face of restricted licence (fsType 2) is not given; "Rbt Embed Mono" of reg is.
    check('restricted: no face', Word.getEmbeddedFonts(await OF.processors.forEditorPackage(await bytesOf(fixture('restricted')))).length, 0);
    const regEditor = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(regEditor, await OF.processors.forEditorPackage(await bytesOf(fixture('reg'))));
    await Word.loadEmbeddedFonts(regEditor);
    check('reg: "Rbt Embed Mono" a canvas font', familyFaces('Rbt Embed Mono').length, 1);
    // (Every character of it is as wide as every other: the layout measured it so.)
    const segments = Word.getLayoutLines(regEditor)[0].Segments.map((segment) => Math.round(segment.Width * 100) / 100);
    check('reg: "MMMMMMMMMM iiiiiiiiii Quiz" measured in it (26 x 6pt = 208px)', Math.round(segments.reduce((a, b) => a + b, 0)), 208);
    Word.dispose(regEditor);

    // The viewer.
    const viewer = viewerModule.default(document.createElement('canvas'));
    viewer.load(calibre);
    check('the viewer: the same faces', key(viewer.getEmbeddedFonts()), key(fromPackage));
    await viewer.loadEmbeddedFonts();

    // The PDF export embeds them (no host fonts given).
    const pdfModule = await import('/word-pdf/Browser.js');
    // (The editor's whole document, as the shell's print takes it.)
    const blob = await pdfModule.exportPdf(editor.Whole, [], false);
    const pdf = new TextDecoder('latin1').decode(new Uint8Array(await blob.arrayBuffer()));
    check('the PDF embeds Ubuntu', /Embedded-Ubuntu/.test(pdf), true);

    // Saved after an edit: the font parts and the table as they were.
    const first = Word.getLayoutLines(editor)[0];
    Word.setSelection(editor, first.BlockId, 0, first.BlockId, 0);
    Word.insertText(editor, 'Edited ');
    const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const before = await window.JSZip.loadAsync(calibreBytes), after = await window.JSZip.loadAsync(out);
    const fontParts = Object.keys(before.files).filter((name) => name.startsWith('word/fonts/') || name.includes('fontTable')).sort();
    const same = [];
    for (const name of fontParts) {
        const a = await before.file(name).async('uint8array'), b = await after.file(name)?.async('uint8array');
        same.push(!!b && a.length === b.length && a.every((v, i) => v === b[i]));
    }
    check(`saved after an edit: ${fontParts.length} font parts as they were`, same, fontParts.map(() => true));
    check('saved after an edit: the faces read again', key(Word.getEmbeddedFonts(await OF.processors.forEditorPackage(out))), key(fromPackage));
    Word.dispose(editor); canvas.remove();
} catch (e) { fail(e); }
finish();
