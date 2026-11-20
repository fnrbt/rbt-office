// The canvas's own measuring (no host measurers): a word is the sum of its characters' own
// advances, unkerned and without ligatures, as Word sets text without w:kern / w14:ligatures -
// against the widths Word 2019 gives the kerning fixtures' words in its PDFs
// (tests/word/fixtures/kerning/word/widths.tsv; tests/word/kerning.fsx checks ShaperText's).
// Measured at a reference size and scaled: not the whole pixels Chromium's canvas gives at the
// text's own size ("AVAVAVAVAV" in Calibri 11pt: Word 63.04pt; the canvas kerned and in whole
// pixels 60.75pt). The faces: dist/fonts' Carlito, Liberation Sans and Liberation Serif under the
// Office families' names (their advances are Calibri's, Arial's and Times New Roman's; Caladea's
// are not all Cambria's, so Cambria is left out, as kerning.fsx does).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    for (const [family, file] of [['Calibri', 'Carlito-Regular.ttf'], ['Arial', 'LiberationSans-Regular.ttf'], ['Times New Roman', 'LiberationSerif-Regular.ttf']]) {
        const face = new FontFace(family, await bytesOf('/fonts/' + file));
        await face.load();
        document.fonts.add(face);
    }
    const rows = (await (await fetch('/repo/tests/word/fixtures/kerning/word/widths.tsv')).text())
        .split('\n').filter((line) => line.length > 0).map((line) => line.split('\t'))
        .map(([fixture, font, size, word, width]) => ({ fixture, font, size: Number(size), word, width: Number(width) }));
    // (Word's widths are its PDF's: glyph positions in 1/600in.)
    const tolerance = 0.25;
    for (const fixture of ['kern-none', 'kern-none-16', 'kern-28', 'kern-none-c14']) {
        const editor = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/word/fixtures/kerning/' + fixture + '.docx')));
        Word.completeLayout(editor);
        const lines = Word.getLayoutLines(editor).filter((line) => !String(line.BlockId).startsWith('__'));
        const want = rows.filter((row) => row.fixture === fixture && row.font !== 'Cambria');
        // (The fixture's paragraphs are its words in order, Cambria's last.)
        const got = want.map((row, i) => {
            const line = lines[i];
            const width = line ? line.Segments.reduce((sum, segment) => sum + segment.Width, 0) * 0.75 : NaN;
            return { row, text: line ? line.Text.trim() : '', width };
        });
        const off = got.filter(({ row, text, width }) => text !== row.word || !(Math.abs(width - row.width) <= tolerance));
        check(`${fixture}: ${want.length} words at Word's widths (within ${tolerance}pt)`,
            off.map(({ row, text, width }) => `${row.font} ${row.word}: ${text} ${width.toFixed(3)}`), []);
        if (fixture === 'kern-none') {
            const av = got.find(({ row }) => row.font === 'Calibri' && row.word === 'AVAVAVAVAV');
            note(`Calibri 11pt AVAVAVAVAV ${av.width.toFixed(3)}pt, Word ${av.row.width}`);
        }
    }
} catch (e) { fail(e); }
finish();
