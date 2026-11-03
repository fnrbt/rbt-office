// Shapes that may not overlap (wp:anchor allowOverlap="0"): the fixture
// synthetic/no-overlap has two in one paragraph, the second placed over the
// first (100 px right, 30 px down). Word moves the second down below the
// first: getShapeRects says it is there, it is painted there (pixels), and
// the text goes around it there. Either shape
// allowed to overlap: the second stays over the first. (The flag was read
// and not used.)
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d', { willReadFrequently: true });
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/no-overlap.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const open = async (xml) => {
        zip.file('word/document.xml', xml);
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        return editor;
    };
    const near = (a, b) => Math.abs(a - b) < 1.5;
    const pixel = (x, y) => Array.from(context.getImageData(Math.round(x), Math.round(y), 1, 1).data.slice(0, 3)).map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    const shapesOf = (editor) => {
        const left = Word.getPageFrames(editor)[0].X ?? 0;
        const rects = Word.getShapeRects(editor);
        const at = (text) => { const rect = rects.find((r) => r.Text === text); return { ...rect, X: rect.X + left }; };
        return [at('first'), at('second')];
    };

    const editor = await open(original);
    const [first, second] = shapesOf(editor);
    note(`first at ${first.X.toFixed(1)}, ${first.Y.toFixed(1)}; second at ${second.X.toFixed(1)}, ${second.Y.toFixed(1)} (${second.Width} x ${second.Height})`);
    check('the second is moved down below the first, across where it was', [near(second.Y, first.Y + first.Height), near(second.X, first.X + 100)], [true, true]);
    check('painted there: the second\'s colour below the first, the first\'s where the second was placed', [pixel(second.X + 150, second.Y + 50), pixel(first.X + 150, first.Y + 50)], ['FEDCBA', 'ABCDEF']);
    // The lines beside the second (below the first) stand clear of it.
    const lines = Word.getLayoutLines(editor).filter((line) => /^\uFFFC*w\d/.test(line.Text ?? ''));
    // (Lines and their segments, as getShapeRects, from the left edge of the page.)
    const pageLeft = Word.getPageFrames(editor)[0].X ?? 0;
    const shapeLeft = second.X - pageLeft, shapeRight = shapeLeft + second.Width;
    const beside = lines.filter((line) => line.Y + line.Height > second.Y + 4 && line.Y < second.Y + second.Height - 4);
    const clear = beside.every((line) => (line.Segments ?? []).filter((seg) => typeof seg.Text === 'string' && seg.Text.replace(/\uFFFC/g, '').trim() !== '').every((seg) => seg.X + seg.Width <= shapeLeft + 0.5 || seg.X >= shapeRight - 0.5));
    note(`${beside.length} lines beside it, starting at ${beside.map((line) => Math.round(line.X)).join(', ')}`);
    check('the text beside it goes around it', [beside.length > 1, clear], [true, true]);
    // Allowed to overlap: as placed.
    for (const which of ['First', 'Second']) {
        const marker = `name="${which}"`;
        const at = original.indexOf(marker);
        const anchorStart = original.lastIndexOf('<wp:anchor ', at);
        const xml = original.slice(0, anchorStart) + original.slice(anchorStart, at).replace('allowOverlap="0"', 'allowOverlap="1"') + original.slice(at);
        const [f, s] = shapesOf(await open(xml));
        check(`the ${which.toLowerCase()} allowed to overlap: the second stays over the first`, [near(s.Y, f.Y + 30), near(s.X, f.X + 100)], [true, true]);
    }
} catch (e) { fail(e); }
finish();
