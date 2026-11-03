// Flipped shapes and pictures in pixels: a right triangle flipped across
// (a:xfrm flipH) has its right angle at the right, a picture flipped across
// shows its right half at the left, one flipped upside down its bottom half
// at the top (the fixture synthetic/flips, its shape not turned, a picture of
// two colours instead of its one pixel); getShapeRects says how a shape is
// flipped. (Flips were not read: both were drawn as if unflipped.)
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const context = canvas.getContext('2d', { willReadFrequently: true });
    // A picture of 2 x 2 pixels: red at the top left, blue at the top right, green below.
    const source = new OffscreenCanvas(2, 2);
    const sourceContext = source.getContext('2d');
    sourceContext.fillStyle = '#FF0000'; sourceContext.fillRect(0, 0, 1, 1);
    sourceContext.fillStyle = '#0000FF'; sourceContext.fillRect(1, 0, 1, 1);
    sourceContext.fillStyle = '#00FF00'; sourceContext.fillRect(0, 1, 2, 1);
    const png = new Uint8Array(await (await source.convertToBlob({ type: 'image/png' })).arrayBuffer());
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/flips.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const open = async (pictureFlip) => {
        // (The shape not turned; the picture 192 x 96 px with the flips given.)
        // (The shape at the right of the column, clear of the text and the picture.)
        const xml = original.replace('<a:xfrm flipH="1" rot="1800000">', '<a:xfrm flipH="1">').replace('<wp:posOffset>95250</wp:posOffset></wp:positionH>', '<wp:posOffset>4500000</wp:posOffset></wp:positionH>')
            .replace('<a:xfrm flipV="1" flipH="1">', `<a:xfrm${pictureFlip}>`)
            .split('cx="457200" cy="228600"').join(`cx="${192 * 9525}" cy="${96 * 9525}"`);
        zip.file('word/document.xml', xml);
        zip.file('word/media/body.png', png);
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        // (The picture decoded before it is painted.)
        const src = Word.exportDocumentPackage(editor).Paragraphs.flatMap((p) => p.Runs ?? []).find((r) => r.IsImage).ImageSrc;
        const image = new Image(); image.src = src; await image.decode();
        Word.setImageCache(editor, new Map([[src, image]]));
        for (let i = 0; i < 40 && colour(pictureOf(editor).X + 4, pictureOf(editor).Y + 4) === 'white'; i++) await new Promise((resolve) => setTimeout(resolve, 50));
        return editor;
    };
    const colour = (x, y) => { const d = context.getImageData(Math.round(x), Math.round(y), 1, 1).data; return d[0] > 200 && d[1] < 80 && d[2] < 80 ? 'red' : d[2] > 200 && d[0] < 80 && d[1] < 80 ? 'blue' : d[1] > 200 && d[0] < 80 && d[2] < 80 ? 'green' : d[0] > 200 && d[1] > 90 && d[1] < 160 && d[2] < 90 ? 'orange' : d[0] > 240 && d[1] > 240 && d[2] > 240 ? 'white' : `rgb(${d[0]},${d[1]},${d[2]})`; };
    const pictureOf = (editor) => {
        const left = Word.getPageFrames(editor)[0].X ?? 0;
        for (const line of Word.getLayoutLines(editor)) for (const segment of line.Segments ?? []) if (segment.Kind === 'image') return { X: left + segment.X, Y: segment.Y ?? line.Y, Width: segment.Width, Height: segment.Height };
        return null;
    };
    const quarters = (editor) => {
        const p = pictureOf(editor);
        return [colour(p.X + p.Width * 0.25, p.Y + p.Height * 0.25), colour(p.X + p.Width * 0.75, p.Y + p.Height * 0.25), colour(p.X + p.Width * 0.5, p.Y + p.Height * 0.75)];
    };

    const plain = await open('');
    const picture = pictureOf(plain);
    note(`picture at ${picture.X.toFixed(1)}, ${picture.Y.toFixed(1)}: ${picture.Width} x ${picture.Height}`);
    check('a picture not flipped: red at the top left, blue at the top right, green below', quarters(plain), ['red', 'blue', 'green']);
    // The triangle: its right angle at the bottom right once flipped across - filled at the right of its lower half, not at the left.
    const shape = Word.getShapeRects(plain).find((rect) => rect.Text === 'upright text');
    const triangle = { ...shape, X: shape.X + (Word.getPageFrames(plain)[0].X ?? 0) };
    check('getShapeRects says how the shape is flipped', [triangle.FlipH, triangle.FlipV, triangle.Rotation], [true, false, 0]);
    check('the triangle flipped across: its right angle at the right',
        [colour(triangle.X + triangle.Width * 0.9, triangle.Y + triangle.Height * 0.8), colour(triangle.X + triangle.Width * 0.1, triangle.Y + triangle.Height * 0.8)], ['orange', 'white']);
    const across = await open(' flipH="1"');
    check('a picture flipped across: blue at the top left, red at the top right', quarters(across), ['blue', 'red', 'green']);
    const upsideDown = await open(' flipV="1"');
    check('a picture flipped upside down: green at the top, red and blue below', [colour(pictureOf(upsideDown).X + 48, pictureOf(upsideDown).Y + 24), colour(pictureOf(upsideDown).X + 48, pictureOf(upsideDown).Y + 72), colour(pictureOf(upsideDown).X + 144, pictureOf(upsideDown).Y + 72)], ['green', 'red', 'blue']);
    const both = await open(' flipH="1" flipV="1"');
    check('a picture flipped both ways: green at the top, blue and red below', [colour(pictureOf(both).X + 96, pictureOf(both).Y + 24), colour(pictureOf(both).X + 48, pictureOf(both).Y + 72), colour(pictureOf(both).X + 144, pictureOf(both).Y + 72)], ['green', 'blue', 'red']);
    // The flips are saved.
    const saved = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(both))).arrayBuffer());
    const savedXml = await (await window.JSZip.loadAsync(saved)).file('word/document.xml').async('string');
    check('saved with the flips', [savedXml.includes('<a:xfrm flipH="1" flipV="1">'), savedXml.includes('<a:xfrm flipH="1">')], [true, true]);
} catch (e) { fail(e); }
finish();
