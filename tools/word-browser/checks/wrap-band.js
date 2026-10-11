// Text beside a floating picture is not put into a band narrower than an
// inch, and stands on the side(s) the picture says (wrapText).
// float-image-wrap-square.docx: the picture is anchored 1463040 EMU (153.6 px)
// from the page's left edge and 1371600 EMU (144 px) from its top, is
// 1463040 x 1097280 EMU (153.6 x 115.2 px), wrapSquare wrapText="bothSides",
// distL = distR = 114300 EMU (12 px), distT = distB = 0; the margins are 96 px.
// Left of the picture there are 153.6 - 96 - 12 = 45.6 px: no text goes there
// (single words stood in that sliver, and the heading above was split by it).
// The lines that pass the picture (144 .. 259.2 px down the page) start
// 153.6 + 153.6 + 12 = 319.2 px from the page's left edge.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const bytes = await bytesOf('/repo/tests/repros/word/float-image-wrap-square.docx');
    const open = async (data) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data)); return editor; };
    // The lines with text, each with where its text starts and ends across the page.
    const textLines = (editor) => Word.getLayoutLines(editor).map((line) => {
        // (Text: not the picture itself, which a line holds as a character of its own.)
        const parts = (line.Segments ?? []).filter((segment) => typeof segment.Text === 'string' && segment.Text.replace(/\uFFFC/g, '').trim() !== '');
        return parts.length === 0 ? null : { text: line.Text, y: line.Y, height: line.Height, left: Math.min(...parts.map((segment) => segment.X)), right: Math.max(...parts.map((segment) => segment.X + segment.Width)), parts: parts.map((segment) => [segment.X, segment.X + segment.Width - (segment.Text.length - segment.Text.trimEnd().length) * 3]) };
    }).filter((line) => line !== null);
    const beside = (lines) => lines.filter((line) => line.y + line.height > 144.5 && line.y < 258.7);
    const editor = await open(bytes);
    const lines = textLines(editor);
    const passing = beside(lines), below = lines.filter((line) => line.y >= 259.2);
    note('lines: ' + JSON.stringify(lines.map((line) => [Math.round(line.left), Math.round(line.right), Math.round(line.y), line.text.slice(0, 18)])));
    check('lines pass the picture', passing.length >= 3, true);
    check('none of them has text left of the picture', passing.filter((line) => line.left < 319.2 - 0.6).map((line) => [Math.round(line.left), line.text.slice(0, 20)]), []);
    check('they start 12 px right of it', [...new Set(passing.map((line) => Math.round(line.left * 10) / 10))], [319.2]);
    check('below the picture the lines start at the margin', [below.length >= 2, [...new Set(below.map((line) => Math.round(line.left)))]], [true, [96]]);
    const heading = lines.find((line) => line.text.startsWith('Float wrap-square repro'));
    check('the heading is one line in one piece', [lines.filter((line) => line.text.includes('wrap-square repro') || line.text.startsWith('Float')).length, heading !== undefined && heading.text.trim()], [1, 'Float wrap-square repro']);
    // Painted: nothing in the sliver left of the picture while lines pass it.
    const context = canvas.getContext('2d');
    const ink = (x0, x1, y0, y1) => { const data = context.getImageData(Math.round(x0), Math.round(y0), Math.round(x1 - x0), Math.round(y1 - y0)).data; let count = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 110) count++; return count; };
    check('painted: no text in the 45.6 px left of the picture, text right of it', [ink(96, 141, 146, 258), ink(320, 720, 146, 258) > 200], [0, true]);

    // The side the text may stand on: the picture moved 2 in to the right (237.6 px of text left of it, 208.8 px right of it).
    const zip = await window.JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml').async('string');
    const anchor = '<wp:positionH relativeFrom="page"><wp:posOffset>1463040</wp:posOffset></wp:positionH>', wrap = '<wp:wrapSquare wrapText="bothSides"/>';
    if (!xml.includes(anchor) || !xml.includes(wrap)) throw new Error('the document changed');
    const sides = async (side) => {
        const copy = await window.JSZip.loadAsync(bytes);
        copy.file('word/document.xml', xml.replace(anchor, anchor.replace('1463040', String(1463040 + 1828800))).replace(wrap, wrap.replace('bothSides', side)));
        const moved = beside(textLines(await open(await copy.generateAsync({ type: 'uint8array' }))));
        // (The picture: 345.6 .. 499.2 px from the page's left edge; with its distances 333.6 .. 511.2.)
        // (A line may have text on both sides of the picture: no piece of it is over the picture.)
        return [moved.some((line) => line.left < 333.6), moved.some((line) => line.right > 520), moved.every((line) => line.parts.every(([from, to]) => to <= 334.2 || from >= 510.6))];
    };
    check('both sides: text left and right of the picture, none over it', await sides('bothSides'), [true, true, true]);
    check('left: text left of it only', await sides('left'), [true, false, true]);
    check('right: text right of it only', await sides('right'), [false, true, true]);
    check('largest: the wider side, the left one', await sides('largest'), [true, false, true]);
} catch (e) { fail(e); }
finish();
