// Table borders in pixels (whole canvas, canvas text): a double border is two
// lines with a clear gap that runs round the corners, a thin-thick border has
// its thin line where the style says (first named = top / left), where two
// cells state different borders for the edge they share the wider one is
// drawn (Word's rule) in one colour, and the corner of a thick single border
// is filled. The tables are made here (synthetic/tables with another body).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/tables.docx')));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, sect = xml.lastIndexOf('<w:sectPr');
    // (Lines of exactly 24px and no spacing: the tables stand at whole pixels, whatever single spacing comes to -
    // a border's pixels are counted below.)
    const p = (text) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const borders = (val, sz, color = '000000') => ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`).join('');
    const cell = (text, tcBorders = '') => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>${tcBorders ? `<w:tcBorders>${tcBorders}</w:tcBorders>` : ''}</w:tcPr>${p(text)}</w:tc>`;
    const table = (tblBorders, rows) => `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/><w:tblBorders>${tblBorders}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${rows.map((row) => `<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="exact"/></w:trPr>${row.join('')}</w:tr>`).join('')}</w:tbl>`;
    const body =
        p('double') + table(borders('double', 12), [[cell('a'), cell('b')], [cell('c'), cell('d')]]) +
        p('thin-thick') + table(['top', 'left', 'bottom', 'right'].map((side) => `<w:${side} w:val="thinThickSmallGap" w:sz="24" w:space="0" w:color="000000"/>`).join(''), [[cell('a'), cell('b')]]) +
        p('conflict') + table(borders('single', 4, '0000FF'), [[cell('a', '<w:right w:val="single" w:sz="24" w:space="0" w:color="FF0000"/>'), cell('b')], [cell('c'), cell('d', '<w:top w:val="single" w:sz="36" w:space="0" w:color="00AA00"/>')]]) +
        p('thick') + table(borders('single', 36), [[cell('a'), cell('b')]]) + p('end');
    zip.file('word/document.xml', xml.slice(0, open) + body + xml.slice(sect));
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
    const context = document.getElementById('c').getContext('2d');
    const at = (x, y) => { const d = context.getImageData(Math.round(x), Math.round(y), 1, 1).data; return [d[0], d[1], d[2]]; };
    const dark = ([r, g, b]) => r < 140 && g < 140 && b < 140;
    // (A hairline is antialiased: its colour, however pale.)
    const red = ([r, g, b]) => r > 150 && r - g > 50 && r - b > 50, blue = ([r, g, b]) => b > 150 && b - r > 40 && b - g > 40, green = ([r, g, b]) => g > 110 && g - r > 40 && g - b > 40;
    // Runs of pixels that satisfy `is` along a row (or a column): [[start, length], ...].
    const runs = (fixed, from, to, horizontal, is) => { const found = []; let start = -1; for (let v = from; v <= to; v++) { const hit = is(horizontal ? at(v, fixed) : at(fixed, v)); if (hit && start < 0) start = v; if (!hit && start >= 0) { found.push([start, v - start]); start = -1; } } if (start >= 0) found.push([start, to - start + 1]); return found; };
    // The tables' boxes from the pixels: the paragraphs between them say where each starts.
    const lines = Word.getLayoutLines(editor);
    const yOf = (text) => lines.find((line) => (line.Text ?? '') === text).Y;
    const box = (label, next) => {
        // Down a column in the second cell, right of its text: the first and the last row with ink are the
        // table's top and bottom; along a row a quarter down: the first and the last column with ink.
        const from = Math.round(yOf(label) + 18), to = Math.round(yOf(next) - 1);
        const inked = ([r, g, b]) => r < 200 || g < 200 || b < 200;
        let top = -1, bottom = -1;
        for (let y = from; y <= to; y++) if (inked(at(400, y))) { if (top < 0) top = y; bottom = y; }
        const row = Math.round(top + (bottom - top) * 0.25);
        let left = -1, right = -1;
        for (let x = 60; x <= 520; x++) if (inked(at(x, row))) { if (left < 0) left = x; right = x; }
        return { left, right, top, bottom };
    };
    const double = box('double', 'thin-thick'), thinThick = box('thin-thick', 'conflict'), conflict = box('conflict', 'thick'), thick = box('thick', 'end');
    note(`boxes ${JSON.stringify({ double, thinThick, conflict, thick })}`);
    const map = (x0, y0, w, h) => { const rows = []; for (let y = y0; y < y0 + h; y++) { let row = ''; for (let x = x0; x < x0 + w; x++) { const v = at(x, y)[0]; row += v < 90 ? '#' : v < 200 ? '+' : '.'; } rows.push(row); } return rows.join(' / '); };
    note(`double, top left ${map(double.left - 2, double.top - 2, 12, 12)}`);
    note(`double, top right ${map(double.right - 9, double.top - 2, 12, 12)}`);
    note(`double, middle ${map(Math.round((double.left + double.right) / 2) - 6, Math.round((double.top + double.bottom) / 2) - 6, 12, 12)}`);
    note(`double, top middle ${map(Math.round((double.left + double.right) / 2) - 6, double.top - 2, 12, 12)}`);
    const midY = (b, part = 0.25) => Math.round(b.top + (b.bottom - b.top) * part), midX = (b, part = 0.25) => Math.round(b.left + (b.right - b.left) * part);

    // ---- double ----
    // (Across a row below its text: the text stands below the row's top border, as in Word.)
    const leftRuns = runs(midY(double, 0.4), double.left - 3, double.left + 12, true, dark);
    const topRuns = runs(midX(double), double.top - 3, double.top + 12, false, dark);
    const rightRuns = runs(midY(double, 0.4), double.right - 12, double.right + 3, true, dark);
    const bottomRuns = runs(midX(double), double.bottom - 12, double.bottom + 3, false, dark);
    check('a double border is two lines, on every side', [leftRuns.length, topRuns.length, rightRuns.length, bottomRuns.length], [2, 2, 2, 2]);
    check('the lines of a side are as wide as each other', [leftRuns, topRuns, rightRuns, bottomRuns].map((side) => side.length === 2 && Math.abs(side[0][1] - side[1][1]) <= 1), [true, true, true, true]);
    if (leftRuns.length === 2 && topRuns.length === 2) {
        const gapX = leftRuns[0][0] + leftRuns[0][1], gapY = topRuns[0][0] + topRuns[0][1];
        const gapWidth = leftRuns[1][0] - gapX, gapHeight = topRuns[1][0] - gapY;
        // Round the top-left corner: the gap is clear all the way, the outer and the inner line are closed.
        const clear = [];
        for (let x = gapX; x < gapX + 30; x++) clear.push(dark(at(x, gapY + Math.floor(gapHeight / 2))));
        for (let y = gapY; y < gapY + 30; y++) clear.push(dark(at(gapX + Math.floor(gapWidth / 2), y)));
        check('the gap of a double border runs round the corner (no line crosses it)', clear.filter((hit) => hit).length, 0);
        check('the outer line is closed at the corner, and so is the inner one', [dark(at(leftRuns[0][0], topRuns[0][0])), dark(at(leftRuns[1][0], topRuns[1][0]))], [true, true]);
        check('nothing sticks out past the outer corner', [dark(at(leftRuns[0][0] - 2, topRuns[0][0])), dark(at(leftRuns[0][0], topRuns[0][0] - 2))], [false, false]);
    }
    const insideV = runs(midY(double, 0.4), midX(double, 0.5) - 12, midX(double, 0.5) + 12, true, dark);
    check('the line between two cells is double too', insideV.length, 2);
    const insideH = runs(midX(double), midY(double, 0.5) - 12, midY(double, 0.5) + 12, false, dark);
    if (insideV.length === 2 && insideH.length === 2) {
        // Where the two inside borders cross, and where the inside one meets the top border: the gaps
        // run into each other (each line turns the corner into the line it meets; none crosses a gap).
        const gapX = insideV[0][0] + insideV[0][1], gapY = insideH[0][0] + insideH[0][1];
        const gapMidX = gapX + Math.floor((insideV[1][0] - gapX) / 2), gapMidY = gapY + Math.floor((insideH[1][0] - gapY) / 2);
        const crossing = [];
        for (let d = -10; d <= 10; d++) { crossing.push(dark(at(gapMidX + d, gapMidY))); crossing.push(dark(at(gapMidX, gapMidY + d))); }
        check('where two double borders cross, their gaps run through each other', crossing.filter((hit) => hit).length, 0);
        const quadrants = [[-1, -1], [1, -1], [-1, 1], [1, 1]].map(([dx, dy]) => dark(at(dx < 0 ? insideV[0][0] : insideV[1][0], dy < 0 ? insideH[0][0] : insideH[1][0])));
        check('and the lines close the four corners round the crossing', quadrants, [true, true, true, true]);
        const topGapY = topRuns.length === 2 ? topRuns[0][0] + topRuns[0][1] : 0;
        const junction = [];
        for (let y = topGapY; y < gapY - 2; y++) junction.push(dark(at(gapMidX, y)));
        check('where the inside border meets the top border, its gap opens into the top border\'s gap; the outer line runs on', [junction.filter((hit) => hit).length, dark(at(gapMidX, topRuns[0][0]))], [0, true]);
    }

    // ---- thin-thick: the first named line is the top / left one ----
    const ttLeft = runs(midY(thinThick, 0.5), thinThick.left - 3, thinThick.left + 14, true, dark);
    const ttTop = runs(midX(thinThick), thinThick.top - 3, thinThick.top + 14, false, dark);
    const ttRight = runs(midY(thinThick, 0.5), thinThick.right - 14, thinThick.right + 3, true, dark);
    check('a thin-thick border is two lines', [ttLeft.length, ttTop.length, ttRight.length], [2, 2, 2]);
    check('thin first, then thick: from the left on the left side, from the top on the top side', [ttLeft.length === 2 && ttLeft[0][1] < ttLeft[1][1], ttTop.length === 2 && ttTop[0][1] < ttTop[1][1]], [true, true]);
    check('and the same order on the right side (the style names the lines left to right)', ttRight.length === 2 && ttRight[0][1] < ttRight[1][1], true);

    // ---- conflicting borders ----
    const sharedX = midX(conflict, 0.5);
    const across = runs(midY(conflict, 0.2), sharedX - 14, sharedX + 14, true, (c) => red(c) || blue(c));
    const redAcross = runs(midY(conflict, 0.2), sharedX - 14, sharedX + 14, true, red);
    check('between a cell with a thick red right border and its neighbour: one line, the thick red one', [across.length, redAcross.length, redAcross.length === 1 && redAcross[0][1] >= 3 && redAcross[0][1] <= 5], [1, 1, true]);
    const below = runs(midY(conflict, 0.8), sharedX - 14, sharedX + 14, true, (c) => red(c) || blue(c));
    check('in the row below the edge is the table\'s thin blue line', [below.length, below.length === 1 && below[0][1] <= 2, blue(at(below.length ? below[0][0] : 0, midY(conflict, 0.8)))], [1, true, true]);
    const sharedY = midY(conflict, 0.5);
    const greenDown = runs(midX(conflict, 0.75), sharedY - 14, sharedY + 14, false, green);
    const blueDown = runs(midX(conflict, 0.25), sharedY - 14, sharedY + 14, false, blue);
    check('a cell\'s thick green top border wins over the row above\'s thin one; under the other cell the thin blue line stays', [greenDown.length === 1 && greenDown[0][1] >= 5 && greenDown[0][1] <= 7, blueDown.length === 1 && blueDown[0][1] <= 2], [true, true]);

    // ---- a thick single border: the corner is filled ----
    // (One pixel in: the outermost row and column can be an antialiased edge.)
    const corner = [at(thick.left + 1, thick.top + 1), at(thick.right - 1, thick.top + 1), at(thick.left + 1, thick.bottom - 1), at(thick.right - 1, thick.bottom - 1)];
    check('the corners of a thick border are filled', corner.map(dark), [true, true, true, true]);
    const thickLeft = runs(midY(thick, 0.75), thick.left - 3, thick.left + 12, true, dark), thickTop = runs(midX(thick), thick.top - 3, thick.top + 12, false, dark);
    check('a 4.5 pt border is 6 pixels wide on both axes', [thickLeft.length === 1 && thickLeft[0][1], thickTop.length === 1 && thickTop[0][1]], [6, 6]);
} catch (e) { fail(e); }
finish();
