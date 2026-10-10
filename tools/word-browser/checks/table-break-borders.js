// The border between two rows of a table is one line, and it belongs to one of
// the two rows. Where a page break parts them each page has the line: the
// last row of a page its bottom border, the first row of the next page its
// top border - also below a header row that is repeated there, and the
// repeated header rows have all of theirs. In pixels (whole canvas, canvas
// text): a column of pixels through a cell, away from its text, has ink at
// the row's edge; rows that follow each other on a page still share one line.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
    const cell = (content) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const rows = 40;
    const table = (headers) => `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>`
        + Array.from({ length: headers }, (_, h) => `<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>${cell(p(`HEAD${h + 1}`))}${cell(p('right'))}</w:tr>`).join('')
        + Array.from({ length: rows }, (_, r) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(p(`row ${r + 1}`))}${cell(p('right'))}</w:tr>`).join('')
        + '</w:tbl>';
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d');
    const open = async (headers) => {
        zip.file('word/document.xml', head + Array.from({ length: 20 }, (_, i) => p(`filler ${i + 1}`)).join('') + table(headers) + p('after the table') + sect + '</w:body></w:document>');
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        return editor;
    };
    // Rows of dark pixels down the column x between two ys: [[y, thickness], ...].
    const linesDown = (x, from, to) => {
        const top = Math.max(0, Math.round(from)), height = Math.max(1, Math.round(to) - top);
        const data = context.getImageData(Math.round(x), top, 1, height).data;
        const found = []; let start = -1;
        for (let i = 0; i < height; i++) { const dark = data[i * 4] < 140 && data[i * 4 + 1] < 140 && data[i * 4 + 2] < 140; if (dark && start < 0) start = i; if (!dark && start >= 0) { found.push([top + start, i - start]); start = -1; } }
        if (start >= 0) found.push([top + start, height - start]);
        return found;
    };
    const study = (editor, label) => {
        const frames = Word.getPageFrames(editor);
        const lines = Word.getLayoutLines(editor);
        const pageIndexOf = (y) => { let index = 0; frames.forEach((frame, i) => { if (frame.Top <= y) index = i; }); return index; };
        const rowLine = (r) => lines.find((line) => (line.Text ?? '') === `row ${r}`);
        const firstOnSecond = Array.from({ length: rows }, (_, r) => r + 1).find((r) => pageIndexOf(rowLine(r).Y) === 1);
        // (A column through the first cell, right of its short text.)
        const x = rowLine(1).X + 180;
        return { frames, lines, rowLine, firstOnSecond, x, pageIndexOf, label };
    };

    // ---- no header rows ----
    {
        const editor = await open(0);
        const { frames, rowLine, firstOnSecond, x } = study(editor);
        check('the table goes on to a second page', firstOnSecond > 2 && firstOnSecond < rows, true);
        const last = rowLine(firstOnSecond - 1), first = rowLine(firstOnSecond), next = rowLine(firstOnSecond + 1);
        const below = linesDown(x, last.Y + last.Height - 2, last.Y + last.Height + 16);
        const above = linesDown(x, first.Y - 16, first.Y + 2);
        note(`row ${firstOnSecond - 1} ends the first page (line at ${Math.round(last.Y)}), row ${firstOnSecond} starts the second (${Math.round(first.Y)}); ink below the last ${JSON.stringify(below)}, above the first ${JSON.stringify(above)}`);
        check('the last row before the break has its bottom border', below.length, 1);
        check('the first row after the break has its top border', above.length, 1);
        check('nothing else is drawn between the last row and the end of the page', linesDown(x, last.Y + last.Height + 16, frames[0].Top + frames[0].Height - 2).length, 0);
        check('nothing above the top border on the second page', linesDown(x, frames[1].Top + 2, first.Y - 16).length, 0);
        check('one line between two rows on a page', [linesDown(x, rowLine(1).Y + rowLine(1).Height - 2, rowLine(2).Y + 2).length, linesDown(x, first.Y + first.Height - 2, next.Y + 2).length], [1, 1]);
        Word.dispose(editor);
    }

    // ---- two header rows, repeated on the second page ----
    {
        const editor = await open(2);
        const { frames, lines, rowLine, firstOnSecond, x, pageIndexOf } = study(editor);
        const heads = (text) => lines.filter((line) => (line.Text ?? '') === text);
        check('the header rows are on both pages', [heads('HEAD1').map((line) => pageIndexOf(line.Y)), heads('HEAD2').map((line) => pageIndexOf(line.Y))], [[0, 1], [0, 1]]);
        const head1 = heads('HEAD1')[1], head2 = heads('HEAD2')[1], first = rowLine(firstOnSecond), last = rowLine(firstOnSecond - 1);
        const header = linesDown(x, frames[1].Top + 2, first.Y + 2);
        note(`on the second page: header rows at ${Math.round(head1.Y)} and ${Math.round(head2.Y)}, row ${firstOnSecond} at ${Math.round(first.Y)}; ink from the top of the page to that row ${JSON.stringify(header)}`);
        check('the repeated header rows have their borders: above, between and below them', header.length, 3);
        check('the first is above the first header row, the last just above the first row', [header[0] && header[0][0] <= head1.Y + 2, header[2] && header[2][0] > head2.Y && header[2][0] <= first.Y + 2], [true, true]);
        check('the last row before the break has its bottom border', linesDown(x, last.Y + last.Height - 2, last.Y + last.Height + 16).length, 1);
        const top1 = heads('HEAD1')[0], top2 = heads('HEAD2')[0];
        check('the header rows at the top of the table: a line above, one between, one below', [linesDown(x, top1.Y - 16, top1.Y + 2).length, linesDown(x, top1.Y + top1.Height - 2, top2.Y + 2).length, linesDown(x, top2.Y + top2.Height - 2, rowLine(1).Y + 2).length], [1, 1, 1]);
        Word.dispose(editor);
    }
} catch (e) { fail(e); }
finish();
