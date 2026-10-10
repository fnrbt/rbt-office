// A vertically merged cell (w:vMerge) whose rows cross a page break: its
// content is split with the rows it spans. It used to run on in one piece -
// through the bottom margin of the page, the gap between the pages and under
// the header row repeated on the next page. In pixels: nothing is painted in
// the bottom margin below the table's last row on the first page, the merged
// column goes on below the repeated header row of the second page; the lines
// (getLayoutLines) are beside rows of the table; a real click on the part on
// the second page puts the caret into the merged cell's text there, and what
// is typed goes into it. Also a row kept with the next ("keep with next" on
// the first paragraph of its first cell) goes to the next page with the row
// that does not fit.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text, extra = '') => `<w:p><w:pPr>${extra}<w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const words = (prefix, count) => Array.from({ length: count }, (_, i) => `${prefix}${i + 1}`).join(' ');
    const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
    const cell = (width, props, content) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${props}</w:tcPr>${content}</w:tc>`;
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const documentOf = (body) => head + body + sect + '</w:body></w:document>';
    const open = async (body, target) => {
        zip.file('word/document.xml', documentOf(body));
        const editor = Word.initializeCanvas(target);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        return editor;
    };
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d');
    const inkBetween = (x0, x1, y0, y1) => { const data = context.getImageData(Math.round(x0), Math.round(y0), Math.round(x1 - x0), Math.max(1, Math.round(y1 - y0))).data; let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 110) ink++; return ink; };

    // ---- the merged cell ----
    const rows = 30;
    // (Two long paragraphs: fewer than the rows, so the importer cannot hand each row one of them.)
    const merged = Array.from({ length: 2 }, (_, i) => p(words(`m${i + 1}x`, 150))).join('');
    const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>`
        + `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(4680, '', p('HEADLEFT'))}${cell(4680, '', p('HEADRIGHT'))}</w:tr>`
        + Array.from({ length: rows }, (_, r) => `<w:tr>${cell(4680, r === 0 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>', r === 0 ? merged : '<w:p/>')}${cell(4680, '', p(`row ${r + 1}`))}</w:tr>`).join('')
        + '</w:tbl>';
    const editor = await open(Array.from({ length: 22 }, (_, i) => p(`filler ${i + 1}`)).join('') + table + p('after the table'), canvas);
    const frames = Word.getPageFrames(editor);
    const lines = Word.getLayoutLines(editor);
    const pageIndexOf = (y) => { let index = 0; frames.forEach((frame, i) => { if (frame.Top <= y) index = i; }); return index; };
    const rowLine = (r) => lines.find((line) => (line.Text ?? '') === `row ${r}`);
    const mergedLines = lines.filter((line) => /^m\d+x\d/.test(line.Text ?? ''));
    note(`rows at ${JSON.stringify(Array.from({ length: rows }, (_, r) => Math.round(rowLine(r + 1).Y)))}, merged lines at ${JSON.stringify(mergedLines.map((line) => Math.round(line.Y)))}, header rows at ${JSON.stringify(lines.filter((line) => line.Text === 'HEADLEFT').map((line) => Math.round(line.Y)))}`);
    const rowsOnFirst = Array.from({ length: rows }, (_, r) => r + 1).filter((r) => pageIndexOf(rowLine(r).Y) === 0);
    const lastOnFirst = rowLine(rowsOnFirst[rowsOnFirst.length - 1]), firstOnSecond = rowLine(rowsOnFirst.length + 1);
    check('two pages; the rows cross the page break', [frames.length, rowsOnFirst.length > 3 && rowsOnFirst.length < rows], [2, true]);
    const first = frames[0], second = frames[1];
    const textBottom = first.Top + first.Height - first.MarginBottom;
    const onFirst = mergedLines.filter((line) => pageIndexOf(line.Y) === 0), onSecond = mergedLines.filter((line) => pageIndexOf(line.Y) === 1);
    check('the merged cell\'s lines are on both pages', [onFirst.length > 3, onSecond.length > 3], [true, true]);
    check('none of them below the last row of the first page (they ran through the bottom margin)', onFirst.filter((line) => line.Y > lastOnFirst.Y + 3).length, 0);
    check('none in the bottom margin, between the pages or in the top margin of the second page', mergedLines.filter((line) => line.Y + 4 > textBottom && line.Y < second.Top + second.MarginTop - 1).length, 0);
    const header = lines.filter((line) => line.Text === 'HEADLEFT');
    check('the header row is repeated; the merged content starts below it, at the first row of the page', [header.length, header.length === 2 && onSecond[0].Y > header[1].Y + 10, Math.abs(onSecond[0].Y - firstOnSecond.Y) < 3], [2, true, true]);
    // Pixels (whole-canvas mode: layout coordinates are canvas pixels).
    const left = first.MarginLeft, middle = first.MarginLeft + (first.Width - first.MarginLeft - first.MarginRight) / 2;
    check('painted: nothing in the bottom margin of the first page, nothing between the pages',
        // (The band below the text area where the next lines of the merged cell were; the page's footer is further down.)
        [inkBetween(left - 4, first.Width - first.MarginRight + 4, textBottom + 3, textBottom + 30), inkBetween(0, first.Width, first.Top + first.Height + 2, second.Top - 2)], [0, 0]);
    check('painted: the merged column goes on below the repeated header row of the second page', [inkBetween(left + 6, middle - 6, firstOnSecond.Y, firstOnSecond.Y + 60) > 40, inkBetween(left + 6, middle - 6, header[1].Y + 2, header[1].Y + 22) > 20], [true, true]);
    check('painted: the repeated header row holds the header text alone (no merged line under it)', inkBetween(left + 6, middle - 6, header[1].Y - 4, header[1].Y + 24) < inkBetween(left + 6, middle - 6, firstOnSecond.Y, firstOnSecond.Y + 28) + 400, true);
    // A real click on the merged text on the second page, and typing there.
    const target = onSecond[1];
    window.state = {};
    window.drive = [
        ['eval', `window.scrollTo(0, ${Math.round(target.Y) - 300})`],
        ['click', Math.round(target.X + 30), 300 + 8],
        ['eval', 'window.afterClick()'],
        ['type', 'TYPED'],
        ['eval', 'window.afterTyping()'],
    ];
    window.afterClick = () => { const selection = Word.getSelection(editor); window.state.clicked = [selection && selection.FocusBlockId === target.BlockId, Word.getCaretContext(editor).InTable, Math.round(Word.getCaretRect(editor).Y)]; };
    window.afterTyping = () => { window.state.typed = Word.getLayoutLines(editor).filter((line) => line.BlockId === target.BlockId).map((line) => line.Text ?? '').join(' '); };
    window.finish = async () => {
        try {
            check('a click on the merged text on the second page: the caret is in that paragraph, on that line', [window.state.clicked[0], window.state.clicked[1], Math.abs(window.state.clicked[2] - Math.round(target.Y)) < 4], [true, true, true]);
            check('typing goes into it', window.state.typed.includes('TYPED'), true);
            const after = Word.getLayoutLines(editor).filter((line) => /^m\d+x\d/.test(line.Text ?? '') || (line.Text ?? '').includes('TYPED'));
            check('and the content is still split with the rows', after.filter((line) => line.Y + 4 > textBottom && line.Y < second.Top + second.MarginTop - 1).length, 0);

            // ---- keep with next on a row ----
            const keepTable = (kept) => `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>`
                + Array.from({ length: 10 }, (_, r) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(4680, '', p(`k${r + 1}`, kept.includes(r + 1) ? '<w:keepNext/>' : ''))}${cell(4680, '', r + 1 === 8 ? Array.from({ length: 6 }, (_, i) => p(`tall ${i + 1}`)).join('') : p(`b${r + 1}`))}</w:tr>`).join('')
                + '</w:tbl>';
            let fillers = 20;
            const pagesOfRows = async (kept) => {
                const other = await open(Array.from({ length: fillers }, (_, i) => p(`filler ${i + 1}`)).join('') + keepTable(kept) + p('after'), document.createElement('canvas'));
                const otherFrames = Word.getPageFrames(other);
                const otherLines = Word.getLayoutLines(other);
                return [5, 6, 7, 8].map((r) => { const line = otherLines.find((candidate) => candidate.Text === `k${r}`); let index = 0; otherFrames.forEach((frame, i) => { if (frame.Top <= line.Y) index = i; }); return index; });
            };
            // (As many paragraphs before the table as put the tall row 8 at the foot of the first page.)
            let plain = await pagesOfRows([]);
            while (fillers < 45 && !(plain[2] === 0 && plain[3] === 1)) { fillers++; plain = await pagesOfRows([]); }
            check('a tall row that does not fit starts the second page; the row before it ends the first', [plain[2], plain[3]], [0, 1]);
            check('kept with the next, the row before it goes with it', await pagesOfRows([7]), [0, 0, 1, 1]);
            check('and so do the rows kept with that one', await pagesOfRows([6, 7]), [0, 1, 1, 1]);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
