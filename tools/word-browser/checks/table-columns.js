// A table in a section of two columns goes from column to column and page to
// page row by row (it was placed whole: it left the rest of its column empty
// and, longer than a column, ran on below the page). The rows' lines
// (getLayoutLines) are in the columns of both pages, inside the text area;
// the header row is repeated at the top of every column; the borders are
// painted beside the rows (pixels) and nothing below the text area; a real
// click on a row in the second column puts the caret there and typing goes
// into that cell.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
    const cell = (width, content) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
    const rows = 90;
    const table = `<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="2160"/><w:gridCol w:w="2160"/></w:tblGrid>`
        + `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(2160, p('HEADA'))}${cell(2160, p('HEADB'))}</w:tr>`
        + Array.from({ length: rows }, (_, r) => `<w:tr>${cell(2160, p(`row ${r + 1}`))}${cell(2160, p(`val ${r + 1}`))}</w:tr>`).join('')
        + '</w:tbl>';
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/></w:sectPr>';
    zip.file('word/document.xml', head + Array.from({ length: 5 }, (_, i) => p(`filler ${i + 1}`)).join('') + table + p('after the table') + sect + '</w:body></w:document>');
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
    const context = canvas.getContext('2d');
    const inkBetween = (x0, x1, y0, y1) => { const data = context.getImageData(Math.round(x0), Math.round(y0), Math.max(1, Math.round(x1 - x0)), Math.max(1, Math.round(y1 - y0))).data; let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 110) ink++; return ink; };
    const frames = Word.getPageFrames(editor);
    const lines = Word.getLayoutLines(editor);
    const first = frames[0];
    const columnWidth = (first.Width - first.MarginLeft - first.MarginRight - 48) / 2;
    const secondColumnX = first.MarginLeft + columnWidth + 48;
    const pageIndexOf = (y) => { let index = 0; frames.forEach((frame, i) => { if (frame.Top <= y) index = i; }); return index; };
    const placeOf = (line) => `${pageIndexOf(line.Y)}/${line.X >= secondColumnX - 2 ? 1 : 0}`;
    const rowLine = (r) => lines.find((line) => line.Text === `row ${r}`);
    const places = [...new Set(Array.from({ length: rows }, (_, r) => placeOf(rowLine(r + 1))))];
    note(`the rows are in ${JSON.stringify(places)} (page/column); ${frames.length} pages`);
    check('the rows are in the columns of more than one page, in order', [places.length >= 3, JSON.stringify(places) === JSON.stringify([...places].sort()), places[0]], [true, true, '0/0']);
    const outside = Array.from({ length: rows }, (_, r) => rowLine(r + 1)).filter((line) => { const frame = frames[pageIndexOf(line.Y)]; return line.Y < frame.Top + frame.MarginTop - 1 || line.Y + 16 > frame.Top + frame.Height - frame.MarginBottom + 1; });
    check('every row is in the text area of its page (the table ran on below it)', outside.length, 0);
    const headers = lines.filter((line) => line.Text === 'HEADA');
    check('the header row is at the top of every column the table is in', [headers.length, JSON.stringify(headers.map(placeOf))], [places.length, JSON.stringify(places)]);
    // The first row of the second column, and the last row of the first.
    const firstOfSecond = Array.from({ length: rows }, (_, r) => rowLine(r + 1)).find((line) => placeOf(line) === '0/1');
    const lastOfFirst = Array.from({ length: rows }, (_, r) => rowLine(r + 1)).filter((line) => placeOf(line) === '0/0').pop();
    check('the second column starts at the top of the text area, below the repeated header row', [Math.abs(headers[1].Y - (first.Top + first.MarginTop)) < 8, firstOfSecond.Y > headers[1].Y + 10], [true, true]);
    check('the first column is filled to the foot of the text area', first.Top + first.Height - first.MarginBottom - lastOfFirst.Y < 80, true);
    // Pixels: borders beside the rows of the second column; none in the gap between the columns; nothing below the text area.
    const textBottom = first.Top + first.Height - first.MarginBottom;
    check('painted: the left and right border of the table beside a row of the second column', [inkBetween(secondColumnX - 3, secondColumnX + 4, firstOfSecond.Y + 4, firstOfSecond.Y + 12) > 3, inkBetween(secondColumnX + 288 - 4, secondColumnX + 288 + 4, firstOfSecond.Y + 4, firstOfSecond.Y + 12) > 3], [true, true]);
    check('painted: its text, in the column', inkBetween(secondColumnX + 4, secondColumnX + 140, firstOfSecond.Y, firstOfSecond.Y + 18) > 20, true);
    check('painted: nothing between the columns, nothing of the table below the text area', [inkBetween(first.MarginLeft + 288 + 6, secondColumnX - 6, first.Top + first.MarginTop, textBottom), inkBetween(first.MarginLeft - 4, first.Width - first.MarginRight + 4, textBottom + 4, textBottom + 30)], [0, 0]);
    check('painted: a line under the last row of the first column', inkBetween(first.MarginLeft + 20, first.MarginLeft + 200, lastOfFirst.Y + 14, lastOfFirst.Y + 40) > 100, true);
    window.state = {};
    window.drive = [
        ['eval', `window.scrollTo(0, ${Math.max(0, Math.round(firstOfSecond.Y) - 300)})`],
        ['click', Math.round(firstOfSecond.X + 20), Math.round(firstOfSecond.Y + 8) - Math.max(0, Math.round(firstOfSecond.Y) - 300)],
        ['eval', 'window.afterClick()'],
        ['type', 'XY'],
        ['eval', 'window.afterTyping()'],
    ];
    window.afterClick = () => { const selection = Word.getSelection(editor); window.state.clicked = [selection && selection.FocusBlockId === firstOfSecond.BlockId, Word.getCaretContext(editor).InTable, Math.round(Word.getCaretRect(editor).X) >= Math.round(secondColumnX)]; };
    // (Through the API as well: a point in the second cell of a row of the first column.)
    const second = lines.find((line) => line.Text === 'val 2');
    const hitSecond = Word.hitTestPosition(editor, second.X + 10, second.Y + 8);
    check('a point in the second cell of a row is in that cell (it was in the first cell)', hitSecond && hitSecond.BlockId, second.BlockId);
    window.afterTyping = () => { window.state.typed = Word.getLayoutLines(editor).filter((line) => line.BlockId === firstOfSecond.BlockId).map((line) => line.Text ?? '').join(''); };
    window.finish = async () => {
        try {
            check('a click on a row in the second column: the caret is in that cell, in that column', window.state.clicked, [true, true, true]);
            check('typing goes into it', [window.state.typed.includes('XY'), window.state.typed.replace('XY', '')], [true, firstOfSecond.Text]);
            // Saved and read again: the same places.
            const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
            const again = Word.initializeCanvas(document.createElement('canvas'));
            Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(bytes));
            const againLines = Word.getLayoutLines(again);
            const againFrames = Word.getPageFrames(again);
            const againPlaces = [...new Set(againLines.filter((line) => /^row \d+$/.test(line.Text ?? '')).map((line) => { let index = 0; againFrames.forEach((frame, i) => { if (frame.Top <= line.Y) index = i; }); return `${index}/${line.X >= secondColumnX - 2 ? 1 : 0}`; }))];
            check('saved and read again: the rows are in the same columns', JSON.stringify(againPlaces), JSON.stringify(places));
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
