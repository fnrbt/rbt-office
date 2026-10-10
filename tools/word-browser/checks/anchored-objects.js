// Anchored pictures in the editor: placed by their anchor (aligned to the
// right margin, to the centre of the page, by an offset from the page, the
// margin and the paragraph - getImageRects against the page frames), and
// text around them - "square": clear of the picture's box; "tight" with a
// wp:wrapPolygon: along the outline, so there is text inside the box where
// the outline leaves room (lines and pixels); "top and bottom": no text
// beside it; no wrapping: text across it. A real click on the text inside the
// picture's box puts the caret into that text; the outline is in the saved
// file.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/floating-image.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const words = (prefix, count) => Array.from({ length: count }, (_, i) => `${prefix}${i + 1}`).join(' ');
    const EMU = 9525;
    // An anchored picture of 192 x 144 px: `h` / `v` are [relativeFrom, align | offset in px], `wrap` the wrap element.
    const picture = (id, h, v, wrap, behind = 0) =>
        `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1">`
        + `<wp:simplePos x="0" y="0"/>`
        + `<wp:positionH relativeFrom="${h[0]}">${typeof h[1] === 'string' ? `<wp:align>${h[1]}</wp:align>` : `<wp:posOffset>${h[1] * EMU}</wp:posOffset>`}</wp:positionH>`
        + `<wp:positionV relativeFrom="${v[0]}">${typeof v[1] === 'string' ? `<wp:align>${v[1]}</wp:align>` : `<wp:posOffset>${v[1] * EMU}</wp:posOffset>`}</wp:positionV>`
        + `<wp:extent cx="${192 * EMU}" cy="${144 * EMU}"/>${wrap}`
        + `<wp:docPr id="${id}" name="Picture ${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${192 * EMU}" cy="${144 * EMU}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const p = (inner) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>${inner}</w:p>`;
    const text = (value) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
    // The outline: the picture's whole width at its top, a point at its bottom left.
    const triangle = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const wraps = { square: '<wp:wrapSquare wrapText="bothSides"/>', tight: `<wp:wrapTight wrapText="bothSides">${triangle}</wp:wrapTight>`, topAndBottom: '<wp:wrapTopAndBottom/>', none: '<wp:wrapNone/>' };
    const open = async (body, target = document.createElement('canvas')) => {
        zip.file('word/document.xml', head + body + sect + '</w:body></w:document>');
        const editor = Word.initializeCanvas(target);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        return editor;
    };
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d');
    const near = (a, b) => Math.abs(a - b) < 1.5;
    // Where a document's picture is: from the lines (the segment of kind "image").
    const pictureOf = (editor) => {
        for (const line of Word.getLayoutLines(editor)) for (const segment of line.Segments ?? []) if (segment.Kind === 'image') return { X: segment.X, Y: segment.Y ?? line.Y, Width: segment.Width, Height: segment.Height, line };
        return null;
    };

    // ---- where the anchor puts it ----
    const placedAt = async (h, v) => {
        const editor = await open(p(text(words('a', 40))) + p(text(words('b', 30)) + picture(1, h, v, wraps.none) + text(words('c', 60))));
        const frame = Word.getPageFrames(editor)[0];
        const found = pictureOf(editor);
        const anchorParagraph = Word.getLayoutLines(editor).filter((line) => (line.Text ?? '').startsWith('b1 '))[0];
        return { x: found.X, y: found.Y, frame, paragraphTop: anchorParagraph.Y, lineTop: found.line.Y };
    };
    const probe = await placedAt(['page', 40], ['page', 300]);
    if (probe.y === undefined || Number.isNaN(probe.y)) throw new Error('the lines do not say where a picture is');
    check('by offsets from the page', [near(probe.x, 40), near(probe.y, 300)], [true, true]);
    const right = await placedAt(['margin', 'right'], ['margin', 'top']);
    check('aligned to the right and to the top of the text area', [near(right.x, right.frame.Width - right.frame.MarginRight - 192), near(right.y, right.frame.MarginTop)], [true, true]);
    const centre = await placedAt(['page', 'center'], ['page', 'bottom']);
    check('in the centre of the page, at its bottom', [near(centre.x, (centre.frame.Width - 192) / 2), near(centre.y, centre.frame.Height - 144)], [true, true]);
    const offsets = await placedAt(['margin', 50], ['margin', 60]);
    check('by offsets from the margins', [near(offsets.x, offsets.frame.MarginLeft + 50), near(offsets.y, offsets.frame.MarginTop + 60)], [true, true]);
    const paragraph = await placedAt(['column', 10], ['paragraph', 25]);
    check('from the column and the top of the anchor\'s paragraph', [near(paragraph.x, paragraph.frame.MarginLeft + 10), near(paragraph.y, paragraph.paragraphTop + 25)], [true, true]);

    // ---- text around it: the picture at the left margin, 40 px below the top of its paragraph ----
    const around = async (wrap, target) => {
        const editor = await open(p(picture(2, ['margin', 0], ['paragraph', 40], wraps[wrap]) + text(words('t', 260))) + p(text(words('after', 20))), target);
        const frame = Word.getPageFrames(editor)[0];
        const found = pictureOf(editor);
        const lines = Word.getLayoutLines(editor).filter((line) => /^\uFFFC?t\d/.test(line.Text ?? ''));
        const beside = lines.filter((line) => line.Y + line.Height > found.Y + 1 && line.Y < found.Y + 144 - 1);
        return { editor, frame, found, lines, beside };
    };
    const square = await around('square');
    check('square: the lines beside the picture start right of its box', [square.beside.length >= 4, square.beside.every((line) => line.X >= square.frame.MarginLeft + 192 - 1)], [true, true]);
    const topAndBottom = await around('topAndBottom');
    check('top and bottom: no line beside the picture, the text goes on below it', [topAndBottom.beside.length, topAndBottom.lines.some((line) => line.Y >= topAndBottom.found.Y + 144 - 1)], [0, true]);
    const none = await around('none');
    check('no wrapping: the lines beside the picture start at the margin', [none.beside.length >= 4, none.beside.every((line) => near(line.X, none.frame.MarginLeft))], [true, true]);
    const tight = await around('tight', canvas);
    const starts = tight.beside.map((line) => Math.round(line.X));
    note(`tight: the lines beside the picture start at ${JSON.stringify(starts)} (the picture: ${Math.round(tight.found.X)} .. ${Math.round(tight.found.X) + 192})`);
    check('tight, with an outline: each line starts right of the outline at its height - further left the further down', [starts.length >= 4, starts.every((x, i) => i === 0 || x <= starts[i - 1] + 1), starts[0] - starts[starts.length - 1] > 60], [true, true, true]);
    check('the last of them inside the picture\'s box (square keeps it clear)', [starts[starts.length - 1] < tight.frame.MarginLeft + 192 - 40, Math.min(...square.beside.map((line) => line.X)) >= square.frame.MarginLeft + 192 - 1], [true, true]);
    // Pixels: text in the lower right half of the picture's box (whole-canvas mode: layout pixels are canvas pixels).
    // (The picture is in front of the text - it is not behindDoc -, and it is half transparent blue: the text under it is dark but for the blue.)
    const inkBetween = (x0, x1, y0, y1) => { const data = context.getImageData(Math.round(x0), Math.round(y0), Math.max(1, Math.round(x1 - x0)), Math.max(1, Math.round(y1 - y0))).data; let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 160) ink++; return ink; };
    const last = tight.beside[tight.beside.length - 1];
    const boxRight = tight.frame.MarginLeft + 192;
    check('painted: text inside the picture\'s box, right of the outline, on the last line beside it', inkBetween(last.X, boxRight - 4, last.Y + 2, last.Y + last.Height - 2) > 30, true);
    // A real click there.
    window.state = {};
    window.drive = [
        ['click', Math.round(last.X + 24), Math.round(last.Y + last.Height / 2)],
        ['eval', 'window.afterClick()'],
        ['type', 'ZZ'],
        ['eval', 'window.afterTyping()'],
        ['click', Math.round(tight.found.X + 20), Math.round(tight.found.Y + 20)],
        ['eval', 'window.afterPictureClick()'],
    ];
    window.afterPictureClick = () => { window.state.picture = Word.getCaretContext(tight.editor).ImageSelected; };
    window.afterClick = () => { const selection = Word.getSelection(tight.editor); window.state.clicked = [selection && selection.FocusBlockId === last.BlockId, selection && selection.FocusOffset >= last.StartOffset && selection.FocusOffset <= last.EndOffset]; };
    window.afterTyping = () => { window.state.typed = Word.getText(tight.editor).includes('ZZ'); };
    window.finish = async () => {
        try {
            check('a click on the text inside the picture\'s box: the caret is in that line', window.state.clicked, [true, true]);
            check('typing goes into it', window.state.typed, true);
            check('a click inside the outline is on the picture: it is selected', window.state.picture, true);
            // Saved: the outline is in the file, and read again the lines are the same.
            const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(tight.editor))).arrayBuffer());
            const saved = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
            check('saved: the outline is in the file', saved.includes('<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'), true);
            const again = Word.initializeCanvas(document.createElement('canvas'));
            Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(bytes));
            const againPicture = pictureOf(again);
            const againStarts = Word.getLayoutLines(again).filter((line) => (line.Text ?? '').startsWith('t') || (line.Text ?? '').includes('ZZ')).filter((line) => line.Y + line.Height > againPicture.Y + 1 && line.Y < againPicture.Y + 144 - 1).map((line) => Math.round(line.X));
            check('read again: the lines go along the outline as before', [againStarts.length, againStarts[0] - againStarts[againStarts.length - 1] > 60], [starts.length, true]);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
