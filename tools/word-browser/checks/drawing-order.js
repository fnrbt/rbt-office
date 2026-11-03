// The order a page is painted in, in pixels: a header's drawings (a text box,
// a watermark's WordArt) below the body; the body's drawings behind the text
// (behindDoc) below its paragraphs' shading and text; the drawings in front of
// the text above it, in the order of their relativeHeight - a text box over a
// picture and a picture over a text box, whichever comes first in the file.
// Text boxes: placed on the page by their anchor, their text inside their
// insets and clipped to the box, a box as high as its text; the text around a
// box that wraps it. Saved: a text box nobody changed is the element it was
// read from (the DrawingML choice and the VML fallback, with what the model
// does not hold) also when its paragraph is written from the model; a changed
// one is written from the model in both.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/floating-image.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const words = (prefix, count) => Array.from({ length: count }, (_, i) => `${prefix}${i + 1}`).join(' ');
    const EMU = 9525;
    const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
    const anchorOpen = (id, o) =>
        `<wp:anchor distT="0" distB="0" distL="${(o.dist ?? 0) * EMU}" distR="${(o.dist ?? 0) * EMU}" simplePos="0" relativeHeight="${o.z}" behindDoc="${o.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`
        + `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${o.relH ?? 'page'}"><wp:posOffset>${o.x * EMU}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${o.relV ?? 'page'}"><wp:posOffset>${o.y * EMU}</wp:posOffset></wp:positionV>`
        + `<wp:extent cx="${o.w * EMU}" cy="${o.h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${o.wrap ?? '<wp:wrapNone/>'}<wp:docPr id="${id}" name="${o.name ?? 'Shape ' + id}"/><wp:cNvGraphicFramePr/>`;
    // A text box as Word writes it: the DrawingML choice and a VML fallback. `o.extra`: what the model does not hold.
    const box = (id, o, content) =>
        `<w:r><mc:AlternateContent><mc:Choice Requires="wps" xmlns:wps="${WPS}"><w:drawing>${anchorOpen(id, o)}`
        + `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${o.w * EMU}" cy="${o.h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
        + `${o.fill ? `<a:solidFill><a:srgbClr val="${o.fill}"/></a:solidFill>` : '<a:noFill/>'}<a:ln w="9525">${o.line ? `<a:solidFill><a:srgbClr val="${o.line}"/></a:solidFill>` : '<a:noFill/>'}</a:ln></wps:spPr>${o.extra ?? ''}`
        + `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${(o.inset ?? [9.6])[0] * EMU}" tIns="${(o.inset ?? [0, 4.8])[1] * EMU}" rIns="${9.6 * EMU}" bIns="${4.8 * EMU}" anchor="${o.anchor ?? 't'}" anchorCtr="0">${o.fit ? '<a:spAutoFit/>' : '<a:noAutofit/>'}</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>`
        + `<mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="vml${id}" type="#_x0000_t202" style="position:absolute;margin-left:${o.x * 0.75}pt;margin-top:${o.y * 0.75}pt;width:${o.w * 0.75}pt;height:${o.h * 0.75}pt;z-index:${o.z}" ${o.fill ? `fillcolor="#${o.fill}"` : 'filled="f"'} stroked="f"><v:textbox><w:txbxContent>${content}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
    const picture = (id, o) =>
        `<w:r><w:drawing>${anchorOpen(id, o)}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${o.w * EMU}" cy="${o.h * EMU}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const text = (value, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${value}</w:t></w:r>`;
    const p = (inner, pPr = '') => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>${pPr}</w:pPr>${inner}</w:p>`;
    const sectWith = (header) => `<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="rIdHeader"/>' : ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
    const rootOpen = head.slice(head.indexOf('<w:document'), head.indexOf('>', head.indexOf('<w:document')) + 1);
    const baseRels = await zip.file('word/_rels/document.xml.rels').async('string');
    // The fixture's picture file, made a picture of one colour (orange) that is decoded before it is painted.
    const pictureTarget = /Id="rId1"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="rId1"/.exec(baseRels);
    const pictureSource = (() => { const c = document.createElement('canvas'); c.width = 8; c.height = 6; const g = c.getContext('2d'); g.fillStyle = '#FF8000'; g.fillRect(0, 0, 8, 6); return c.toDataURL('image/png'); })();
    zip.file('word/' + (pictureTarget[1] ?? pictureTarget[2]), Uint8Array.from(atob(pictureSource.slice(pictureSource.indexOf(',') + 1)), (ch) => ch.charCodeAt(0)));
    const pictureImage = new Image(); pictureImage.src = pictureSource; await pictureImage.decode();
    const baseTypes = await zip.file('[Content_Types].xml').async('string');
    const open = async (body, header, target = document.createElement('canvas')) => {
        zip.file('word/document.xml', head + body + sectWith(header) + '</w:body></w:document>');
        if (header) {
            zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rootOpen.replace('<w:document', '<w:hdr')}${header}</w:hdr>`);
            zip.file('word/_rels/document.xml.rels', baseRels.replace('</Relationships>', '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
            zip.file('[Content_Types].xml', baseTypes.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
        } else {
            zip.remove('word/header1.xml');
            zip.file('word/_rels/document.xml.rels', baseRels);
            zip.file('[Content_Types].xml', baseTypes);
        }
        const bytes = await zip.generateAsync({ type: 'uint8array' });
        const editor = Word.initializeCanvas(target);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.setImageCache(editor, new Map([[pictureSource, pictureImage]]));
        return { editor, bytes };
    };
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d', { willReadFrequently: true });
    const near = (a, b, slack = 1.5) => Math.abs(a - b) < slack;
    // The share of the pixels of a rectangle (page coordinates) that `is` says yes to.
    const share = (frame, x, y, w, h, is) => {
        const data = context.getImageData(Math.round((frame.X ?? 0) + x), Math.round((frame.Y ?? 0) + y), w, h).data;
        let hits = 0;
        for (let i = 0; i < data.length; i += 4) if (is(data[i], data[i + 1], data[i + 2])) hits++;
        return hits / (w * h);
    };
    const green = (r, g, b) => r < 60 && g > 200 && b < 60, blue = (r, g, b) => r < 60 && g < 60 && b > 200, magenta = (r, g, b) => r > 200 && g < 60 && b > 200;
    const cyan = (r, g, b) => r < 60 && g > 200 && b > 200, yellow = (r, g, b) => r > 200 && g > 200 && b < 80, white = (r, g, b) => r > 245 && g > 245 && b > 245;

    // ---- the layers of a page ----
    // The body: a shaded paragraph down to about y = 420; behind it a green box that sticks out into the left margin;
    // in front a blue box (height 20) over a picture (10) over a magenta box (5) - the magenta one last in the file.
    const shaded = '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>';
    const bodyXml =
        p(box(11, { x: 20, y: 150, w: 200, h: 100, z: 3, behind: true, fill: '00FF00' }, p(text('behind')))
            + box(12, { x: 400, y: 120, w: 150, h: 100, z: 20, fill: '0000FF' }, p(text('blue')))
            + picture(13, { x: 450, y: 150, w: 192, h: 144, z: 10 })
            + text(words('shaded', 110))
            + box(14, { x: 500, y: 240, w: 100, h: 80, z: 5, fill: 'FF00FF' }, p(text('magenta'))), shaded)
        + p(text('after the shaded paragraph'));
    // The header: a cyan box at the left edge of the page, half of it under the body's shaded paragraph; a watermark.
    const watermark = '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:400pt;height:100pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="#FF0000" stroked="f"><v:textpath style="font-family:&quot;Arial&quot;;font-size:1pt" string="DRAFT"/></v:shape></w:pict></w:r>';
    const headerXml = p(text('the header') + box(21, { x: 20, y: 300, w: 200, h: 100, z: 1, fill: '00FFFF' }, p(text('in the header'))) + watermark);
    const layered = await open(bodyXml, headerXml, canvas);
    const frame = Word.getPageFrames(layered.editor)[0];
    const shadedLines = Word.getLayoutLines(layered.editor).filter((line) => (line.Text ?? '').includes('shaded'));
    const shadedBottom = Math.max(...shadedLines.map((line) => line.Y + line.Height)) - (frame.Y ?? 0);
    note(`the shaded paragraph ends at y = ${Math.round(shadedBottom)}`);
    check('the shaded paragraph reaches below the boxes', shadedBottom > 410, true);
    const pictureColour = (r, g, b) => r > 235 && g > 110 && g < 145 && b < 30;
    check('the picture is painted over the paragraph\'s shading and text', share(frame, 600, 225, 40, 60, pictureColour) > 0.98, true);
    check('behind the text: the box shows in the margin', share(frame, 30, 190, 50, 50, green) > 0.95, true);
    check('behind the text: the paragraph\'s shading and text cover it', [share(frame, 110, 160, 100, 80, green), share(frame, 110, 160, 100, 80, yellow) > 0.5], [0, true]);
    check('in front of the text: the box covers the shading and the text', share(frame, 405, 170, 40, 45, blue) > 0.98, true);
    check('a box of a higher relativeHeight covers the picture', [share(frame, 460, 160, 80, 50, blue) > 0.98, share(frame, 460, 160, 80, 50, pictureColour)], [true, 0]);
    check('the picture covers a box of a lower relativeHeight that comes after it in the file', [share(frame, 505, 245, 90, 44, pictureColour) > 0.95, share(frame, 505, 245, 90, 44, magenta)], [true, 0]);
    check('that box shows below the picture', share(frame, 505, 300, 90, 16, magenta) > 0.9, true);
    check('the header\'s box shows in the margin', share(frame, 30, 310, 50, 80, cyan) > 0.9, true);
    check('the body covers the header\'s box', [share(frame, 110, 310, 100, 80, cyan), share(frame, 110, 310, 100, 80, yellow) > 0.5], [0, true]);
    // The watermark: red letters at half... (no opacity given: opaque) across the middle of the text area, turned; the body's shading covers them.
    const red = (r, g, b) => r > 200 && g < 70 && b < 70;
    const textMiddleY = (frame.MarginTop + (frame.Height - frame.MarginBottom)) / 2;
    const below = share(frame, 200, Math.round(shadedBottom) + 30, 420, Math.round(frame.Height - frame.MarginBottom - shadedBottom) - 60, red);
    note(`the watermark: ${(below * 100).toFixed(1)}% of the text area below the shaded paragraph is red`);
    check('the watermark is painted below the body\'s text area', below > 0.01, true);
    check('turned: its letters are found left below and right above the middle', [share(frame, 150, Math.round(textMiddleY) + 60, 180, 160, red) > 0.005, share(frame, 150, Math.round(shadedBottom) + 5, 150, 40, red) < 0.005], [true, true]);
    check('the body\'s shading covers the watermark', share(frame, 300, 200, 90, 180, red), 0);
    // The same pixels through a viewport (the chunks of a band).
    const pkg = Word.exportDocumentPackage(layered.editor);
    check('the header\'s shapes are read: a text box and a watermark anchored in its paragraph', (pkg.HeaderParts[0].Paragraphs ?? []).map((item) => [item.Type, item.FitText === true, item.AnchorBlock ?? null]).slice(1), [['TextBox', false, null], ['TextBox', true, null]]);

    // ---- where a text box is, and its text ----
    const placed = await open(
        p(text(words('a', 60)))
        + p(text('anchor ') + box(31, { relH: 'column', relV: 'paragraph', x: 50, y: 30, w: 220, h: 90, z: 2, fill: 'FFFFFF', line: '000000', inset: [30, 20] }, p(text('one')) + p(text('two')) + p(text('three')) + p(text('four')) + p(text('five')) + p(text('six')) + p(text('seven')))
            + text(words('b', 20)))
        + p(text(words('c', 100)))
        + p(text('second ') + box(32, { relH: 'column', relV: 'paragraph', x: 300, y: 0, w: 200, h: 30, z: 2, fill: 'EEEEEE', fit: true }, p(text('fitted one')) + p(text('fitted two')) + p(text('fitted three')))), null, canvas);
    const placedFrame = Word.getPageFrames(placed.editor)[0];
    const anchorLine = Word.getLayoutLines(placed.editor).filter((line) => (line.Text ?? '').startsWith('anchor'))[0];
    const shapes = typeof Word.getShapeRects === 'function' ? Word.getShapeRects(placed.editor) : [];
    check('getShapeRects: two shapes', shapes.length, 2);
    if (shapes.length === 2) {
        const first = shapes[0], second = shapes[1];
        check('from the column and its anchor\'s paragraph, by its offsets', [near(first.X, placedFrame.MarginLeft + 50), near(first.Y, anchorLine.Y + 30), near(first.Width, 220), near(first.Height, 90)], [true, true, true, true]);
        check('a box that follows its text is as high as it', [second.Height > 45, second.Height < 90], [true, true]);
        // Its text: inside the insets, clipped at the box's bottom.
        const ink = (x, y, w, h) => share(placedFrame, x, y - (placedFrame.Y ?? 0), w, h, (r, g, b) => r < 110 && g < 110 && b < 110);
        check('the text starts at the left inset', [ink(first.X + 2, first.Y + 20, 26, 60), ink(first.X + 30, first.Y + 20, 40, 60) > 0.02], [0, true]);
        check('and at the top inset', ink(first.X + 2, first.Y + 2, 200, 16), 0);
        check('the text is clipped to the box', ink(first.X + 2, first.Y + first.Height + 2, 200, 40) < 0.05, true);
    }

    // ---- text around a box ----
    const wrapped = async (wrap) => {
        const opened = await open(p(box(41, { relH: 'column', relV: 'paragraph', x: 0, y: 20, w: 200, h: 120, z: 2, fill: 'DDDDDD', wrap, dist: 10 }, p(text('boxed'))) + text(words('t', 260))) + p(text(words('after', 20))));
        const lines = Word.getLayoutLines(opened.editor).filter((line) => /^t\d/.test(line.Text ?? ''));
        const frame = Word.getPageFrames(opened.editor)[0];
        const top = lines[0].Y + 20;
        return { lines, frame, beside: lines.filter((line) => line.Y + line.Height > top + 1 && line.Y < top + 120 - 1) };
    };
    const square = await wrapped('<wp:wrapSquare wrapText="bothSides"/>');
    check('square: the lines beside the box start right of it and its distance from text', [square.beside.length >= 4, square.beside.every((line) => line.X >= square.frame.MarginLeft + 210 - 1)], [true, true]);
    const topAndBottom = await wrapped('<wp:wrapTopAndBottom/>');
    check('top and bottom: no line beside the box', topAndBottom.beside.length, 0);
    const none = await wrapped('<wp:wrapNone/>');
    check('no wrapping: the lines beside the box start at the margin', [none.beside.length >= 4, none.beside.every((line) => near(line.X, none.frame.MarginLeft))], [true, true]);

    // ---- saved ----
    const extra = '<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>';
    const keptBox = box(51, { relH: 'column', relV: 'paragraph', x: 40, y: 40, w: 200, h: 80, z: 7, fill: 'CCDDEE', inset: [20, 10], anchor: 'ctr', extra }, p(text('kept box')));
    const saved = await open(p(text('first paragraph')) + p(text('anchor paragraph of the box') + keptBox) + p(text('last paragraph')));
    const documentOf = async (editor) => (await (await window.JSZip.loadAsync(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer()))).file('word/document.xml').async('string'));
    const source = keptBox.slice('<w:r>'.length, -'</w:r>'.length);
    check('an unchanged document: the text box as it was read', (await documentOf(saved.editor)).includes(source), true);
    // Part of the anchor's paragraph made bold: the change is made in the paragraph's own XML, the box in it is the text it was.
    const anchorBlock = Word.getLayoutLines(saved.editor).filter((line) => (line.Text ?? '').startsWith('anchor paragraph'))[0];
    Word.setSelection(saved.editor, anchorBlock.BlockId, 0, anchorBlock.BlockId, 6);
    Word.toggleBold(saved.editor);
    const boldSaved = await documentOf(saved.editor);
    check('its paragraph changed: the bold is saved', /<w:b\/>[^]*?anchor/.test(boldSaved), true);
    check('its paragraph changed: the box is the text it was read from', boldSaved.includes(source), true);
    // A bookmark set in the paragraph: it is written from its source too (docs/save-architecture.md) - the box in it
    // is the text it was read from.
    Word.setSelection(saved.editor, anchorBlock.BlockId, 7, anchorBlock.BlockId, 16);
    Word.setBookmark(saved.editor, 'beside_the_box');
    const modelSaved = await documentOf(saved.editor);
    check('a bookmark set in its paragraph: the bookmark is saved', modelSaved.includes('w:name="beside_the_box"'), true);
    check('a bookmark set in its paragraph: the box is the text it was read from - choice, fallback and what the model does not hold', [modelSaved.includes(source), modelSaved.includes(extra)], [true, true]);
    check('read again: the same box', (await OF.processors.forEditorPackage(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(saved.editor))).arrayBuffer()))).Paragraphs.filter((item) => item.Type === 'TextBox').map((item) => [Math.round(item.InsetLeft), Math.round(item.InsetTop), item.VerticalAnchor, item.Width, item.Height]), [[20, 10, 'ctr', 200, 80]]);
    // The box itself changed (its fill): both the choice and the fallback are written from the model, with its insets and anchor.
    Word.setCurrentTextBoxStyle(saved.editor, 40, 40, 200, 80, '#112233', '#445566', 1);
    const changedSaved = await documentOf(saved.editor);
    check('the box changed: the choice says the new fill, and keeps the insets and the anchor', [/<a:srgbClr val="112233"\/>/.test(changedSaved), /<wps:bodyPr[^>]*lIns="190500"[^>]*tIns="95250"[^>]*anchor="ctr"/.test(changedSaved)], [true, true]);
    check('the box changed: the fallback says the new fill too', /<mc:Fallback><w:pict><v:shape[^>]*fillcolor="#112233"[^>]*>\s*<v:textbox inset="15\.00pt,7\.50pt,7\.20pt,3\.60pt"/.test(changedSaved), true);
    check('the box changed: what the model does not hold of it is gone with the old element', changedSaved.includes('<wps:style>'), false);
    check('its text is in both', (changedSaved.match(/kept box/g) ?? []).length, 2);

    // ---- a text box edited in place: real clicks and keys ----
    const edit = await open(
        p(text(words('body', 40)))
        + p(text('anchor ') + box(61, { x: 300, y: 300, w: 240, h: 120, z: 9, fill: 'FFFFCC', line: '000000' }, p(text('first line')) + p(text('second line'))))
        + p(text(words('tail', 40))), null, canvas);
    const editor = edit.editor;
    const editFrame = Word.getPageFrames(editor)[0];
    const boxRect = () => Word.getShapeRects(editor)[0];
    const rect = boxRect();
    const inBoxX = Math.round((editFrame.X ?? 0) + rect.X + 60), inBoxY = Math.round(rect.Y + 14);
    const state = { before: Word.getCaretContext(editor).Story };
    window.steps = {
        entered: () => { const c = Word.getCaretContext(editor); state.entered = [c.Story, c.TextBoxId === rect.Id, c.InHeaderFooter]; },
        typed: () => { state.typed = [boxRect().Text, Word.canUndo(editor)]; },
        selectedAll: () => {
            const selection = Word.getSelection(editor);
            state.selectedAll = selection.AnchorBlockId !== selection.FocusBlockId;
            Word.toggleBold(editor);
            state.bold = Word.getSelectionFormat(editor).Bold;
        },
        undone: () => { state.undone = [Word.getSelectionFormat(editor).Bold, boxRect().Text, Word.getCaretContext(editor).Story]; },
        escaped: () => { state.escaped = [Word.getCaretContext(editor).Story, boxRect().Text]; },
        again: () => { state.again = Word.getCaretContext(editor).Story; },
        outside: () => { state.outside = Word.getCaretContext(editor).Story; },
        bodyTyped: () => { state.bodyTyped = [boxRect().Text.includes('Q'), Word.getText(editor).includes('Q')]; },
    };
    window.drive = [
        ['click', inBoxX, inBoxY], ['eval', 'window.steps.entered()'],
        ['type', 'ZZ'], ['eval', 'window.steps.typed()'],
        ['press', 'Control+a'], ['eval', 'window.steps.selectedAll()'],
        ['press', 'Control+z'], ['eval', 'window.steps.undone()'],
        ['press', 'Escape'], ['eval', 'window.steps.escaped()'],
        ['click', inBoxX, inBoxY], ['eval', 'window.steps.again()'],
        ['click', Math.round((editFrame.X ?? 0) + editFrame.MarginLeft + 80), Math.round(editFrame.MarginTop + 8)], ['eval', 'window.steps.outside()'],
        ['type', 'Q'], ['eval', 'window.steps.bodyTyped()'],
    ];
    window.finish = async () => {
        try {
            check('before the click: the caret is in the body', state.before, 'body');
            check('a click in the box: the caret is in its text (getCaretContext: Story "textbox", its id; not a header)', state.entered, ['textbox', true, false]);
            check('typing goes into the box', [/ZZ/.test(state.typed[0]), state.typed[0].includes('second line'), state.typed[1]], [true, true, true]);
            check('Ctrl+A selects the box\'s text, and it is made bold', [state.selectedAll, state.bold], [true, true]);
            check('Ctrl+Z takes the bold back, in the box', [state.undone[0], /ZZ/.test(state.undone[1]), state.undone[2]], [false, true, 'textbox']);
            check('Escape: back in the body, the box keeps its text', [state.escaped[0], /ZZ/.test(state.escaped[1])], ['body', true]);
            check('a click in the box again, then one outside it', [state.again, state.outside], ['textbox', 'body']);
            check('typing then goes into the body', state.bodyTyped, [false, true]);
            // By the API.
            check('enterTextBox / exitTextBox', [Word.enterTextBox(editor, rect.Id), Word.getCaretContext(editor).Story, Word.exitTextBox(editor), Word.getCaretContext(editor).Story, Word.exitTextBox(editor), Word.enterTextBox(editor, 'no such box')], [true, 'textbox', true, 'body', false, false]);
            // Saved: the edited box from the model, its text in the choice and in the fallback; read again it is the same box.
            const editedXml = await documentOf(editor);
            const editedBox = editedXml.slice(editedXml.indexOf('<mc:AlternateContent'), editedXml.indexOf('</mc:AlternateContent>'));
            check('saved: the typed text is in the choice and in the fallback', [(editedBox.match(/ZZ/g) ?? []).length, /<wps:txbx><w:txbxContent>[^]*ZZ[^]*<\/wps:txbx>/.test(editedBox), /<v:textbox[^>]*><w:txbxContent>[^]*ZZ/.test(editedBox)], [2, true, true]);
            const againPkg = await OF.processors.forEditorPackage(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer()));
            const againBoxes = againPkg.Paragraphs.filter((item) => item.Type === 'TextBox');
            check('read again: one box, with that text, where it was', [againBoxes.length, againBoxes[0].Paragraphs.length, Math.round(againBoxes[0].X), Math.round(againBoxes[0].Y), againBoxes[0].Width, againBoxes[0].Height], [1, 2, 300, 300, 240, 120]);
            // Undo to the start: the document is the one that was opened again.
            let steps = 0;
            while (Word.canUndo(editor) && steps < 20) { Word.undo(editor); steps++; }
            check('everything undone: the box has the text it had, and the save is the file that was opened', [boxRect().Text, Word.exportDocumentPackage(editor).Unchanged], ['first line\nsecond line', true]);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
