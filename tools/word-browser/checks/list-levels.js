// A list item moved to another level of its list (Tab / Shift+Tab at its
// start, increaseIndent / decreaseIndent / setListLevel) is at once what it
// is after the file is saved and read again: the number in the level's
// format and text ("I.a)" with the number of the level above written as
// that level writes it), the level's indents, the look of its number. The
// reader hands all the levels of a list over with its paragraphs. Saved, the
// paragraph says its new level (w:ilvl) and no indent of its own, and the
// numbering part is the one that was read.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/lists.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('>', original.indexOf('<w:body')) + 1);
    const numberingOriginal = await zip.file('word/numbering.xml').async('string');
    const numberingHead = numberingOriginal.slice(0, numberingOriginal.indexOf('>', numberingOriginal.indexOf('<w:numbering')) + 1);
    const level = (ilvl, inner) => `<w:lvl w:ilvl="${ilvl}">${inner}</w:lvl>`;
    const ind = (left, hanging) => `<w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr>`;
    const numbering = numberingHead
        + '<w:abstractNum w:abstractNumId="0">'
        + level(0, `<w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/>${ind(720, 360)}<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>`)
        + level(1, `<w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/>${ind(1440, 432)}`)
        + level(2, `<w:start w:val="3"/><w:numFmt w:val="decimal"/><w:isLgl/><w:lvlText w:val="%1.%2.%3"/>${ind(2160, 504)}`)
        + '</w:abstractNum>'
        + '<w:abstractNum w:abstractNumId="1">'
        + level(0, `<w:numFmt w:val="bullet"/><w:lvlText w:val=""/>${ind(720, 360)}<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>`)
        + level(1, `<w:numFmt w:val="bullet"/><w:lvlText w:val="o"/>${ind(1440, 360)}<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr>`)
        + '</w:abstractNum>'
        + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        + '<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'
        + '</w:numbering>';
    const item = (numId, ilvl, text) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const bytesFor = async (body) => { zip.file('word/numbering.xml', numbering); zip.file('word/document.xml', head + body + sect + '</w:body></w:document>'); return zip.generateAsync({ type: 'uint8array' }); };
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const open = async (bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes)); return editor; };
    const save = async (editor) => {
        const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const out = await window.JSZip.loadAsync(bytes);
        return { bytes, document: await out.file('word/document.xml').async('string'), numbering: await out.file('word/numbering.xml').async('string') };
    };
    // What a line says of its list item: the number as shown, where the text starts, where the paragraph is indented.
    const items = (editor) => Word.getLayoutLines(editor).filter((line) => line.BlockId.length > 0 && line.StartOffset === 0).map((line) => [line.Text, line.ListMarker, Math.round(line.X * 10) / 10, Math.round(line.Indent * 10) / 10]);
    const caretIn = (editor, text) => { const line = Word.getLayoutLines(editor).find((entry) => entry.Text === text); Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0); return line; };
    const paragraphXml = (xml, text) => { const end = xml.indexOf(`<w:t>${text}</w:t>`); const start = xml.lastIndexOf('<w:p>', end); return xml.slice(start, xml.indexOf('</w:p>', end) + 6); };

    const body = item(1, 0, 'one') + item(1, 0, 'two') + item(1, 0, 'three') + item(2, 0, 'dot') + item(2, 0, 'another') + item(3, 0, 'five') + item(3, 0, 'six');
    const editor = await open(await bytesFor(body));
    const read = items(editor);
    check('read: the numbers and bullets of the first levels', read.map((entry) => entry[1].trim()), ['I.', 'II.', 'III.', '', '', 'V.', 'VI.']);

    // Tab on "two" (through the API, as the toolbar does).
    caretIn(editor, 'two');
    Word.increaseIndent(editor);
    const moved = items(editor);
    const reread = items(await open(await bytesFor(item(1, 0, 'one') + item(1, 1, 'two') + item(1, 0, 'three') + item(2, 0, 'dot') + item(2, 0, 'another') + item(3, 0, 'five') + item(3, 0, 'six'))));
    check('"two" a level down: its number, where its text starts and its indent are those of the file read with it at that level - at once', moved, reread);
    check('the number is the level\'s: I.a), and the item after it is II. again', [moved[1][1].trim(), moved[2][1].trim(), moved[1][2] > read[1][2] + 40], ['I.a)', 'II.', true]);
    check('getSelectionFormat says the level', Word.getSelectionFormat(editor).Paragraph.ListLevel, 2);
    // Saved: w:ilvl 1, no indent of the paragraph's own, and the numbering part as it was read.
    let saved = await save(editor);
    const two = paragraphXml(saved.document, 'two');
    check('saved: the paragraph says its level and nothing of an indent', [/<w:ilvl w:val="1"\/>/.test(two), /<w:numId w:val="1"\/>/.test(two), /<w:ind /.test(two)], [true, true, false]);
    check('the numbering part is the one that was read', saved.numbering === numbering, true);
    const again = await open(saved.bytes);
    check('read again: what the editor showed', items(again), moved);

    // The pixels: the number of the moved item stands further right than the first level's (which is red and right-aligned:
    // it ends where a left-aligned number would start).
    const context = canvas.getContext('2d');
    const firstInk = (y, height) => { const data = context.getImageData(0, Math.round(y), 600, Math.round(height)).data; let first = -1; for (let x = 0; x < 600 && first < 0; x++) for (let row = 0; row < Math.round(height); row++) { const i = (row * 600 + x) * 4; if (data[i + 1] < 140 && data[i + 2] < 140) { first = x; break; } } return first; };
    Word.repaint(editor);
    const lineTwo = Word.getLayoutLines(editor).find((entry) => entry.Text === 'two'), lineOne = Word.getLayoutLines(editor).find((entry) => entry.Text === 'one');
    const inkTwo = firstInk(lineTwo.Y, lineTwo.Height), inkOne = firstInk(lineOne.Y, lineOne.Height);
    note(`the first ink of "one" is at x ${inkOne}, of "two" (a level down) at x ${inkTwo}`);
    check('painted: the number of the item a level down stands right of the first level\'s; the first level\'s is right-aligned (it starts left of its position, 24 px right of the margin)', [inkTwo > inkOne + 20, inkOne < 96 + 24 - 4, inkTwo >= 96 + 96 - 30], [true, true, true]);

    // Tab again: the third level is legal - decimal numbers of all levels - and starts at 3.
    Word.increaseIndent(editor);
    check('a level further down: 1.1.3 (a legal level that starts at 3)', items(editor)[1][1].trim(), '1.1.3');
    Word.setListLevel(editor, 1);
    check('setListLevel 1: as read', items(editor), read);
    Word.undo(editor); Word.undo(editor);
    check('undone twice: at the second level again', items(editor), moved);
    Word.undo(editor);
    check('and once more: as read', items(editor), read);

    // The keyboard: Tab and Shift+Tab at the start of a list item.
    const three = caretIn(editor, 'three');
    const input = Word.getInputElement(editor);
    const key = (name, shift) => input.dispatchEvent(new KeyboardEvent('keydown', { key: name, code: name, shiftKey: !!shift, bubbles: true, cancelable: true }));
    key('Tab');
    check('Tab at the start of "three": II.a)', [items(editor)[2][1].trim(), items(editor)[2][2] > read[2][2] + 40], ['II.a)', true]);
    key('Tab', true);
    check('Shift+Tab: III. again', items(editor), read);

    // A bullet a level down: the second level's bullet at its indent.
    caretIn(editor, 'another');
    Word.increaseIndent(editor);
    const bullet = items(editor)[4];
    check('a bullet a level down: the second level\'s bullet, at its indent', [bullet[1].trim(), bullet[2] > read[4][2] + 40], ['o', true]);
    // A list that starts at 5: its second item a level down.
    caretIn(editor, 'six');
    Word.increaseIndent(editor);
    check('in a list that starts at 5 (w:startOverride): V.a)', items(editor)[6][1].trim(), 'V.a)');
    saved = await save(editor);
    check('saved with three items moved: the numbering part still the one that was read', [saved.numbering === numbering, /<w:ilvl w:val="1"\/>/.test(paragraphXml(saved.document, 'another')), /<w:ilvl w:val="1"\/>/.test(paragraphXml(saved.document, 'six'))], [true, true, true]);
    check('and read again', items(await open(saved.bytes)), items(editor));
    // file-sample.docx: Tab at the start of the list item "Lists" (it used to stay "2." at 48 px until the file was read again).
    {
        const sample = await open(await bytesOf('/repo/tests/test-documents-online/file-sample.docx'));
        const marked = (target) => Word.getLayoutLines(target).filter((line) => line.ListMarker && line.ListMarker.trim() !== '').map((line) => [line.Text.trim(), line.ListMarker.trim(), Math.round(line.Indent)]);
        const line = Word.getLayoutLines(sample).find((entry) => entry.ListMarker && entry.Text.startsWith('Lists'));
        Word.setSelection(sample, line.BlockId, 0, line.BlockId, 0);
        Word.getInputElement(sample).dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true }));
        const format = Word.getSelectionFormat(sample).Paragraph;
        const shown = marked(sample);
        check('file-sample.docx, Tab on "Lists": shown as "a." at the second level\'s indent, the items after it count on', [shown.slice(0, 3), format.ListLevel, format.IndentLeft], [[['Headings', '1.', 48], ['Lists', 'a.', 96], ['Links', '2.', 48]], 2, 96]);
        const written = await save(sample);
        check('and it is what the saved file shows when it is read again', marked(await open(written.bytes)), shown);
    }
} catch (e) { fail(e); }
finish();
