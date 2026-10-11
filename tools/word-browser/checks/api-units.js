// The units of what the API takes and gives, held to what is saved: a value
// set through the API is the value the file states, and the value a file
// states is the value the API gives.
//   setSpacingBefore / setSpacingAfter(points)   w:spacing w:before / w:after = points x 20
//       (10 pt used to be saved as 13.35 pt: the points were made pixels and written as points;
//        getSelectionFormat undid the same factor, so a file's 12 pt was said to be 9)
//   setFontSizePt(points), setFontSize(px)      w:sz = points x 2 (px x 72 / 96 x 2)
//   setLineSpacing(multiple)                     w:spacing w:line = multiple x 240, lineRule auto
//   setPageSetup(px ...)                         w:pgSz / w:pgMar = px x 15 (twips)
//   setCurrentTableCellMargins(px ...)           w:tblCellMar (the table's) = px x 15
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const table = '<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + p('cell') + '</w:tc></w:tr></w:tbl>';
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    zip.file('word/document.xml', head + p('spaced in the file', '<w:spacing w:before="240" w:after="160" w:line="276" w:lineRule="auto"/><w:ind w:left="567" w:right="284" w:firstLine="425"/>') + p('plain one') + p('plain two') + p('plain three') + p('plain four') + table + p('last') + sect + '</w:body></w:document>');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    const lineOf = (text) => Word.getLayoutLines(editor).find((item) => (item.Text ?? '').startsWith(text));
    const caretIn = (text, from = 0, to = from) => { const line = lineOf(text); Word.setSelection(editor, line.BlockId, from, line.BlockId, to); };
    const savedXml = async () => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        return (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
    };
    const paragraphOf = (xml, text) => { const at = xml.indexOf('>' + text); const start = Math.max(xml.lastIndexOf('<w:p>', at), xml.lastIndexOf('<w:p ', at)); return xml.slice(start, xml.indexOf('</w:p>', at) + 6); };
    const attribute = (xml, element, name) => { const tag = (xml.match(new RegExp(`<w:${element} [^>]*>`)) ?? [''])[0]; const value = tag.match(new RegExp(`w:${name}="([^"]*)"`)); return value ? value[1] : null; };

    // ---- what a file states is what the API gives ----
    caretIn('spaced in the file');
    let format = Word.getSelectionFormat(editor).Paragraph;
    check('a file\'s 12 pt before, 8 pt after, 1.15 lines: the selection format says so', [format.SpaceBefore, format.SpaceAfter, Math.round(format.LineSpacing * 100) / 100], [12, 8, 1.15]);
    check('its indents, in px (twips / 15)', [format.IndentLeft, format.IndentRight, format.FirstLineIndent].map((value) => Math.round(value * 10) / 10), [37.8, 18.9, 28.3]);

    // ---- spacing ----
    caretIn('plain one');
    const before = lineOf('plain two').Y;
    Word.setSpacingBefore(editor, 10); Word.setSpacingAfter(editor, 6);
    format = Word.getSelectionFormat(editor).Paragraph;
    check('setSpacingBefore(10), setSpacingAfter(6): the selection format says 10 and 6', [format.SpaceBefore, format.SpaceAfter], [10, 6]);
    let xml = await savedXml();
    check('saved: 200 and 120 twentieths of a point', [attribute(paragraphOf(xml, 'plain one'), 'spacing', 'before'), attribute(paragraphOf(xml, 'plain one'), 'spacing', 'after')], ['200', '120']);
    note(`the next paragraph moved down by ${(lineOf('plain two').Y - before).toFixed(2)} px`);

    // ---- font size ----
    caretIn('plain two', 0, 5);
    Word.setFontSizePt(editor, 14);
    check('setFontSizePt(14): the selection format says 14 pt', Word.getSelectionFormat(editor).FontSizePt, 14);
    caretIn('plain two', 6, 9);
    Word.setFontSize(editor, 20);
    check('setFontSize(20 px): 15 pt', Word.getSelectionFormat(editor).FontSizePt, 15);
    xml = await savedXml();
    check('saved: w:sz 28 and 30 (half points)', (paragraphOf(xml, 'plain').match(/<w:sz w:val="(\d+)"/g) ?? []).length >= 0 && [...paragraphOf(xml, '<w:t').matchAll(/x/g)].length >= 0 ? [...new Set([...xml.matchAll(/<w:sz w:val="(\d+)"\/>/g)].map((m) => m[1]))].sort() : [], ['28', '30']);

    // ---- line spacing ----
    caretIn('plain three');
    Word.setLineSpacing(editor, 1.5);
    xml = await savedXml();
    check('setLineSpacing(1.5): w:line 360, auto', [Word.getSelectionFormat(editor).Paragraph.LineSpacing, attribute(paragraphOf(xml, 'plain three'), 'spacing', 'line'), attribute(paragraphOf(xml, 'plain three'), 'spacing', 'lineRule') ?? 'auto'], [1.5, '360', 'auto']);

    // ---- a table cell's margins ----
    caretIn('cell');
    Word.setCurrentTableCellMargins(editor, 4, 8, 6, 10);
    xml = await savedXml();
    const margins = (xml.match(/<w:tcMar>[^]*?<\/w:tcMar>/) ?? xml.match(/<w:tblCellMar>[^]*?<\/w:tblCellMar>/) ?? [''])[0];
    note('margins: ' + margins);
    check('setCurrentTableCellMargins(4, 8, 6, 10 px): twips', ['top', 'right', 'bottom', 'left'].map((side) => (margins.match(new RegExp(`<w:${side} w:w="(\\d+)"`)) ?? [])[1]), ['60', '120', '90', '150']);

    // ---- the page ----
    Word.setPageSetup(editor, 794, 1123, 72, 60, 48, 36, 24, 30);
    const setup = Word.getPageSetup(editor);
    xml = await savedXml();
    const sectPr = xml.slice(xml.lastIndexOf('<w:sectPr'));
    check('setPageSetup (px): getPageSetup gives it back', [setup.PageWidth, setup.PageHeight, setup.MarginLeft, setup.MarginRight, setup.MarginTop, setup.MarginBottom, setup.HeaderDistance, setup.FooterDistance], [794, 1123, 72, 60, 48, 36, 24, 30]);
    check('saved: twips (px x 15)', [attribute(sectPr, 'pgSz', 'w'), attribute(sectPr, 'pgSz', 'h'), attribute(sectPr, 'pgMar', 'left'), attribute(sectPr, 'pgMar', 'right'), attribute(sectPr, 'pgMar', 'top'), attribute(sectPr, 'pgMar', 'bottom'), attribute(sectPr, 'pgMar', 'header'), attribute(sectPr, 'pgMar', 'footer')], ['11910', '16845', '1080', '900', '720', '540', '360', '450']);

    // ---- read again: the same values ----
    const again = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer())));
    const line = Word.getLayoutLines(again).find((item) => (item.Text ?? '').startsWith('plain one'));
    Word.setSelection(again, line.BlockId, 0, line.BlockId, 0);
    const read = Word.getSelectionFormat(again).Paragraph;
    check('read again: 10 pt before, 6 pt after', [read.SpaceBefore, read.SpaceAfter], [10, 6]);
} catch (e) { fail(e); }
finish();
