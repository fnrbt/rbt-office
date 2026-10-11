// A row or a column inserted into a table of a file is saved like the one it
// was inserted next to: the new cells have the neighbour's w:tcPr (width,
// borders, shading, vertical alignment, margins), the new row its w:trPr, each
// new cell a paragraph with the neighbour's paragraph properties - they used
// to be bare <w:tc><w:p/></w:tc> - and w:tblGrid has the columns the table
// has (a column more after an insert, one less after a delete). The rows and
// cells that were there are the XML that was read.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text) => `<w:p><w:pPr><w:spacing w:before="40" w:after="20"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const tcPr = (width, fill) => `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/><w:bottom w:val="double" w:sz="6" w:space="0" w:color="0000FF"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/><w:tcMar><w:left w:w="200" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>`;
    const cell = (width, fill, text) => `<w:tc>${tcPr(width, fill)}${p(text)}</w:tc>`;
    const table = '<w:tbl><w:tblPr><w:tblW w:w="4500" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>'
        + `<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>${cell(1500, 'FFEEDD', 'a1')}${cell(3000, 'DDEEFF', 'b1')}</w:tr>`
        + `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="900" w:hRule="atLeast"/></w:trPr>${cell(1500, 'EEFFDD', 'a2')}${cell(3000, 'DDFFEE', 'b2')}</w:tr>`
        + '</w:tbl>';
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    zip.file('word/document.xml', head + p('before') + table + p('after') + sect + '</w:body></w:document>');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const canvas = document.getElementById('c');
    const open = async () => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes)); return editor; };
    const caretIn = (editor, text) => { const line = Word.getLayoutLines(editor).find((item) => item.Text === text); Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0); };
    const saved = async (editor) => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const tableXml = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>') + 8);
        const rows = tableXml.match(/<w:tr[ >][^]*?<\/w:tr>/g) ?? [];
        return { xml, tableXml, grid: (tableXml.match(/<w:gridCol w:w="(\d+)"/g) ?? []).map((item) => Number(item.match(/\d+/)[0])), rows: rows.map((row) => ({ xml: row, cells: row.match(/<w:tc>[^]*?<\/w:tc>/g) ?? [] })), bytes: out };
    };
    // What of a cell's XML says how it looks, and of its paragraph.
    const cellLook = (xml) => ({
        width: (xml.match(/<w:tcW w:w="(\d+)"/) ?? [])[1], top: /<w:top w:val="single" w:sz="12"[^>]*w:color="FF0000"/.test(xml), bottom: /<w:bottom w:val="double" w:sz="6"[^>]*w:color="0000FF"/.test(xml),
        fill: (xml.match(/<w:shd [^>]*w:fill="([0-9A-F]+)"/) ?? [])[1], left: (xml.match(/<w:tcMar>[^]*?<w:left w:w="(\d+)"/) ?? [])[1], vAlign: (xml.match(/<w:vAlign w:val="(\w+)"/) ?? [])[1],
        centred: /<w:jc w:val="center"\/>/.test(xml), spacing: (xml.match(/<w:spacing [^>]*w:before="(\d+)"[^>]*w:after="(\d+)"/) ?? []).slice(1).join('/'), text: (xml.match(/<w:t[ >][^<]*<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join(''),
    });
    const like = (width, fill, text) => ({ width: String(width), top: true, bottom: true, fill, left: '200', vAlign: 'center', centred: true, spacing: '40/20', text });

    // ---- a row below the second ----
    {
        const editor = await open();
        caretIn(editor, 'a2');
        Word.insertTableRowAfter(editor);
        const out = await saved(editor);
        check('three rows; the grid as it was', [out.rows.length, out.grid], [3, [1500, 3000]]);
        check('the new row\'s cells are like those above them, empty', out.rows[2].cells.map(cellLook), [like(1500, 'EEFFDD', ''), like(3000, 'DDFFEE', '')]);
        check('the new row has the row properties of its neighbour', [/<w:trPr>[^]*<w:cantSplit\/>[^]*<\/w:trPr>/.test(out.rows[2].xml), /<w:trHeight w:val="900" w:hRule="atLeast"\/>/.test(out.rows[2].xml)], [true, true]);
        check('the rows that were there are the XML that was read', [out.rows[0].xml, out.rows[1].xml], [table.match(/<w:tr>[^]*?<\/w:tr>/g)[0], table.match(/<w:tr>[^]*?<\/w:tr>/g)[1]]);
        // In the editor: the new row is as high as its neighbour, its cells shaded.
        const again = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out.bytes));
        const pkg = Word.exportDocumentPackage(again);
        const item = pkg.Paragraphs.find((entry) => entry.Type === 'Table');
        check('read again: three rows of two cells, the last like the second', [item.Rows.length, item.Rows[2].Cells.map((entry) => entry.BackgroundColor), item.Rows[2].Cells.map((entry) => entry.VerticalAlignment)], [3, item.Rows[1].Cells.map((entry) => entry.BackgroundColor), ['center', 'center']]);
    }
    // ---- a row above the first ----
    {
        const editor = await open();
        caretIn(editor, 'b1');
        Word.insertTableRowBefore(editor);
        const out = await saved(editor);
        check('a row above the first: like the first', [out.rows.length, out.rows[0].cells.map(cellLook), /<w:trHeight w:val="600" w:hRule="atLeast"\/>/.test(out.rows[0].xml)], [3, [like(1500, 'FFEEDD', ''), like(3000, 'DDEEFF', '')], true]);
    }
    // ---- a column right of the first ----
    {
        const editor = await open();
        caretIn(editor, 'a1');
        Word.insertTableColumnAfter(editor);
        const out = await saved(editor);
        check('a column right of the first: the grid has a column more, as wide as it', out.grid, [1500, 1500, 3000]);
        check('every row has three cells: the cell that was, one like it, the other that was', out.rows.map((row) => row.cells.map(cellLook)), [[like(1500, 'FFEEDD', 'a1'), like(1500, 'FFEEDD', ''), like(3000, 'DDEEFF', 'b1')], [like(1500, 'EEFFDD', 'a2'), like(1500, 'EEFFDD', ''), like(3000, 'DDFFEE', 'b2')]]);
        check('the rows keep their properties', [/<w:trHeight w:val="600"/.test(out.rows[0].xml), /<w:cantSplit\/>/.test(out.rows[1].xml)], [true, true]);
        // And taken out again: the grid that was.
        // (The new cell: 1500 twips = 100 px right of the first, as wide as it. A point in its middle is in its paragraph.)
        const first = Word.getLayoutLines(editor).find((item) => item.Text === 'a1'), frame = Word.getPageFrames(editor)[0];
        const inNew = Word.hitTestPosition(editor, frame.MarginLeft + 150, first.Y + first.Height / 2);
        check('a point in the new cell is in a paragraph of its own', [inNew !== null, inNew.BlockId !== first.BlockId, Word.setSelection(editor, inNew.BlockId, 0, inNew.BlockId, 0)], [true, true, true]);
        Word.deleteTableColumn(editor);
        const back = await saved(editor);
        check('the column deleted again: two columns, the grid that was', [back.grid, back.rows.map((row) => row.cells.map((item) => cellLook(item).text))], [[1500, 3000], [['a1', 'b1'], ['a2', 'b2']]]);
    }
    // ---- a column left of the second; a column deleted ----
    {
        const editor = await open();
        caretIn(editor, 'b2');
        Word.insertTableColumnBefore(editor);
        const out = await saved(editor);
        check('a column left of the second: as wide as the second, its cells like the second\'s', [out.grid, out.rows.map((row) => cellLook(row.cells[1]))], [[1500, 3000, 3000], [like(3000, 'DDEEFF', ''), like(3000, 'DDFFEE', '')]]);
        const other = await open();
        caretIn(other, 'a1');
        Word.deleteTableColumn(other);
        const less = await saved(other);
        check('the first column deleted: the grid has the second', [less.grid, less.rows.map((row) => row.cells.map((item) => cellLook(item).text))], [[3000], [['b1'], ['b2']]]);
    }
} catch (e) { fail(e); }
finish();
