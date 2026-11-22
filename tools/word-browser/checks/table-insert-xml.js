// A row or column inserted into a table of a file takes its properties from
// the cell at its insertion position: after the first takes the following
// cell, while appending takes the last. New cells copy that cell's w:tcPr
// (width, borders, shading, alignment and margins) and paragraph properties;
// new rows copy their neighbour's w:trPr. Original cells keep their XML.
// The grid and a fixed table's absolute preferred width gain or lose the
// inserted/deleted column's width.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const { getTableGeometry } = await import('/word/Word.js');
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
    const open = async (data = bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data)); return editor; };
    const caretIn = (editor, text) => { const line = Word.getLayoutLines(editor).find((item) => item.Text === text); Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0); };
    const tableCellIds = (editor) => {
        const order = Word.getParagraphOrder(editor), lines = Word.getLayoutLines(editor);
        const row = (first, last) => order.slice(order.indexOf(lines.find((item) => item.Text === first).BlockId), order.indexOf(lines.find((item) => item.Text === last).BlockId) + 1);
        return [row('a1', 'b1'), row('a2', 'b2')];
    };
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
        check('a column right of the first: the grid has the following column twice', out.grid, [1500, 3000, 3000]);
        check('every row has three cells: original first, one like the following cell, original second', out.rows.map((row) => row.cells.map(cellLook)), [[like(1500, 'FFEEDD', 'a1'), like(3000, 'DDEEFF', ''), like(3000, 'DDEEFF', 'b1')], [like(1500, 'EEFFDD', 'a2'), like(3000, 'DDFFEE', ''), like(3000, 'DDFFEE', 'b2')]]);
        check('the rows keep their properties', [/<w:trHeight w:val="600"/.test(out.rows[0].xml), /<w:cantSplit\/>/.test(out.rows[1].xml)], [true, true]);
        check('insertion widens the fixed table by the inserted column', /<w:tblW w:w="7500" w:type="dxa"\/>/.test(out.tableXml), true);
        // The fixed table grows with the inserted column. Its rendered
        // positions need not follow the original grid or page margin; choose
        // interior points from the actual layout.
        const ids = tableCellIds(editor);
        const lines = Word.getLayoutLines(editor);
        const firstLine = lines.find((item) => item.BlockId === ids[0][0]);
        const cells = getTableGeometry(editor).filter((item) => item.Kind === 'table-cell' && item.Table === 1 && item.Row === 1)
            .sort((a, b) => a.Cell - b.Cell);
        check('three first-row cells have positive rendered widths', [cells.length, cells.every((cell) => cell.Width > 0)], [3, true]);
        const hits = cells.map((cell) => {
            return Word.hitTestPosition(editor, cell.X + cell.Width / 2, firstLine.BlockY + firstLine.BlockHeight / 2)?.BlockId ?? null;
        });
        check('interior points of all three rendered cells hit their own paragraphs', hits, ids[0]);
        check('the inserted cell has a paragraph distinct from both neighbours', [ids[0][1] !== ids[0][0], ids[0][1] !== ids[0][2]], [true, true]);
        Word.setSelection(editor, hits[1], 0, hits[1], 0);
        Word.deleteTableColumn(editor);
        const back = await saved(editor);
        check('delete by hit test removes only the inserted column', [back.grid, back.rows.map((row) => row.cells.map((item) => cellLook(item).text))], [[1500, 3000], [['a1', 'b1'], ['a2', 'b2']]]);
        check('deletion restores the original fixed table width', /<w:tblW w:w="4500" w:type="dxa"\/>/.test(back.tableXml), true);
        // Independently select the inserted paragraph by identity. This
        // separates a column mutation error from a coordinate/hit-test error.
        const direct = await open();
        caretIn(direct, 'a1');
        Word.insertTableColumnAfter(direct);
        const insertedId = tableCellIds(direct)[0][1];
        Word.setSelection(direct, insertedId, 0, insertedId, 0);
        Word.deleteTableColumn(direct);
        const directBack = await saved(direct);
        check('delete by inserted paragraph identity removes only the inserted column', [directBack.grid, directBack.rows.map((row) => row.cells.map((item) => cellLook(item).text))], [[1500, 3000], [['a1', 'b1'], ['a2', 'b2']]]);
    }
    // ---- a column left of the second; a column deleted ----
    {
        const editor = await open();
        caretIn(editor, 'b2');
        Word.insertTableColumnBefore(editor);
        const out = await saved(editor);
        check('a column left of the second: as wide as the second, its cells like the second\'s', [out.grid, out.rows.map((row) => cellLook(row.cells[1]))], [[1500, 3000, 3000], [like(3000, 'DDEEFF', ''), like(3000, 'DDFFEE', '')]]);
        check('insertion before the second widens the fixed table', /<w:tblW w:w="7500" w:type="dxa"\/>/.test(out.tableXml), true);
        const other = await open();
        caretIn(other, 'a1');
        Word.deleteTableColumn(other);
        const less = await saved(other);
        check('the first column deleted: the grid has the second', [less.grid, less.rows.map((row) => row.cells.map((item) => cellLook(item).text))], [[3000], [['b1'], ['b2']]]);
        check('deleting the first shrinks the fixed table by its width', /<w:tblW w:w="3000" w:type="dxa"\/>/.test(less.tableXml), true);
    }
    // ---- a fixed width different from the input grid's sum ----
    {
        const otherZip = await window.JSZip.loadAsync(bytes);
        otherZip.file('word/document.xml', head + p('before') + table.replaceAll('3000', '1500') + p('after') + sect + '</w:body></w:document>');
        const mismatch = await otherZip.generateAsync({ type: 'uint8array' });
        const editor = await open(mismatch);
        caretIn(editor, 'a1');
        Word.insertTableColumnAfter(editor);
        const out = await saved(editor);
        // Word normalizes 1500/1500 to the 4500-twip fixed table's
        // rendered 2250/2250, then adds a 2250-twip following column.
        check('fixed width differing from grid: insertion uses resolved widths', [out.grid, out.rows.map((row) => row.cells.map((item) => cellLook(item).width)), /<w:tblW w:w="6750" w:type="dxa"\/>/.test(out.tableXml)], [[2250, 2250, 2250], [['2250', '2250', '2250'], ['2250', '2250', '2250']], true]);
        const again = await open(out.bytes);
        const againIds = tableCellIds(again)[0];
        const againLines = Word.getLayoutLines(again);
        const againLine = againLines.find((item) => item.BlockId === againIds[0]);
        const againCells = getTableGeometry(again).filter((item) => item.Kind === 'table-cell' && item.Table === 1 && item.Row === 1)
            .sort((a, b) => a.Cell - b.Cell);
        check('read again: normalized fixed widths are the rendered cell widths', againCells.map((cell) => Math.abs(cell.Width - 150) < 1), [true, true, true]);
        check('read again: all three cells hit their own paragraphs', againCells.map((cell) => {
            return Word.hitTestPosition(again, cell.X + cell.Width / 2, againLine.BlockY + againLine.BlockHeight / 2)?.BlockId ?? null;
        }), againIds);
        Word.setSelection(again, againIds[1], 0, againIds[1], 0);
        Word.deleteTableColumn(again);
        const back = await saved(again);
        check('deleting the added column preserves normalized widths and restores fixed table width', [back.grid, back.rows.map((row) => row.cells.map((item) => cellLook(item).width)), /<w:tblW w:w="4500" w:type="dxa"\/>/.test(back.tableXml)], [[2250, 2250], [['2250', '2250'], ['2250', '2250']], true]);
    }
} catch (e) { fail(e); }
finish();
