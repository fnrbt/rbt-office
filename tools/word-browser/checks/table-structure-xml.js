// A row or a column put into a table of a file, or taken out of it, is saved as the table's XML as it was read
// with that change made in it, byte for byte: the rows and cells that were there are their bytes; a new row is
// made like the row it was put in next to - that row's start tag without its w14 ids, its w:trPr, its cells'
// w:tcPr, and in each cell an empty paragraph made like the cell's first (its start tag without its w14 ids, its
// w:pPr: the paragraph's and its mark's properties); a new cell likewise (one column wide where its neighbour
// spans two: w:gridSpan taken out); w:tblGrid has a w:gridCol put in like its neighbour's or taken out. A row or a
// cell taken out leaves the marks of the ranges in it (bookmarks, permission ranges) where it stood, and takes a
// content control around it that holds nothing else with it. The table stands after a bookmark's end between
// blocks, which used to make the whole table be written from the model.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    let head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    if (!head.includes('xmlns:w14=')) head = head.replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ');
    // A paragraph with a w14 id, rsids, properties the model does not write (w:suppressAutoHyphens) and a mark's
    // properties; marks of ranges may stand before and after its run.
    const pPr = '<w:pPr><w:suppressAutoHyphens/><w:spacing w:before="40" w:after="20"/><w:jc w:val="center"/><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr></w:pPr>';
    const p = (id, text, before = '', after = '') => `<w:p w14:paraId="${id}" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3">${pPr}${before}<w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>${text}</w:t></w:r>${after}</w:p>`;
    const tcPr = (width, fill, span = '') => `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${span}<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/><w:noWrap/><w:hideMark/></w:tcPr>`;
    const tc = (width, fill, para, span = '') => `<w:tc>${tcPr(width, fill, span)}${para}</w:tc>`;
    const trPr = (height) => `<w:trPr><w:cnfStyle w:val="000000100000"/><w:trHeight w:val="${height}" w:hRule="atLeast"/></w:trPr>`;
    const tr = (id, height, cells) => `<w:tr w:rsidR="00D1E2F3" w14:paraId="${id}" w14:textId="77777777">${trPr(height)}${cells}</w:tr>`;
    const a1 = tc(1500, 'FFEEDD', p('10000001', 'a1')), b1 = tc(3000, 'DDEEFF', p('10000002', 'b1'));
    const a2 = tc(1500, 'EEFFDD', p('10000003', 'a2', '<w:bookmarkStart w:id="12" w:name="cross"/><w:bookmarkStart w:id="11" w:name="inner"/>', '<w:bookmarkEnd w:id="11"/>'));
    const b2 = tc(3000, 'DDFFEE', p('10000004', 'b2', '<w:permStart w:id="13" w:edGrp="everyone"/>', '<w:permEnd w:id="13"/>'));
    const a3 = tc(1500, 'FFFFDD', p('10000005', 'a3', '', '<w:bookmarkEnd w:id="12"/>')), b3 = tc(3000, 'DDDDFF', p('10000006', 'b3'));
    const c4 = tc(4500, 'EEEEEE', p('10000007', 'c4'), '<w:gridSpan w:val="2"/>');
    const r1 = tr('20000001', 600, a1 + b1), r2 = tr('20000002', 700, a2 + b2), r3 = tr('20000003', 800, a3 + b3), r4 = tr('20000004', 900, c4);
    // (A row in a content control: a repeating section's item.)
    const e5 = tc(1500, 'DDDDDD', p('10000008', 'e5')), f5 = tc(3000, 'CCCCCC', p('10000009', 'f5'));
    const r5 = tr('20000005', 1000, e5 + f5);
    const item = `<w:sdt><w:sdtPr><w:alias w:val="item"/><w:id w:val="4242"/></w:sdtPr><w:sdtContent>${r5}</w:sdtContent></w:sdt>`;
    const grid = '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>';
    const table = '<w:tbl>\n  <w:tblPr><w:tblW w:w="4500" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCaption w:val="kept"/></w:tblPr>\n  ' + grid
        + '\n  ' + r1 + '\n  ' + r2 + '\n  ' + r3 + '\n  ' + r4 + '\n  ' + item + '\n</w:tbl>';
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    const body = '<w:p><w:bookmarkStart w:id="1" w:name="around"/><w:r><w:t>before</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>' + table + '<w:p><w:r><w:t>after</w:t></w:r></w:p>' + sect;
    const documentXml = head + body + '</w:body></w:document>';
    zip.file('word/document.xml', documentXml);
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const canvas = document.getElementById('c');
    const open = async (data) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data)); return editor; };
    const caretIn = (editor, text) => { const line = Word.getLayoutLines(editor).find((item) => item.Text === text); Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0); };
    const save = async (editor) => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        return { xml: await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string'), bytes: out, report: OF.word.getLastSaveReport() };
    };
    // What is made like an element that was read: its start tag without its w14 ids.
    const withoutIds = (startTag) => startTag.replace(/ w14:(paraId|textId)="[^"]*"/g, '');
    const startTagOf = (xml) => xml.slice(0, xml.indexOf('>') + 1);
    const emptyLike = (paragraph) => withoutIds(startTagOf(paragraph)) + pPr + '</w:p>';
    const cellLike = (cell, tcPrNow) => '<w:tc>' + (tcPrNow ?? cell.match(/<w:tcPr>[^]*?<\/w:tcPr>/)[0]) + emptyLike(cell.match(/<w:p [^]*?<\/w:p>/)[0]) + '</w:tc>';
    const rowLike = (row) => withoutIds(startTagOf(row)) + row.match(/<w:trPr>[^]*?<\/w:trPr>/)[0] + (row.match(/<w:tc>[^]*?<\/w:tc>/g) ?? []).map((cell) => cellLike(cell)).join('') + '</w:tr>';
    // (Where the saved part first differs from the one expected, to say so.)
    const firstDifference = (got, want) => { let at = 0; while (at < got.length && got[at] === want[at]) at++; return at >= got.length && at >= want.length ? 'the same' : `at ${at}: ${JSON.stringify(got.slice(Math.max(0, at - 80), at + 120))} where ${JSON.stringify(want.slice(Math.max(0, at - 80), at + 120))}`; };
    const textOfTable = async (data) => {
        const pkg = await OF.processors.forEditorPackage(data);
        const item = pkg.Paragraphs.find((entry) => entry.Type === 'Table');
        return item.Rows.map((row) => row.Cells.map((cell) => cell.Paragraphs.map((para) => para.Runs.map((run) => run.Text).join('')).join('/')));
    };

    // ---- a row below the first ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'a1');
        Word.insertTableRowAfter(editor);
        const out = await save(editor);
        check('a row below the first: the document is the one read with a row like the first put in after it', firstDifference(out.xml, documentXml.replace(r1, r1 + rowLike(r1))), 'the same');
        check('a row below the first: no paragraph written from the model', out.report.GeneratedParagraphs, 0);
        check('read again: the new row is empty', await textOfTable(out.bytes), [['a1', 'b1'], ['', ''], ['a2', 'b2'], ['a3', 'b3'], ['c4'], ['e5', 'f5']]);
    }
    // ---- a row above the third ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'b3');
        Word.insertTableRowBefore(editor);
        const out = await save(editor);
        check('a row above the third: put in before it, like it', firstDifference(out.xml, documentXml.replace(r3, rowLike(r3) + r3)), 'the same');
        check('a row above the third: no paragraph written from the model', out.report.GeneratedParagraphs, 0);
    }
    // ---- the second row deleted: its marks stay where it stood ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'a2');
        Word.deleteTableRow(editor);
        const out = await save(editor);
        const marks = '<w:bookmarkStart w:id="12" w:name="cross"/><w:bookmarkStart w:id="11" w:name="inner"/><w:bookmarkEnd w:id="11"/><w:permStart w:id="13" w:edGrp="everyone"/><w:permEnd w:id="13"/>';
        check('the second row deleted: the document is the one read with the row\'s marks in its place', firstDifference(out.xml, documentXml.replace(r2, marks)), 'the same');
        check('read again: the rows that are left', await textOfTable(out.bytes), [['a1', 'b1'], ['a3', 'b3'], ['c4'], ['e5', 'f5']]);
    }
    // ---- the row in a content control deleted: the control, which holds nothing else, goes with it ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'f5');
        Word.deleteTableRow(editor);
        const out = await save(editor);
        check('the row in a content control deleted: the document is the one read without the control', firstDifference(out.xml, documentXml.replace('\n  ' + item, '')), 'the same');
        check('read again: the rows that are left', await textOfTable(out.bytes), [['a1', 'b1'], ['a2', 'b2'], ['a3', 'b3'], ['c4']]);
    }
    // ---- a column right of the first ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'a1');
        Word.insertTableColumnAfter(editor);
        const out = await save(editor);
        const want = documentXml
            .replace(grid, '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>')
            .replace(a1, a1 + cellLike(a1)).replace(a2, a2 + cellLike(a2)).replace(a3, a3 + cellLike(a3))
            // (Next to the cell that spans both columns: one column wide.)
            .replace(c4, c4 + cellLike(c4, tcPr(4500, 'EEEEEE')))
            .replace(e5, e5 + cellLike(e5));
        check('a column right of the first: every row has a cell like its first after it, the grid a column like the first', firstDifference(out.xml, want), 'the same');
        check('a column right of the first: no paragraph written from the model', out.report.GeneratedParagraphs, 0);
        check('read again: the new column is empty', await textOfTable(out.bytes), [['a1', '', 'b1'], ['a2', '', 'b2'], ['a3', '', 'b3'], ['c4', ''], ['e5', '', 'f5']]);
    }
    // ---- a column left of the second ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'b1');
        Word.insertTableColumnBefore(editor);
        const out = await save(editor);
        // (The row of the cell that spans both columns has no second cell: nothing is put into it.)
        const want = documentXml
            .replace(grid, '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>')
            .replace(b1, cellLike(b1) + b1).replace(b2, cellLike(b2) + b2).replace(b3, cellLike(b3) + b3).replace(f5, cellLike(f5) + f5);
        check('a column left of the second: every row has a cell like its second before it, the grid a column as wide', firstDifference(out.xml, want), 'the same');
        check('a column left of the second: no paragraph written from the model', out.report.GeneratedParagraphs, 0);
    }
    // ---- the second column deleted ----
    {
        const editor = await open(bytes);
        caretIn(editor, 'b1');
        Word.deleteTableColumn(editor);
        const out = await save(editor);
        const want = documentXml
            .replace(grid, '<w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>')
            .replace(b1, '').replace(b2, '<w:permStart w:id="13" w:edGrp="everyone"/><w:permEnd w:id="13"/>').replace(b3, '').replace(f5, '');
        check('the second column deleted: its cells taken out (the marks of one stay where it stood), the grid without its column', firstDifference(out.xml, want), 'the same');
        check('read again: the first column', await textOfTable(out.bytes), [['a1'], ['a2'], ['a3'], ['c4'], ['e5']]);
    }
} catch (e) { fail(e); }
finish();
