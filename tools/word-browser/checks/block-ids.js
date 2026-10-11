// No block id twice: a position is a block id and an offset, so two blocks
// with one id are one block to the caret, to the selection and to an edit.
// Every document under tests/ is read (as a whole and front first), and
// getDuplicateBlockIds finds nothing - nor after the edits that make blocks:
// a paragraph split, a row and a column inserted, a table inserted in a cell.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.createElement('canvas');
    const corpus = await (await fetch('/corpus-list')).json();
    const withDuplicates = [];
    let nested = 0;
    for (const path of corpus) {
        let pkg;
        try { pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/' + path)); } catch { continue; }
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, pkg);
        const twice = Word.getDuplicateBlockIds(editor);
        if (twice.length > 0) withDuplicates.push([path, twice.slice(0, 4)]);
        if ((pkg.Paragraphs ?? []).some((item) => item.Type === 'Table' && item.Rows.some((row) => row.Cells.some((cell) => (cell.NestedTables ?? []).length > 0)))) nested++;
    }
    note(`${corpus.length} documents, ${nested} with a table in a table`);
    check('no document has a block id twice', withDuplicates, []);
    check('documents with tables in tables were among them', nested > 3, true);
    // After edits, in a document with a table in a table.
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/tables.docx'))));
    const lineOf = (text) => Word.getLayoutLines(editor).find((item) => item.Text === text);
    const at = (text, offset = 0) => { const line = lineOf(text); Word.setSelection(editor, line.BlockId, offset, line.BlockId, offset); };
    const steps = [];
    at('inner', 2); Word.insertParagraph ? Word.insertParagraph(editor) : Word.insertText(editor, '\n'); steps.push(Word.getDuplicateBlockIds(editor));
    at('before nested'); Word.insertTableRowAfter(editor); steps.push(Word.getDuplicateBlockIds(editor));
    at('before nested'); Word.insertTableColumnAfter(editor); steps.push(Word.getDuplicateBlockIds(editor));
    at('b1'); Word.insertTable(editor, 2, 2); steps.push(Word.getDuplicateBlockIds(editor));
    at('a1'); Word.insertTableRowBefore(editor); steps.push(Word.getDuplicateBlockIds(editor));
    check('none after a split, a row, a column, a table in a cell, a row above', steps, [[], [], [], [], []]);
    // (The row inserted below the row with the nested table has no table in it, and no paragraph of that table's.)
    const pkg = Word.exportDocumentPackage(editor);
    const table = pkg.Paragraphs.find((item) => item.Type === 'Table');
    note('rows: ' + JSON.stringify(table.Rows.map((row) => row.Cells.map((cell) => [cell.Paragraphs.length, (cell.NestedTables ?? []).length]))));
} catch (e) { fail(e); }
finish();
