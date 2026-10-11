// The commands a toolbar needs: sizes in points, colours, fonts, styles,
// paragraph spacing, clear formatting, page break, table rows / columns /
// merge / split and hyperlinks. (Picture handles: image-handles.js.)
import { check, note, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initialize('c');
    Word.setParagraphs(editor, ['A paragraph of text to format.', 'A second paragraph.', 'A third one.']);
    const ids = () => [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
    const select = (para, from, to) => Word.setSelection(editor, ids()[para], from, ids()[para], to);
    const format = () => Word.getSelectionFormat(editor);
    const undoSteps = () => { let n = 0; while (Word.undo(editor)) n++; for (let i = 0; i < n; i++) Word.redo(editor); return n; };
    const stepsBefore = undoSteps();

    // ---- font size in points ----
    select(0, 2, 11);
    Word.setFontSizePt(editor, 12);
    check('setFontSizePt(12): 12 pt is 16 layout pixels', [format().FontSizePt, format().FontSize], [12, 16]);
    Word.setFontSizePt(editor, 14);
    check('setFontSizePt(14): 14 pt, 18.67 px', [format().FontSizePt, Math.round(format().FontSize * 100) / 100], [14, 18.67]);
    Word.setFontSize(editor, 24);
    check('setFontSize is still in pixels: 24 px is 18 pt', [format().FontSizePt, format().FontSize], [18, 24]);
    Word.setFontSizePt(editor, 10.5);
    check('half points', format().FontSizePt, 10.5);
    select(0, 0, 20);
    check('a mixed selection says so in both units', [format().FontSizePt, format().FontSize], ['mixed', 'mixed']);

    // ---- colour, highlight, font family ----
    select(0, 2, 11);
    Word.setColor(editor, '#C00000'); Word.setHighlight(editor, 'yellow'); Word.setFontFamily(editor, 'Georgia');
    check('text colour, highlight and font family', [format().Color, format().Highlight, format().FontFamily], ['#C00000', 'yellow', 'Georgia']);
    check('getFontFamilies: the document\'s fonts when no text stack is installed', Word.getFontFamilies(editor).includes('Georgia'), true);
    Word.setFontFamilies(editor, ['Calibri', 'Arial', 'Times New Roman', 'georgia']);
    check('with the text stack\'s families: those and the document\'s, sorted, each once', Word.getFontFamilies(editor), ['Arial', 'Calibri', 'georgia', 'Times New Roman']);
    Word.clearFormatting(editor);
    check('clear formatting', [format().Color, format().Highlight, format().FontFamily !== 'Georgia'], [null, null, true]);

    // ---- styles ----
    check('a document made in the editor has no named styles', Word.getStyles(editor), []);
    {
        const { loadFormats, bytesOf } = await import('./common.js');
        const OF = await loadFormats();
        const styled = Word.initializeCanvas(Object.assign(document.createElement('canvas'), { width: 10, height: 10 }));
        Word.importFromDocxPackage(styled, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents/sample-formatted.docx')));
        const styles = Word.getStyles(styled);
        note('styles ' + JSON.stringify(styles.slice(0, 3)));
        check('getStyles lists ids, names and types', [styles.length > 5, styles.every((style) => typeof style.Id === 'string' && typeof style.Name === 'string' && typeof style.Type === 'string')], [true, true]);
        const heading = styles.find((style) => style.Type === 'paragraph' && style.HeadingLevel === 1);
        check('heading styles say their level', !!heading, true);
        const target = styles.find((style) => style.Type === 'paragraph' && !style.IsDefault && style.HeadingLevel === 2) ?? heading;
        const body = Word.getLayoutLines(styled).find((line) => (line.Text || '').trim().length > 20);
        Word.setSelection(styled, body.BlockId, 0, body.BlockId, 0);
        Word.setParagraphStyle(styled, target.Id);
        check('setParagraphStyle takes an id of getStyles', Word.getSelectionFormat(styled).Paragraph.StyleId, target.Id);
        check('and is one undo step', [Word.undo(styled), Word.getSelectionFormat(styled).Paragraph.StyleId !== target.Id], [true, true]);
        check('getFontFamilies includes the fonts of the document\'s styles and text', Word.getFontFamilies(styled).length > 0, true);
        Word.dispose(styled);
    }

    // ---- spacing ----
    select(1, 0, 0);
    const lineBefore = Word.getLayoutLines(editor).filter((line) => line.BlockId === ids()[2])[0].Y;
    Word.setLineSpacing(editor, 2); Word.setSpacingBefore(editor, 18); Word.setSpacingAfter(editor, 6);
    const pkgPara = (index) => (Word.exportDocumentPackage(editor).Paragraphs ?? [])[index];
    // (Points, as a package holds paragraph spacing - twips / 20 - and as they were given. It used to be 24 and 8: the points made pixels.)
    check('line spacing and space before / after are on the paragraph', [pkgPara(1).LineSpacing, pkgPara(1).SpacingBefore, pkgPara(1).SpacingAfter], [2, 18, 6]);
    check('and move what follows', Word.getLayoutLines(editor).filter((line) => line.BlockId === ids()[2])[0].Y > lineBefore + 20, true);

    // ---- page break ----
    const pagesBefore = Word.getPageCount(editor);
    select(2, 2, 2);
    Word.insertPageBreak(editor);
    check('a page break adds a page and splits the paragraph', [Word.getPageCount(editor), Word.getText(editor).split('\n').slice(2, 4)], [pagesBefore + 1, ['A ', 'third one.']]);
    Word.undo(editor);
    check('and is one undo step', [Word.getPageCount(editor), Word.getText(editor).split('\n')[2]], [pagesBefore, 'A third one.']);

    // ---- tables ----
    select(2, 12, 12);
    Word.insertTable(editor, 2, 2);
    const table = () => (Word.exportDocumentPackage(editor).Paragraphs ?? []).find((item) => item.Type === 'Table');
    const shape = () => table().Rows.map((row) => row.Cells.map((cell) => cell.ColSpan ?? cell.GridSpan ?? 1));
    const cellIds = () => table().Rows.map((row) => row.Cells.map((cell) => cell.Paragraphs[0].Id ?? cell.Paragraphs[0].BlockId));
    check('a 2 x 2 table', shape(), [[1, 1], [1, 1]]);
    // The caret is in the first cell after the table is inserted; the keyboard moves from cell to cell.
    const key = (name, options = {}) => Word.getInputElement(editor).dispatchEvent(new KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true, ...options }));
    const here = () => Word.getSelection(editor).FocusBlockId;
    const firstCell = here();
    check('the caret is in the table\'s first cell', ids().includes(firstCell), false);
    Word.insertText(editor, 'a1');
    Word.insertTableRowAfter(editor); check('row after', shape().length, 3);
    Word.insertTableRowBefore(editor); check('row before: the caret\'s row is now the second', [shape().length, table().Rows[1].Cells[0].Paragraphs[0].Runs?.map((run) => run.Text).join('') ?? ''], [4, 'a1']);
    Word.insertTableColumnAfter(editor); check('column after', shape()[1].length, 3);
    Word.insertTableColumnBefore(editor); check('column before', shape().map((row) => row.length), [4, 4, 4, 4]);
    Word.deleteTableRow(editor); Word.deleteTableColumn(editor);
    check('delete row and column', [shape().length, shape()[0].length], [3, 3]);
    // Merge two neighbouring cells of a row, then split them again.
    check('nothing to merge with a caret in one cell', Word.canMergeTableCells(editor), false);
    const left = here();
    key('End', { ctrlKey: false }); key('ArrowRight');
    const right = here();
    check('ArrowRight at the end of a cell goes to the next cell', right !== left, true);
    Word.setSelection(editor, left, 0, right, 0);
    check('a selection over two cells can be merged', Word.canMergeTableCells(editor), true);
    const rowOfCaret = () => table().Rows.findIndex((row) => row.Cells.some((cell) => (cell.ColSpan ?? 1) > 1));
    Word.mergeTableCells(editor);
    const mergedRow = rowOfCaret();
    check('merged: one cell spanning two columns', [mergedRow >= 0, mergedRow >= 0 && shape()[mergedRow].reduce((a, b) => a + b, 0), mergedRow >= 0 && shape()[mergedRow].length], [true, 3, 2]);
    Word.splitTableCell(editor);
    check('split again', shape().map((row) => row.length), [3, 3, 3]);
    Word.undo(editor);
    check('the split is an undo step', rowOfCaret(), mergedRow);
    Word.undo(editor);
    check('and so is the merge', shape().map((row) => row.length), [3, 3, 3]);
    void cellIds;

    // ---- hyperlinks ----
    select(0, 2, 11);
    Word.insertHyperlink(editor, 'https://example.com/a', '');
    check('the selection becomes a link', [format().Hyperlink, Word.getHyperlink(editor) && [Word.getHyperlink(editor).Target, Word.getHyperlink(editor).StartOffset, Word.getHyperlink(editor).EndOffset, Word.getHyperlink(editor).Text]], ['https://example.com/a', ['https://example.com/a', 2, 11, 'paragraph']]);
    select(0, 5, 5);
    check('with the caret in it', Word.getHyperlink(editor)?.Target, 'https://example.com/a');
    Word.setHyperlinkTarget(editor, 'https://example.com/b');
    check('setHyperlinkTarget changes the whole link', [Word.getHyperlink(editor).Target, Word.getHyperlink(editor).StartOffset, Word.getHyperlink(editor).EndOffset], ['https://example.com/b', 2, 11]);
    select(0, 20, 20);
    check('no link at the caret elsewhere', [Word.getHyperlink(editor), format().Hyperlink], [null, null]);
    Word.insertHyperlink(editor, 'https://example.com/c', 'link text');
    select(0, 22, 22);
    check('a link inserted at the caret, with its text', [Word.getHyperlink(editor)?.Target, Word.getHyperlink(editor)?.Text, Word.getText(editor).split('\n')[0].includes('link text')], ['https://example.com/c', 'link text', true]);
    Word.setHyperlinkTarget(editor, '');
    check('an empty target removes the link and keeps the text', [Word.getHyperlink(editor), Word.getText(editor).split('\n')[0].includes('link text')], [null, true]);

    // (Resizing a picture by its handles: see image-handles.js, which finds the picture in the painted pixels.)
    void canvas; void stepsBefore;
} catch (e) { fail(e); }
finish();
