// Headers and footers edited in place, the cases beside the default header of
// a plain document: undo / redo of "link to previous" keeps the caret in the
// header (landscape-section.docx); the page count and the page number in a
// header are those of the pages the document shows, the page of the endnotes
// included, while editing and after saving and reading again
// (sample-formatted.docx: 10 pages, 9 of body text); a header's text has the
// size the size box says (the default paragraph style's, as the body);
// entering a header changes nothing - a page without one, a title page and an
// even page show a blank one until something is typed, which undo takes away
// again; a copy made by unlinking is the section's own (tests/word/header-footer.fsx holds the ids of its blocks to be new); first and even footers;
// the same without a viewport (whole-canvas mode); and a long document laid
// out progressively shows the final page count once it is laid out.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const viewport = (top) => ({ scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
    const inkOf = (data, limit = 100) => { let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < limit && data[i + 1] < limit && data[i + 2] < limit) ink++; return ink; };
    // The rows of a region that hold dark pixels: [first, last] (null when none).
    const inkRows = (context, x, y, w, h, limit = 110) => {
        const data = context.getImageData(x, y, w, h).data; let first = -1, last = -1;
        for (let row = 0; row < h; row++) for (let col = 0; col < w; col++) { const i = (row * w + col) * 4; if (data[i] < limit && data[i + 1] < limit && data[i + 2] < limit) { if (first < 0) first = row; last = row; break; } }
        return first < 0 ? null : [first, last];
    };
    const partTexts = (parts) => parts.map((part) => (part.Paragraphs ?? []).map((para) => (para.Runs ?? []).map((run) => run.Text ?? '').join('')).join('|'));
    const story = (editor) => { const c = Word.getCaretContext(editor); return c.HeaderFooter ? [c.Story, c.HeaderFooter.Variant, c.HeaderFooter.PageIndex] : [c.Story]; };
    const save = async (editor) => new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const open = async (bytes, withViewport = true, target = document.createElement('canvas')) => {
        const editor = Word.initializeCanvas(target);
        if (withViewport) Word.setViewport(editor, viewport(0));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        return editor;
    };

    // ---- 1. undo of "link to previous" stays in the header ----
    {
        const editor = await open(await bytesOf('/repo/tests/repros/word/landscape-section.docx'));
        const pages = Word.getPageCount(editor);
        // (Two pages, as the document says of itself: "Landscape section (page 1)", "Portrait section (page 2)".)
        check('landscape-section: the header of page 2', [pages, Word.enterHeaderFooter(editor, 1, 'header'), story(editor)[0], Word.getCaretContext(editor).HeaderFooter.CanLinkToPrevious], [2, true, 'header', true]);
        const linkedBefore = Word.getHeaderFooterLinkToPrevious(editor);
        const textBefore = Word.getText(editor);
        check('link to previous is switched', [Word.setHeaderFooterLinkToPrevious(editor, !linkedBefore), Word.getHeaderFooterLinkToPrevious(editor), story(editor)[0]], [true, !linkedBefore, 'header']);
        check('undo: the link is as it was and the caret is still in the header (it was in the body)', [Word.undo(editor), Word.getCaretContext(editor).Story, Word.getHeaderFooterLinkToPrevious(editor), story(editor)[2], Word.getText(editor) === textBefore], [true, 'header', linkedBefore, 1, true]);
        check('redo: switched again, in the header', [Word.redo(editor), Word.getCaretContext(editor).Story, Word.getHeaderFooterLinkToPrevious(editor)], [true, 'header', !linkedBefore]);
        check('undo: nothing left to undo, not modified', [Word.undo(editor), Word.canUndo(editor), Word.isModified(editor), Word.getCaretContext(editor).Story], [true, false, false, 'header']);
        // A body edit, then a header edit; undo from the body goes to the header, undo there to the body at its edit.
        Word.exitHeaderFooter(editor);
        Word.insertText(editor, 'BODY');
        const bodyText = Word.getText(editor);
        const bodyCaret = Word.getSelection(editor);
        Word.enterHeaderFooter(editor, 0, 'header');
        Word.insertText(editor, 'HEAD');
        Word.exitHeaderFooter(editor);
        check('undo of a header edit from the body: the caret is in that header', [Word.undo(editor), story(editor), Word.getText(editor).includes('HEAD')], [true, ['header', 'default', 0], false]);
        check('undo of the body edit from the header: back in the body, at the edit', [Word.undo(editor), story(editor), Word.getText(editor) === bodyText, Word.getText(editor).includes('BODY')], [true, ['body'], false, false]);
        check('redo: the body edit in the body (the caret where it was), then the header edit in the header', [Word.redo(editor), story(editor), JSON.stringify(Word.getSelection(editor)) === JSON.stringify(bodyCaret), Word.redo(editor), story(editor), Word.getText(editor).includes('HEAD')], [true, ['body'], true, true, ['header', 'default', 0], true]);
        Word.dispose?.(editor);
    }

    // ---- 2 / 3 / 4. sample-formatted.docx: fields, sizes, entering changes nothing ----
    const sampleBytes = await bytesOf('/repo/tests/test-documents/sample-formatted.docx');
    const fieldsAndSizes = async (withViewport, label) => {
        const target = withViewport ? canvas : document.createElement('canvas');
        if (!withViewport) { target.width = 1000; target.height = 12000; }
        const editor = await open(sampleBytes, withViewport, target);
        const context = target.getContext('2d');
        const pageCount = Word.getPageCount(editor);
        const frames = Word.getPageFrames(editor);
        check(`${label}: ten pages (nine of body text and the page of the endnotes)`, [pageCount, frames.length], [10, 10]);
        // The body's size at the start of the text.
        // (The document's default paragraph style - Normal - is 12 pt Ubuntu; the document defaults are 11 pt.)
        const bodyFormat = { FontSizePt: 12, FontFamily: 'Ubuntu' };
        const revision = Word.getRevision(editor);
        const before = Word.exportDocumentPackage(editor);
        check(`${label}: no header yet`, [before.HeaderParts.length, before.FooterParts.length], [0, 0]);
        check(`${label}: entering the header of page 1`, [Word.enterHeaderFooter(editor, 0, 'header'), story(editor), Word.getText(editor)], [true, ['header', 'default', 0], '']);
        check(`${label}: entering makes no header and no undo step`, [Word.exportDocumentPackage(editor).HeaderParts.length, Word.isModified(editor), Word.canUndo(editor), Word.getRevision(editor) === revision], [0, false, false, true]);
        const format = Word.getSelectionFormat(editor);
        check(`${label}: the size box says the body's size in the empty header (${bodyFormat.FontSizePt} pt)`, [format.FontSizePt, format.FontFamily], [bodyFormat.FontSizePt, bodyFormat.FontFamily]);
        Word.insertPageNumberField(editor);
        Word.insertText(editor, ' of ');
        Word.insertPageCountField(editor);
        check(`${label}: page number and page count: "1 of 10" (it was "1 of 9")`, [Word.getText(editor), Word.getPageCount(editor)], ['1 of 10', 10]);
        const made = Word.exportDocumentPackage(editor);
        const runs = made.HeaderParts.length === 1 ? made.HeaderParts[0].Paragraphs[0].Runs : [];
        check(`${label}: the first edit made the header, with its fields`, [made.HeaderParts.length, runs.filter((run) => run.FieldCode).map((run) => [run.FieldCode.trim(), run.Text])], [1, [['PAGE', '1'], ['NUMPAGES', '10']]]);
        check(`${label}: the fields have the formatting of the text around them`, [new Set(runs.map((run) => run.FontSize)).size, new Set(runs.map((run) => run.FontFamily ?? null)).size, runs.every((run) => !run.Bold)], [1, 1, true]);
        // The size that is painted: capitals in the header at the size the box says, then at twice that.
        Word.insertText(editor, ' HHHH');
        const first = Word.getPageFrames(editor)[0];
        const headerBand = () => inkRows(context, Math.round(first.X ?? 0), Math.round(first.Top) + 4, Math.round(first.Width), Math.round(first.MarginTop) - 8);
        const rowsAtSize = headerBand();
        const sizePt = Word.getSelectionFormat(editor).FontSizePt;
        Word.selectAll(editor);
        Word.setFontSizePt(editor, sizePt * 1.5);
        const rowsAtDouble = headerBand();
        const height = (rows) => rows ? rows[1] - rows[0] + 1 : 0;
        // (A capital is some 0.7 em high: 11.5 px at 12 pt, 10.5 px at 11 pt.)
        check(`${label}: header text is painted at the size the box says (${sizePt} pt: capitals ${height(rowsAtSize)} px high, ${height(rowsAtDouble)} px at one and a half times the size)`,
            [sizePt, height(rowsAtSize) >= Math.floor(sizePt * 96 / 72 * 0.66) && height(rowsAtSize) <= Math.ceil(sizePt * 96 / 72 * 0.80), Math.abs(height(rowsAtDouble) - 1.5 * height(rowsAtSize)) <= 2], [bodyFormat.FontSizePt, true, true]);
        Word.undo(editor); Word.undo(editor);
        check(`${label}: back to "1 of 10", in the header`, [Word.getText(editor), story(editor)[0]], ['1 of 10', 'header']);
        // The last page - the page of the endnotes - has the header, numbered 10.
        check(`${label}: the header of the last page`, [Word.enterHeaderFooter(editor, 9, 'header'), Word.getText(editor), story(editor)], [true, '10 of 10', ['header', 'default', 9]]);
        if (withViewport) {
            Word.setViewport(editor, viewport(Math.max(0, Math.round(frames[9].Top) - 20)));
            Word.repaint(editor);
            const last = Word.getPageFrames(editor)[9];
            const ink = inkOf(context.getImageData(Math.round(last.X ?? 0), 20, Math.round(last.Width), Math.round(last.MarginTop) - 30).data, 140);
            check(`${label}: and it is painted there (${ink} dark pixels in the header area of page 10)`, ink > 20, true);
            Word.setViewport(editor, viewport(0));
        }
        Word.exitHeaderFooter(editor);
        // Saved and read again.
        const reopened = await open(await save(editor), withViewport);
        check(`${label}: saved and read again: still ten pages, "1 of 10" and "7 of 10"`, [Word.getPageCount(reopened), Word.enterHeaderFooter(reopened, 0, 'header'), Word.getText(reopened), Word.enterHeaderFooter(reopened, 6, 'header'), Word.getText(reopened), Word.isModified(reopened)], [10, true, '1 of 10', true, '7 of 10', false]);
        // Undo of everything takes the header away again, in the header.
        Word.enterHeaderFooter(editor, 0, 'header');
        let steps = 0;
        while (Word.canUndo(editor) && steps < 40) { Word.undo(editor); steps++; }
        check(`${label}: undo of the edits takes the header away again; the caret stays in it`, [Word.exportDocumentPackage(editor).HeaderParts.length, story(editor), Word.getText(editor), Word.isModified(editor)], [0, ['header', 'default', 0], '', false]);
        Word.redo(editor);
        check(`${label}: redo makes it again`, [Word.exportDocumentPackage(editor).HeaderParts.length, story(editor)[0]], [1, 'header']);
        while (Word.canRedo(editor)) Word.redo(editor);
        // (All of it again, the capitals and their size too: without those.)
        Word.undo(editor); Word.undo(editor);
        check(`${label}: redone to "1 of 10"`, Word.getText(editor), '1 of 10');
        Word.exitHeaderFooter(editor);

        // ---- 4. a title page and even pages: a blank header until something is typed ----
        const setup = Word.getPageSetup(editor);
        const options = (first, even) => Word.setPageSetupOptions(editor, setup.PageWidth, setup.PageHeight, setup.MarginLeft, setup.MarginRight, setup.MarginTop, setup.MarginBottom, setup.HeaderDistance, setup.FooterDistance, first, even);
        const headerInk = (pageIndex) => {
            if (!withViewport) { const frame = Word.getPageFrames(editor)[pageIndex]; return inkOf(context.getImageData(0, Math.round(frame.Top) + 4, Math.round(frame.Width), Math.round(frame.MarginTop) - 30).data, 140); }
            const frame = Word.getPageFrames(editor)[pageIndex];
            Word.setViewport(editor, viewport(Math.round(frame.Top)));
            Word.repaint(editor);
            const ink = inkOf(context.getImageData(Math.round(frame.X ?? 0), 4, Math.round(frame.Width), Math.round(frame.MarginTop) - 30).data, 140);
            Word.setViewport(editor, viewport(0));
            return ink;
        };
        check(`${label}: every page shows the header`, [headerInk(0) > 20, headerInk(1) > 20, headerInk(2) > 20], [true, true, true]);
        options(true, true);
        check(`${label}: a title page and even pages: page 1 and page 2 show a blank header, page 3 the default one`, [headerInk(0), headerInk(1), headerInk(2) > 20], [0, 0, true]);
        const partsBefore = Word.exportDocumentPackage(editor).HeaderParts.length;
        const revisionBefore = Word.getRevision(editor);
        for (const [pageIndex, variant] of [[0, 'first'], [1, 'even']]) {
            check(`${label}: entering the ${variant} page header changes nothing`, [Word.enterHeaderFooter(editor, pageIndex, 'header'), story(editor), Word.getText(editor), Word.exportDocumentPackage(editor).HeaderParts.length, Word.getRevision(editor) === revisionBefore], [true, ['header', variant, pageIndex], '', partsBefore, true]);
            check(`${label}: the size box in the empty ${variant} header`, Word.getSelectionFormat(editor).FontSizePt, bodyFormat.FontSizePt);
            Word.exitHeaderFooter(editor);
            check(`${label}: left again: no ${variant} header was made, page 3 still shows the default one`, [Word.exportDocumentPackage(editor).HeaderParts.length, Word.getRevision(editor) === revisionBefore, headerInk(pageIndex), headerInk(2) > 20], [partsBefore, true, 0, true]);
            Word.enterHeaderFooter(editor, pageIndex, 'header');
            Word.insertText(editor, variant.toUpperCase());
            const kinds = Word.exportDocumentPackage(editor).HeaderParts.map((part) => part.Kind);
            check(`${label}: typing makes the ${variant} header`, [kinds.filter((kind) => kind === variant).length, kinds.length, Word.getText(editor), headerInk(pageIndex) > 20], [1, partsBefore + 1, variant.toUpperCase(), true]);
            check(`${label}: undo takes it away again, the caret still in the ${variant} header`, [Word.undo(editor), Word.exportDocumentPackage(editor).HeaderParts.length, story(editor), Word.getText(editor), Word.getRevision(editor) === revisionBefore], [true, partsBefore, ['header', variant, pageIndex], '', true]);
            Word.exitHeaderFooter(editor);
        }
        // Footers: a default one with the fields, then a first-page and an even one.
        check(`${label}: the footer of page 3`, [Word.enterHeaderFooter(editor, 2, 'footer'), story(editor), Word.getText(editor)], [true, ['footer', 'default', 2], '']);
        Word.insertText(editor, 'Page ');
        Word.insertPageNumberField(editor);
        Word.insertText(editor, ' / ');
        Word.insertSectionPageCountField(editor);
        check(`${label}: page number and the pages of the section in the footer`, Word.getText(editor), 'Page 3 / 10');
        const footerCaret = Word.getCaretRect(editor);
        const third = Word.getPageFrames(editor)[2];
        check(`${label}: the caret is at the foot of page 3`, [footerCaret.Y > third.Top + third.Height - third.MarginBottom - 2, footerCaret.Y < third.Top + third.Height], [true, true]);
        for (const [pageIndex, variant] of [[0, 'first'], [1, 'even']]) {
            check(`${label}: the ${variant} page footer is a blank one`, [Word.enterHeaderFooter(editor, pageIndex, 'footer'), story(editor), Word.getText(editor)], [true, ['footer', variant, pageIndex], '']);
            Word.insertText(editor, variant + ' foot ');
            Word.insertPageNumberField(editor);
            check(`${label}: typed, with the page number`, Word.getText(editor), `${variant} foot ${pageIndex + 1}`);
        }
        Word.exitHeaderFooter(editor);
        const footers = Word.exportDocumentPackage(editor).FooterParts;
        check(`${label}: three footers: default, first, even`, footers.map((part) => part.Kind).sort(), ['default', 'even', 'first']);
        const again = await open(await save(editor), withViewport);
        const shown = (pageIndex, kind) => { Word.enterHeaderFooter(again, pageIndex, kind); const text = Word.getText(again); const where = story(again); Word.exitHeaderFooter(again); return [text, where[1]]; };
        check(`${label}: saved and read again: each page shows its footer and header`, [shown(0, 'footer'), shown(1, 'footer'), shown(2, 'footer'), shown(4, 'footer'), shown(3, 'footer'), shown(0, 'header'), shown(1, 'header'), shown(2, 'header')],
            [['first foot 1', 'first'], ['even foot 2', 'even'], ['Page 3 / 10', 'default'], ['Page 5 / 10', 'default'], ['even foot 4', 'even'], ['', 'first'], ['', 'even'], ['3 of 10', 'default']]);
        check(`${label}: reading it changed nothing`, [Word.isModified(again), Word.canUndo(again)], [false, false]);
        return editor;
    };
    await fieldsAndSizes(true, 'viewport');
    await fieldsAndSizes(false, 'whole canvas');

    // ---- unlinking copies every block with an id of its own ----
    {
        const editor = await open(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
        Word.enterHeaderFooter(editor, 0, 'header');
        Word.insertTable(editor, 2, 2);
        Word.insertText(editor, 'cell');
        Word.enterHeaderFooter(editor, 1, 'header');
        Word.setHeaderFooterLinkToPrevious(editor, true);
        check('the second section shows the first one\'s header with its table', [Word.getHeaderFooterLinkToPrevious(editor), Word.getText(editor).includes('cell')], [true, true]);
        Word.setHeaderFooterLinkToPrevious(editor, false);
        Word.exitHeaderFooter(editor);
        // Typing in the copy's table leaves the first section's alone.
        Word.enterHeaderFooter(editor, 1, 'header');
        const copyText = Word.getText(editor);
        Word.insertText(editor, 'own');
        Word.exitHeaderFooter(editor);
        Word.enterHeaderFooter(editor, 0, 'header');
        check('typing in the copy leaves the header it was copied from as it was', [Word.getText(editor) === copyText, Word.getText(editor).includes('own')], [true, false]);
    }

    // ---- a long document laid out in slices: the page count in the header is the final one ----
    {
        const long = await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx');
        const author = await open(long, true);
        Word.enterHeaderFooter(author, 0, 'footer');
        Word.insertText(author, 'Page ');
        Word.insertPageNumberField(author);
        Word.insertText(author, ' of ');
        Word.insertPageCountField(author);
        const authored = Word.getPageCount(author);
        check(`a long document (${authored} pages): its footer`, [authored > 40, Word.getText(author).startsWith(`Page 1 of ${authored}`)], [true, true]);
        Word.exitHeaderFooter(author);
        const bytes = await save(author);
        // (Read again the document has other pages - it is laid out without the page breaks its file recorded: the count is that of the file as saved.)
        const total = Word.getPageCount(await open(bytes, true));
        const progressive = await open(bytes, true, canvas);
        const context = canvas.getContext('2d');
        const frame = Word.getPageFrames(progressive)[0];
        // (The footer area of page 1 in a viewport scrolled to the page's last 800 px.)
        const footerTop = Math.round(frame.Height - frame.MarginBottom) - (Math.round(frame.Height) - 800) + 2;
        const footerPixels = () => { Word.repaint(progressive); return Array.from(context.getImageData(0, footerTop, Math.round(frame.Width), 800 - footerTop).data).join(','); };
        const early = { complete: Word.isLayoutComplete(progressive), pages: Word.getPageFrames(progressive).length, pixels: null };
        // (The footer of page 1 is below an 800 px viewport: scroll to it.)
        Word.setViewport(progressive, viewport(Math.round(frame.Height) - 800));
        early.pixels = footerPixels();
        early.complete = Word.isLayoutComplete(progressive);
        early.pages = Word.getPageFrames(progressive).length;
        await Word.whenLayoutComplete(progressive);
        await new Promise((resolve) => requestAnimationFrame(() => resolve()));
        const late = { pages: Word.getPageFrames(progressive).length, pixels: footerPixels() };
        const whole = await open(bytes, true, canvas);
        Word.completeLayout(whole);
        Word.setViewport(whole, viewport(Math.round(frame.Height) - 800));
        const reference = (() => { Word.repaint(whole); return Array.from(context.getImageData(0, footerTop, Math.round(frame.Width), 800 - footerTop).data).join(','); })();
        check(`laid out in slices (${early.pages} pages at first, complete: ${early.complete}), then whole: ${late.pages} pages`, [early.complete, early.pages < total, late.pages], [false, true, total]);
        check('the footer of page 1 is painted with the final page count once the layout is complete (the pixels of a layout made at once)', [late.pixels === reference, early.pixels !== late.pixels], [true, true]);
        check('and entered, it says so', [Word.enterHeaderFooter(progressive, 0, 'footer'), Word.getText(progressive).startsWith(`Page 1 of ${total}`)], [true, true]);
    }
} catch (e) { fail(e); }
finish();
