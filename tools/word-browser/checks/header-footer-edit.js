// Headers and footers edited in place: enterHeaderFooter / a double click on
// the page's header area puts the caret into that page's header (the body is
// dimmed, a labelled line marks the header), typing, undo / redo, tables and
// fields work there, Escape / a double click on the body / exitHeaderFooter go
// back; getCaretContext says where the caret is. What is typed is part of the
// document (saved, read again), undo steps made in a header and in the body
// interleave, a header is made where a page has none, each section has its
// own header ("link to previous" on and off), and the body moves down when
// the header grows. Real mouse and keyboard events drive the first part.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const viewport = (top) => ({ scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
    const sections = await OF.processors.forEditorPackage(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const editor = Word.initializeCanvas(canvas);
    Word.setViewport(editor, viewport(0));
    Word.importFromDocxPackage(editor, sections);
    const context = canvas.getContext('2d');
    const region = (x, y, w, h) => context.getImageData(x, y, w, h).data;
    const inkOf = (data, limit = 100) => { let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < limit && data[i + 1] < limit && data[i + 2] < limit) ink++; return ink; };
    const blueOf = (data) => { let blue = 0; for (let i = 0; i < data.length; i += 4) if (data[i + 2] > 150 && data[i] < 140 && data[i + 2] - data[i] > 60) blue++; return blue; };
    const headerTexts = (pkg) => pkg.HeaderParts.map((part) => (part.Paragraphs ?? []).map((para) => (para.Runs ?? []).map((run) => run.Text ?? '').join('')).join('|'));
    const footerTexts = (pkg) => pkg.FooterParts.map((part) => (part.Paragraphs ?? []).map((para) => (para.Runs ?? []).map((run) => run.Text ?? '').join('')).join('|'));
    const exported = () => Word.exportDocumentPackage(editor);
    const story = () => { const c = Word.getCaretContext(editor); return [c.InHeaderFooter, c.Story]; };
    const frames = Word.getPageFrames(editor);
    const first = frames[0];
    const state = {};

    // ---- before: the body ----
    check('in the body', [story(), Word.getCaretContext(editor).HeaderFooter], [[false, 'body'], null]);
    check('the headers of the three sections', headerTexts(exported()), ['header one', 'header two', 'header three']);
    const bodyBefore = inkOf(region(0, first.MarginTop + 2, 1000, 300));
    const changes = [];
    Word.onChange(editor, (change) => changes.push(change));
    const revision = Word.getRevision(editor);

    // ---- a real double click on the header area of the first page ----
    window.hf = {
        afterEnter() {
            state.entered = story();
            state.info = Word.getCaretContext(editor).HeaderFooter;
            state.modifiedAtEnter = [Word.isModified(editor), Word.canUndo(editor), Word.getRevision(editor) === revision];
            state.changeAtEnter = changes.length ? [changes[changes.length - 1].DocumentChanged, changes[changes.length - 1].SelectionChanged] : null;
            state.bodyDimmed = inkOf(region(0, first.MarginTop + 2, 1000, 300));
            state.boundary = blueOf(region(0, first.MarginTop - 6, 1000, 40));
            state.textAtEnter = Word.getText(editor);
            state.caret = Word.getCaretRect(editor);
        },
        afterTyping() {
            state.typed = Word.getText(editor);
            state.typedHeaders = headerTexts(exported());
            state.afterTyping = [Word.isModified(editor), Word.canUndo(editor), story()];
            state.headerInk = inkOf(region(0, 0, 1000, first.MarginTop - 8));
        },
        afterUndo() { state.undone = [Word.getText(editor), headerTexts(exported())[0], story(), Word.getRevision(editor) === revision, Word.isModified(editor)]; },
        afterRedo() { state.redone = [Word.getText(editor), story()]; },
        afterEscape() { state.escaped = [story(), Word.getText(editor).includes('section one'), headerTexts(exported())[0]]; },
        afterBodyTyping() { state.body = [Word.getText(editor).includes('Bsection one') || Word.getText(editor).includes('B'), story()]; state.bodyText = Word.getText(editor); },
        afterUndoBody() { state.undoBody = [Word.getText(editor) === state.bodyTextBefore, story(), headerTexts(exported())[0]]; },
        afterUndoHeader() { state.undoHeader = [story(), headerTexts(exported())[0], Word.getText(editor)]; },
        afterRedoHeader() { state.redoHeader = [story(), headerTexts(exported())[0]]; },
        afterRedoBody() { state.redoBody = [story(), Word.getText(editor) === state.bodyText]; },
        beforeBody() { state.bodyTextBefore = Word.getText(editor); },
        afterReenter() { state.reentered = story(); },
        afterBodyDoubleClick() { state.left = story(); },
    };
    const headerY = Math.round(first.MarginTop / 2);
    window.drive = [
        ['dblclick', 300, headerY], ['eval', 'window.hf.afterEnter()'],
        ['type', 'Typed '], ['eval', 'window.hf.afterTyping()'],
        ['press', 'Control+z'], ['press', 'Control+z'], ['press', 'Control+z'], ['press', 'Control+z'], ['press', 'Control+z'], ['press', 'Control+z'], ['eval', 'window.hf.afterUndo()'],
        ['press', 'Control+y'], ['press', 'Control+y'], ['press', 'Control+y'], ['press', 'Control+y'], ['press', 'Control+y'], ['press', 'Control+y'], ['eval', 'window.hf.afterRedo()'],
        ['press', 'Escape'], ['eval', 'window.hf.afterEscape()'],
        ['click', 300, first.MarginTop + 12], ['eval', 'window.hf.beforeBody()'], ['type', 'B'], ['eval', 'window.hf.afterBodyTyping()'],
        ['press', 'Control+z'], ['eval', 'window.hf.afterUndoBody()'],
        ['press', 'Control+z'], ['eval', 'window.hf.afterUndoHeader()'],
        ['press', 'Control+y'], ['eval', 'window.hf.afterRedoHeader()'],
        ['press', 'Control+y'], ['eval', 'window.hf.afterRedoBody()'],
        ['dblclick', 300, headerY], ['eval', 'window.hf.afterReenter()'],
        ['dblclick', 300, first.MarginTop + 200], ['eval', 'window.hf.afterBodyDoubleClick()'],
    ];
    window.finish = async () => {
        try {
            check('a double click on the header area: the caret is in the header', state.entered, [true, 'header']);
            check('getCaretContext says which', state.info, { Kind: 'header', Variant: 'default', PageIndex: 0, Section: 0, LinkedToPrevious: false, CanLinkToPrevious: false });
            check('entering is no edit: not modified, nothing to undo, the same revision; a selection change', [state.modifiedAtEnter, state.changeAtEnter], [[false, false, true], [false, true]]);
            check('the document the commands act on is the header', state.textAtEnter, 'header one');
            check('the caret is in the header area', [state.caret !== null && state.caret.Y < first.MarginTop, state.caret !== null && state.caret.Y > 0], [true, true]);
            check(`the body is dimmed (dark pixels ${bodyBefore} -> ${state.bodyDimmed})`, [bodyBefore > 50, state.bodyDimmed < bodyBefore / 2], [true, true]);
            check('a labelled line marks the header', state.boundary > 150, true);
            check('typed with the keyboard into the header', [state.typed.includes('Typed '), state.typed.replace('Typed ', '')], [true, 'header one']);
            check('it is the document\'s header (the first section\'s; the others are as they were)', [state.typedHeaders[0].includes('Typed '), state.typedHeaders.slice(1)], [true, ['header two', 'header three']]);
            check('an edit: modified, undoable, still in the header', state.afterTyping, [true, true, [true, 'header']]);
            check('Ctrl+Z takes the typing back, in the header; the saved revision is back', state.undone, ['header one', 'header one', [true, 'header'], true, false]);
            check('Ctrl+Y brings it back', [state.redone[0].includes('Typed '), state.redone[1]], [true, [true, 'header']]);
            check('Escape goes back to the body; the header keeps what was typed', [state.escaped[0], state.escaped[1], state.escaped[2].includes('Typed ')], [[false, 'body'], true, true]);
            check('typing in the body', state.body, [true, [false, 'body']]);
            check('undo: the body edit first (still the body, the header as typed)', [state.undoBody[0], state.undoBody[1], state.undoBody[2].includes('Typed ')], [true, [false, 'body'], true]);
            check('undo again: the last header edit (a keystroke), and the caret is back in the header', [state.undoHeader[0], state.undoHeader[1], state.undoHeader[2]], [[true, 'header'], state.typedHeaders[0].replace('Typed ', 'Typed'), state.typed.replace('Typed ', 'Typed')]);
            check('redo: the header edit (in the header), then the body edit (in the body)', [state.redoHeader[0], state.redoHeader[1], state.redoBody], [[true, 'header'], state.typedHeaders[0], [[false, 'body'], true]]);
            check('a double click on the header again, then on the body', [state.reentered, state.left], [[true, 'header'], [false, 'body']]);

            // ---- through the API: a table and a field in a footer, a footer where there was none ----
            check('enterHeaderFooter(page 0, "footer") where the page has no footer: the caret is in an empty one', [Word.enterHeaderFooter(editor, 0, 'footer'), story(), Word.getCaretContext(editor).HeaderFooter.Kind, Word.getText(editor)], [true, [true, 'footer'], 'footer', '']);
            Word.insertText(editor, 'Page ');
            Word.insertPageNumberField(editor);
            Word.insertText(editor, ' of ');
            Word.insertPageCountField(editor);
            const pageCount = Word.getPageCount(editor);
            check('text and page number fields in the footer', Word.getText(editor), `Page 1 of ${pageCount}`);
            check('the caret is at the foot of the page', (() => { const rect = Word.getCaretRect(editor); return rect.Y > first.Height - first.MarginBottom - 2 && rect.Y < first.Height; })(), true);
            check('exitHeaderFooter', [Word.exitHeaderFooter(editor), story(), Word.exitHeaderFooter(editor)], [true, [false, 'body'], false]);
            const withFooter = exported();
            const made = withFooter.FooterParts.find((part) => (part.Paragraphs ?? []).some((para) => (para.Runs ?? []).some((run) => (run.Text ?? '').startsWith('Page'))));
            check('the new footer is a part of the document, referred to by the first section', [made !== undefined, made !== undefined && typeof made.RelationshipId === 'string', footerTexts(withFooter).includes('footer of section two')], [true, true, true]);
            check('with its fields', made && made.Paragraphs[0].Runs.filter((run) => run.FieldCode).map((run) => run.FieldCode.trim()), ['PAGE', 'NUMPAGES']);
            // A table in the second section's header (page 1).
            check('the header of a page of the second section', [Word.enterHeaderFooter(editor, 1, 'header'), Word.getText(editor), Word.getCaretContext(editor).HeaderFooter.Section], [true, 'header two', 1]);
            Word.insertTable(editor, 2, 2);
            Word.insertText(editor, 'cell');
            check('a table in the header: the caret is in its first cell', (() => { const c = Word.getCaretContext(editor); return [c.InTable, c.Row, c.Column, c.InHeaderFooter]; })(), [true, 0, 0, true]);
            const bodyTopBefore = Word.getPageFrames(editor)[1].MarginTop;
            // ---- link to previous ----
            check('the second section has its own header', [Word.getHeaderFooterLinkToPrevious(editor), Word.getCaretContext(editor).HeaderFooter.CanLinkToPrevious], [false, true]);
            check('link to previous: it shows the first section\'s header', [Word.setHeaderFooterLinkToPrevious(editor, true), Word.getHeaderFooterLinkToPrevious(editor), story(), Word.getText(editor).includes('Typed ')], [true, true, [true, 'header'], true]);
            check('undo takes the link away again, and the caret stays in the header', [Word.undo(editor), Word.getHeaderFooterLinkToPrevious(editor), story(), Word.getText(editor).includes('header two')], [true, false, [true, 'header'], true]);
            check('redo links it again, in the header', [Word.redo(editor), Word.getHeaderFooterLinkToPrevious(editor), story()], [true, true, [true, 'header']]);
            check('and undo once more', [Word.undo(editor), Word.getHeaderFooterLinkToPrevious(editor), story()], [true, false, [true, 'header']]);
            Word.exitHeaderFooter(editor);
            check('entered again, the section has its own header as before', [Word.enterHeaderFooter(editor, 1, 'header'), Word.getHeaderFooterLinkToPrevious(editor), Word.getText(editor).includes('header two')], [true, false, true]);
            Word.setHeaderFooterLinkToPrevious(editor, true);
            check('unlinking gives the section a copy of what it showed, to change on its own', [Word.setHeaderFooterLinkToPrevious(editor, false), Word.getHeaderFooterLinkToPrevious(editor), Word.getText(editor).includes('Typed ')], [true, false, true]);
            Word.insertText(editor, 'own ');
            Word.exitHeaderFooter(editor);
            const unlinked = headerTexts(exported());
            check('typing there leaves the first section\'s header alone', [unlinked[0].includes('own '), unlinked.some((text) => text.includes('own ') && text.includes('Typed '))], [false, true]);

            // ---- the body moves down when the header grows ----
            Word.enterHeaderFooter(editor, 0, 'header');
            const topBefore = Word.getPageFrames(editor)[0].MarginTop;
            const input = Word.getInputElement(editor);
            for (let i = 0; i < 4; i++) input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
            Word.insertText(editor, 'a fifth header line');
            const topAfter = Word.getPageFrames(editor)[0].MarginTop;
            check(`the body starts further down when the header has more lines (${topBefore} -> ${topAfter})`, topAfter > topBefore + 20, true);
            Word.exitHeaderFooter(editor);

            // ---- saved and read again ----
            const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(exported())).arrayBuffer());
            const reopened = await OF.processors.forEditorPackage(bytes);
            const before = exported();
            check('saved and read again: the headers and footers are what was typed', [headerTexts(reopened).sort(), footerTexts(reopened).sort()], [headerTexts(before).sort(), footerTexts(before).sort()]);
            check('the table in the header is in the file', reopened.HeaderParts.some((part) => (part.Blocks ?? part.Paragraphs ?? []).some((block) => block.Type === 'Table' || block.Rows !== undefined)), true);
            const again = Word.initializeCanvas(document.createElement('canvas'));
            Word.importFromDocxPackage(again, reopened);
            check('and the document opens with them', [Word.getPageCount(again) >= 3, Word.enterHeaderFooter(again, 0, 'footer'), Word.getText(again).startsWith('Page 1 of')], [true, true, true]);
            note(`headers ${JSON.stringify(headerTexts(before))} footers ${JSON.stringify(footerTexts(before))}`);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
