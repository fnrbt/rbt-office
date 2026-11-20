// How sections start, and where their text stands on a page, read from a file,
// laid out, changed and saved:
//  - a section that starts on an odd / even page (w:type oddPage / evenPage in
//    its own properties) has a blank page before it where the next page is of
//    the other kind; the blank page is numbered and has the section's header;
//  - w:vAlign (center / bottom / both) moves the text of a page that is not full;
//  - a page break in a section with columns starts a new page, a column break
//    (w:br w:type="column") goes on in the next column - and is saved as one;
//  - setSectionBreak(kind) says how the section that STARTS at the break
//    starts: the w:type is written into the properties of the section after
//    the caret's paragraph, and the paragraph's own w:sectPr has the
//    properties of the section it was cut from.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const p = (text, pPr = '', extra = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${run(text)}${extra}</w:p>`;
    const page = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
    const bytesFor = async (body) => { zip.file('word/document.xml', head + body + '</w:body></w:document>'); return zip.generateAsync({ type: 'uint8array' }); };
    const canvas = document.getElementById('c');
    const open = async (bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes)); return editor; };
    const savedXml = async (editor) => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        return { xml: await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string'), bytes: out };
    };
    const lineOf = (editor, text) => Word.getLayoutLines(editor).find((line) => (line.Text ?? '').startsWith(text));
    const pageOf = (editor, text) => { const frames = Word.getPageFrames(editor); const y = lineOf(editor, text).Y; let index = 0; frames.forEach((frame, i) => { if ((frame.Top ?? frame.Y) <= y) index = i; }); return index; };
    const context = canvas.getContext('2d');
    const inkIn = (frame) => {
        // (Between the margins: the page's header and footer are not its text.)
        const top = Math.round(frame.Top ?? frame.Y), data = context.getImageData(Math.round(frame.X ?? 0) + 4, top + 96, Math.round(frame.Width) - 8, Math.round(frame.Height) - 192).data;
        let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 110) ink++; return ink;
    };

    // ---- odd and even page sections -------------------------------------
    // Section 1 (one page); section 2 starts on an odd page; section 3 on an even page.
    {
        const body = p('section one') + `<w:p><w:pPr><w:sectPr>${page}</w:sectPr></w:pPr></w:p>`
            + p('section two') + `<w:p><w:pPr><w:sectPr><w:type w:val="oddPage"/>${page}</w:sectPr></w:pPr></w:p>`
            + p('section three') + `<w:sectPr><w:type w:val="evenPage"/>${page}</w:sectPr>`;
        const editor = await open(await bytesFor(body));
        const frames = Word.getPageFrames(editor);
        check('an odd-page section after page 1: four pages, a blank second one', [frames.length, pageOf(editor, 'section one'), pageOf(editor, 'section two'), pageOf(editor, 'section three')], [4, 0, 2, 3]);
        Word.repaint(editor);
        const ink = frames.map(inkIn);
        check('the blank page has no text (it has the header and footer its section gives it), the others have theirs', [ink[0] > 20, ink[1], ink[2] > 20, ink[3] > 20], [true, 0, true, true]);
        const { xml } = await savedXml(editor);
        check('saved as read', [(xml.match(/<w:type w:val="oddPage"\/>/g) ?? []).length, (xml.match(/<w:type w:val="evenPage"\/>/g) ?? []).length], [1, 1]);
        // Typing in the first section until it has two pages: the odd-page section needs no blank page then.
        const one = lineOf(editor, 'section one');
        Word.setSelection(editor, one.BlockId, 0, one.BlockId, 0);
        Word.insertPageBreak(editor);
        check('the first section two pages long: the odd-page section follows on page 3, no blank page', [Word.getPageFrames(editor).length, pageOf(editor, 'section one'), pageOf(editor, 'section two'), pageOf(editor, 'section three')], [4, 1, 2, 3]);
        Word.undo(editor);
        check('and undone: the blank page is back', [Word.getPageFrames(editor).length, pageOf(editor, 'section two')], [4, 2]);
        Word.dispose(editor);
    }

    // ---- vAlign ------------------------------------------------------------
    {
        const titleBody = (vAlign) => p('Title') + `<w:p><w:pPr><w:sectPr>${page}${vAlign ? `<w:vAlign w:val="${vAlign}"/>` : ''}</w:sectPr></w:pPr></w:p>` + p('Body') + `<w:sectPr>${page}</w:sectPr>`;
        const yOf = async (vAlign) => { const editor = await open(await bytesFor(titleBody(vAlign))); const result = { title: lineOf(editor, 'Title').Y, body: lineOf(editor, 'Body').Y, saved: (await savedXml(editor)).xml, editor }; return result; };
        const top = await yOf(null), center = await yOf('center'), bottom = await yOf('bottom');
        note(`a title page: the title at y ${top.title.toFixed(1)} (top), ${center.title.toFixed(1)} (center), ${bottom.title.toFixed(1)} (bottom)`);
        check('center: the title is near the middle of the page, bottom: near the bottom margin', [top.title < 130, center.title > 420 && center.title < 560, bottom.title > 860 && bottom.title < 960], [true, true, true]);
        check('the section after it is where it was', [center.body, bottom.body], [top.body, top.body]);
        // Written from the model (the paragraph that ends the section is changed: its alignment), the vAlign is still there.
        const end = Word.getLayoutLines(center.editor).filter((line) => line.Y > center.title && line.Y < center.body)[0];
        check('saved as read: w:vAlign in the section\'s properties', [/<w:vAlign w:val="center"\/>/.test(center.saved), /<w:vAlign/.test(top.saved)], [true, false]);
        if (end) {
            Word.setSelection(center.editor, end.BlockId, 0, end.BlockId, 0);
            Word.alignRight(center.editor);
            const again = (await savedXml(center.editor)).xml;
            check('and with the paragraph that ends the section changed', /<w:vAlign w:val="center"\/>/.test(again), true);
        }
        // The body's own.
        const single = await open(await bytesFor(p('Alone') + `<w:sectPr>${page}<w:vAlign w:val="center"/></w:sectPr>`));
        check('a document of one section, centred', [lineOf(single, 'Alone').Y > 420, /<w:vAlign w:val="center"\/>/.test((await savedXml(single)).xml)], [true, true]);
        Word.insertText(single, 'x');
        check('and after an edit', /<w:vAlign w:val="center"\/>/.test((await savedXml(single)).xml), true);
        // Set through the API: of the section the caret is in.
        {
            const editor = await open(await bytesFor(titleBody(null)));
            const title = Word.findText(editor, 'Title', true)[0].BlockId, body = Word.findText(editor, 'Body', true)[0].BlockId;
            const before = [lineOf(editor, 'Title').Y, lineOf(editor, 'Body').Y];
            Word.setSelection(editor, title, 0, title, 0);
            check('setSectionVerticalAlignment("center") with the caret in the first section: its text is in the middle of its page, the next section where it was',
                [Word.getSectionVerticalAlignment(editor), Word.setSectionVerticalAlignment(editor, 'center'), Word.getSectionVerticalAlignment(editor), Math.abs(lineOf(editor, 'Title').Y - center.title) < 1, lineOf(editor, 'Body').Y === before[1]],
                ['top', true, 'center', true, true]);
            check('set again it changes nothing', Word.setSectionVerticalAlignment(editor, 'center'), false);
            const xml = (await savedXml(editor)).xml;
            check('saved in that section\'s properties, not in the last section\'s', [/<w:pPr><w:sectPr>(?:(?!<\/w:sectPr>).)*<w:vAlign w:val="center"\/>/.test(xml), (xml.match(/<w:vAlign /g) ?? []).length], [true, 1]);
            check('one undo: from the top again', [Word.undo(editor), Word.getSectionVerticalAlignment(editor), lineOf(editor, 'Title').Y], [true, 'top', before[0]]);
            // The last section (no paragraph ends it).
            Word.setSelection(editor, body, 0, body, 0);
            check('with the caret in the last section: "bottom" is that section\'s',
                [Word.setSectionVerticalAlignment(editor, 'bottom'), Word.getSectionVerticalAlignment(editor), lineOf(editor, 'Body').Y > 860, lineOf(editor, 'Title').Y === before[0]], [true, 'bottom', true, true]);
            const last = (await savedXml(editor)).xml;
            check('saved in the body\'s section properties', [/<w:vAlign w:val="bottom"\/>(?:(?!<w:p[ >]).)*<\/w:sectPr><\/w:body>/.test(last), (last.match(/<w:vAlign /g) ?? []).length], [true, 1]);
            check('"top" takes it away', [Word.setSectionVerticalAlignment(editor, 'top'), /<w:vAlign/.test((await savedXml(editor)).xml)], [true, false]);
            Word.dispose(editor);
        }
    }

    // ---- page break and column break in columns -----------------------------
    {
        const cols = '<w:cols w:num="2" w:space="720"/>';
        const withBreak = (type) => p('first column', '', `<w:r><w:br w:type="${type}"/></w:r>`) + p('after the break') + `<w:sectPr>${page}${cols}</w:sectPr>`;
        const pageBreak = await open(await bytesFor(withBreak('page')));
        check('a page break in two columns: the text after it is on the next page', [pageOf(pageBreak, 'first column'), pageOf(pageBreak, 'after the break'), Word.getPageFrames(pageBreak).length], [0, 1, 2]);
        const columnBreak = await open(await bytesFor(withBreak('column')));
        const first = lineOf(columnBreak, 'first column'), after = lineOf(columnBreak, 'after the break');
        // (Below the paragraph mark the break's paragraph leaves at the column's top - an empty line in the
        // mark's font, and its space after: Word 2010 and 2019, fixtures float-wrap cbr-text, cbr-own-a12.)
        check('a column break: in the next column of the same page, a line below its top', [pageOf(columnBreak, 'after the break'), after.Y - first.Y > first.Height / 2 && after.Y - first.Y <= first.BlockHeight + 0.5, after.X > first.X + 200, Word.getPageFrames(columnBreak).length], [0, true, true, 1]);
        // Saved: unchanged it is the XML that was read; changed (made bold: written from the model or patched) it is still a column break.
        Word.setSelection(columnBreak, first.BlockId, 0, first.BlockId, 5);
        Word.toggleBold(columnBreak);
        const { xml, bytes } = await savedXml(columnBreak);
        check('a column break is saved as one', [(xml.match(/<w:br w:type="column"\/>/g) ?? []).length, /<w:br w:type="page"\/>/.test(xml)], [1, false]);
        const reopened = await open(bytes);
        check('and read again', [pageOf(reopened, 'after the break'), Word.getPageFrames(reopened).length], [0, 1]);
        // The paragraph written from the model as a whole (a list made of it).
        Word.setSelection(reopened, lineOf(reopened, 'first column').BlockId, 0, lineOf(reopened, 'first column').BlockId, 0);
        Word.toggleBulletList(reopened);
        const listed = (await savedXml(reopened)).xml;
        check('and written from the model', [(listed.match(/<w:br w:type="column"\/>/g) ?? []).length, /<w:br w:type="page"\/>/.test(listed)], [1, false]);
    }

    // ---- setSectionBreak -----------------------------------------------------
    {
        const editor = await open(await bytesFor(p('one') + p('two') + p('three') + `<w:sectPr>${page}<w:pgNumType w:start="5"/></w:sectPr>`));
        const caretIn = (text) => { const line = lineOf(editor, text); Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0); };
        const sections = (xml) => (xml.match(/<w:sectPr[ >][^]*?<\/w:sectPr>/g) ?? []).map((sectPr) => ({ type: (sectPr.match(/<w:type w:val="(\w+)"/) ?? [])[1] ?? null, page: /<w:pgSz w:w="12240" w:h="15840"\/>/.test(sectPr), start: (sectPr.match(/<w:pgNumType [^>]*w:start="(\d+)"/) ?? [])[1] ?? null }));
        caretIn('one');
        Word.setSectionBreak(editor, 'continuous');
        let saved = await savedXml(editor);
        check('a continuous break after "one": the section after it is continuous (its type in the body\'s sectPr), the paragraph\'s own sectPr says none; both have the page',
            sections(saved.xml), [{ type: 'nextPage', page: true, start: '5' }, { type: 'continuous', page: true, start: null }]);
        check('the paragraph that ends the first section is "one"', /<w:p[ >](?:(?!<\/w:p>)[^])*<w:sectPr>(?:(?!<\/w:p>)[^])*one<\/w:t>/.test(saved.xml), true);
        check('laid out: one page', [Word.getPageFrames(editor).length, pageOf(editor, 'two')], [1, 0]);
        // Read again: the same sections.
        const reopened = await open(saved.bytes);
        check('read again: one page, and saved the same', [Word.getPageFrames(reopened).length, JSON.stringify(sections((await savedXml(reopened)).xml))], [1, JSON.stringify(sections(saved.xml))]);
        // A next-page break further down: what starts after "two" starts on a new page; "two" stays continuous.
        caretIn('two');
        Word.setSectionBreak(editor, 'oddPage');
        saved = await savedXml(editor);
        check('an odd-page break after "two"', sections(saved.xml).map((section) => section.type), ['nextPage', 'continuous', 'oddPage']);
        check('laid out: "one" and "two" on page 1, a blank page, "three" on page 3', [Word.getPageFrames(editor).length, pageOf(editor, 'two'), pageOf(editor, 'three')], [3, 0, 2]);
        // The type of the break after "one" changed.
        caretIn('one');
        Word.setSectionBreak(editor, 'nextPage');
        saved = await savedXml(editor);
        check('the break after "one" made a next-page break', [sections(saved.xml).map((section) => section.type), pageOf(editor, 'two')], [['nextPage', 'nextPage', 'oddPage'], 1]);
        // Taken away.
        Word.setSectionBreak(editor, 'none');
        saved = await savedXml(editor);
        check('the break after "one" taken away: two sections, the first starts as it did', [sections(saved.xml).map((section) => section.type), pageOf(editor, 'two')], [['nextPage', 'oddPage'], 0]);
        Word.undo(editor); Word.undo(editor); Word.undo(editor); Word.undo(editor);
        saved = await savedXml(editor);
        check('all undone: one section, as read', sections(saved.xml), [{ type: null, page: true, start: '5' }]);
    }
} catch (e) { fail(e); }
finish();
