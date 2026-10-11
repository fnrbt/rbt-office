// Sections on the page: a section's text stands at its own left margin (the
// layouts of a section whose left margin is not the document's are moved by
// the difference, so painting, the caret and the pointer agree), and the
// layout queries find a page by the pages as they are - a landscape section
// among portrait ones has shorter pages, so a page is not y divided by one
// page's height.
// The document: synthetic/three-sections with the second (landscape) section
// given a left margin of 2 inches.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const xml = await zip.file('word/document.xml').async('string');
    const landscape = '<w:pgSz w:w="15840" w:h="12240"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"';
    if (!xml.includes(landscape)) throw new Error('the fixture changed');
    zip.file('word/document.xml', xml.replace(landscape, landscape.replace('w:left="1440"', 'w:left="2880"')));
    const editorPackage = await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' }));
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, editorPackage);
    const frames = Word.getPageFrames(editor);
    // (A section starts on the page after the one the section before it ends on: its sectPr has no
    // w:type, which is nextPage. The empty paragraph that ends a section is the section's last line,
    // not a page - there used to be an empty page after each.)
    check('a portrait page, a landscape page, a portrait page', frames.map((frame) => [frame.Width, frame.Height]), [[816, 1056], [1056, 816], [816, 1056]]);
    check('the landscape section has its own left margin', frames.map((frame) => frame.MarginLeft), [96, 192, 96]);
    const pageOf = (y) => { let index = 0; frames.forEach((frame, i) => { if (frame.Top <= y) index = i; }); return index; };
    const lines = Word.getLayoutLines(editor);
    const lineOf = (text) => lines.find((line) => (line.Text ?? '').startsWith(text));
    const one = lineOf('section one'), two = lineOf('section two'), three = lineOf('section three');
    note(`lines: one x ${one.X} page ${one.Page}; two x ${two.X} y ${two.Y} page ${two.Page} pageY ${two.PageY}; three y ${three.Y} page ${three.Page} pageY ${three.PageY}`);
    // (c) the left margin
    check('a line of the first section starts at the document\'s left margin, one of the landscape section at its own', [Math.round(one.X), Math.round(two.X)], [96, 192]);
    Word.setSelection(editor, two.BlockId, 0, two.BlockId, 0);
    const caret = Word.getCaretRect(editor);
    check('the caret at the start of that line is at the section\'s margin', Math.round(caret.X), 192);
    const context = canvas.getContext('2d');
    const inkBetween = (x0, x1, y0, y1) => { const data = context.getImageData(x0, y0, x1 - x0, y1 - y0).data; let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 110) ink++; return ink; };
    const lineTop = Math.floor(two.Y) - 2, lineBottom = Math.ceil(two.Y) + 26;
    check('painted there: nothing left of the margin, text right of it', [inkBetween(60, 186, lineTop, lineBottom), inkBetween(190, 400, lineTop, lineBottom) > 20], [0, true]);
    const hit = Word.hitTestPosition(editor, 192 + 30, two.Y + 8);
    check('a point in that line is in its paragraph, a few characters in', [hit && hit.BlockId === two.BlockId, hit && hit.Offset > 0 && hit.Offset < 8], [true, true]);
    // (f) pages by their frames
    // (After a landscape page the last page does not start at a multiple of a portrait page's stride.)
    check('each line is on the page its y is in', [one.Page, two.Page, three.Page], [pageOf(one.Y) + 1, pageOf(two.Y) + 1, pageOf(three.Y) + 1]);
    check('the last section is on the third page, which starts after a shorter one', [three.Page, frames[2].Top < 2 * frames[1].Top, Math.floor(three.Y / frames[1].Top) + 1], [3, true, 2]);
    const onPage = (line) => Math.round((line.Y - frames[pageOf(line.Y)].Top) - line.PageY);
    check('a line\'s place on its page is measured from that page\'s top (the same offset on every page)', [onPage(two), onPage(three), three.PageY > 0 && three.PageY < 300], [onPage(one), onPage(one), true]);
    const information = Word.getRangeInformation(editor, three.BlockId, 0, 1);
    check('getRangeInformation: the page of a range after the landscape pages', [information.wdActiveEndPageNumber, information.wdNumberOfPagesInDocument], [3, 3]);
    const bounds = Word.getRangeBounds(editor, three.BlockId, 0, 1);
    check('getRangeBounds: its page and its place on it', [bounds.Page, Math.round(bounds.PageY - (bounds.Y - frames[2].Top))], [3, 0]);
    // A continuous section break: the section before it keeps its own columns and header.
    // phpword-sample-11: paragraphs 1-6 are the first section (its sectPr, in the sixth: one column,
    // w:titlePg, a first-page header "First page header"), the seventh alone the second
    // (w:type continuous, w:cols w:num="2" w:space="720"), the rest the last (continuous, one column).
    // The page is 11906 twips wide with 1418 at each side: 9070 twips = 604.67 px of text; a column
    // is (9070 - 720) / 2 = 4175 twips = 278.33 px.
    {
        const second = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(second, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/phpword-sample-11-readword2007.docx')));
        const all = Word.getLayoutLines(second);
        const first = (text) => all.find((line) => (line.Text ?? '').startsWith(text));
        const library = first('PHPWord is a pure PHP library'), formats = first('PHPWord can apply font formats'), justified = first('PHPWord can also format paragraph'), multi = first('PHPWord can also create'), lists = first('PHPWord can create bulleted lists');
        check('the first section is one column wide', [library, formats, justified].map((line) => Math.round(line.AvailableWidth)), [605, 605, 605]);
        check('the continuous section after it has two columns', Math.round(multi.AvailableWidth), 278);
        check('the section after that one again', Math.round(lists.AvailableWidth), 605);
        check('all on the first page, one below the other', [library.Page, multi.Page, lists.Page, multi.Y > justified.Y, lists.Y > multi.Y], [1, 1, 1, true, true]);
        check('a line of the first section runs past the middle of the page (it was broken at the first column\'s end)', library.Width > 320, true);
        check('the first page shows the first section\'s first-page header', [Word.enterHeaderFooter(second, 0, 'header'), Word.getText(second)], [true, 'First page header']);
    }
    // Sections that start on a new page: the page after the one the section before ends on.
    {
        const third = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(third, await OF.processors.forEditorPackage(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx'))));
        const text = Word.getLayoutLines(third).filter((line) => (line.Text ?? '').includes('section')).map((line) => [line.Text, line.Page]);
        check('three sections, three pages', [Word.getPageFrames(third).length, text], [3, [['section one', 1], ['section two', 2], ['still section two', 2], ['section three', 3]]]);
    }
} catch (e) { fail(e); }
finish();
