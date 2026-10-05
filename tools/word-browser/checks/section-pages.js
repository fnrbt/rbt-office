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
    // (Each of the first two sections ends with an empty paragraph on a page of its own.)
    check('portrait pages, landscape pages, a portrait page', frames.map((frame) => [frame.Width, frame.Height]), [[816, 1056], [816, 1056], [1056, 816], [1056, 816], [816, 1056]]);
    check('the landscape section has its own left margin', frames.map((frame) => frame.MarginLeft), [96, 96, 192, 192, 96]);
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
    // (After two landscape pages the last page does not start at a multiple of a portrait page's stride.)
    check('each line is on the page its y is in', [one.Page, two.Page, three.Page], [pageOf(one.Y) + 1, pageOf(two.Y) + 1, pageOf(three.Y) + 1]);
    check('the last section is on the fifth page, which starts after two shorter ones', [three.Page, frames[4].Top < 4 * frames[1].Top, Math.floor(three.Y / frames[1].Top) + 1], [5, true, 4]);
    const onPage = (line) => Math.round((line.Y - frames[pageOf(line.Y)].Top) - line.PageY);
    check('a line\'s place on its page is measured from that page\'s top (the same offset on every page)', [onPage(two), onPage(three), three.PageY > 0 && three.PageY < 300], [onPage(one), onPage(one), true]);
    const information = Word.getRangeInformation(editor, three.BlockId, 0, 1);
    check('getRangeInformation: the page of a range after the landscape pages', [information.wdActiveEndPageNumber, information.wdNumberOfPagesInDocument], [5, 5]);
    const bounds = Word.getRangeBounds(editor, three.BlockId, 0, 1);
    check('getRangeBounds: its page and its place on it', [bounds.Page, Math.round(bounds.PageY - (bounds.Y - frames[4].Top))], [5, 0]);
} catch (e) { fail(e); }
finish();
