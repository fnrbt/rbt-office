// The fixture the native painter is tested with (tests/word/native-painter.fsx
// renders it through word-native to a bitmap): the browser lays it out as
// tests/word/native-painter/expected.json says - the page, where each line of
// text is, the rectangle of the shape, the table's cells. The same numbers are
// what the native test holds the painted pixels against, so the two painters
// are compared through them.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const text = async (url) => new TextDecoder().decode(await bytesOf(url));
    const documentXml = await text('/repo/tests/word/native-painter/document.xml');
    const expected = JSON.parse(await text('/repo/tests/word/native-painter/expected.json'));
    const zip = new window.JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/document.xml', documentXml);
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    Word.completeLayout(editor);
    const frame = Word.getPageFrames(editor)[0];
    const top = frame.Top ?? frame.Y ?? 0, left = frame.Left ?? frame.X ?? 0;
    const round = (value) => Math.round(value * 100) / 100;
    // (Relative to the page: the native painter's bitmap is the page.)
    const lines = Word.getLayoutLines(editor).map((line) => ({ text: line.Text, x: round(line.X - left), y: round(line.Y - top), height: round(line.Height) }));
    const shape = Word.getShapeRects(editor).map((rect) => ({ x: round(rect.X - left), y: round(rect.Y - top), width: round(rect.Width), height: round(rect.Height) }))[0];
    const got = { page: { width: frame.Width, height: frame.Height }, lines, shape };
    note('the browser has: ' + JSON.stringify(got));
    const near = (a, b) => Math.abs(a - b) <= 0.5;
    check('the page', [got.page.width, got.page.height], [expected.page.width, expected.page.height]);
    check('the lines of text', lines.map((line) => line.text), expected.lines.map((line) => line.text));
    check('each line where the fixture\'s numbers say (within half a pixel)', lines.map((line, index) => { const want = expected.lines[index] ?? {}; return near(line.x, want.x) && near(line.y, want.y) && near(line.height, want.height); }), lines.map(() => true));
    check('the shape\'s rectangle', shape && [near(shape.x, expected.shape.x), near(shape.y, expected.shape.y), near(shape.width, expected.shape.width), near(shape.height, expected.shape.height)], [true, true, true, true]);
} catch (e) { fail(e); }
finish();
