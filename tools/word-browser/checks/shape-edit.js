// Shapes moved, resized and deleted through the API: moveShape, resizeShape,
// deleteShape (each one undo step), getShapeHandles and hitTestShapeHandle
// for a host that shows a shape selected and drags it. A shape that was read
// and only moved or resized is saved as the element it was read from with
// the offsets and the extent of its anchor changed in it - its geometry, its
// gradient, whatever the model does not hold of it stay the text they were;
// the other shapes are not touched.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/shapes.docx'));
    const xmlOf = async (data) => (await window.JSZip.loadAsync(data)).file('word/document.xml').async('string');
    const original = await xmlOf(bytes);
    const canvas = document.getElementById('c');
    const open = async (data = bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data)); Word.completeLayout(editor); return editor; };
    // The three shapes of the fixture, by their widths: the oval (140), the gradient box (100), the triangle (80).
    const shape = (editor, width) => Word.getShapeRects(editor).find((rect) => Math.abs(rect.Width - width) < 1.5);
    const place = (rect) => rect && [rect.X, rect.Y, rect.Width, rect.Height].map((value) => Math.round(value * 100) / 100);
    const saved = async (editor) => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const again = await open(out);
        const result = { xml: await xmlOf(out), bytes: out, shapes: Word.getShapeRects(again).map(place) };
        Word.dispose(again);
        return result;
    };
    // The element of a shape in document.xml: its mc:AlternateContent, by the shape's name.
    // (A shape in a paragraph that is written from the model is the element as the reader has it: the namespaces it uses declared on it, its attributes in their order.)
    const elementOf = (xml, name) => xml.split(/<mc:AlternateContent(?: [^>]*)?>/).find((piece) => piece.includes(`name="${name}"`)) ?? '';
    const sameShape = (a, b) => { const tags = (text) => (text.match(/<[^>]+>|[^<]+/g) ?? []).map((tag) => tag.startsWith('<') ? tag.replace(/ xmlns:\w+="[^"]*"/g, '').replace(/^(<[\w:]+)((?: [\w:]+="[^"]*")*)(\s*\/?>)$/, (_, open, attributes, close) => open + (attributes.match(/ [\w:]+="[^"]*"/g) ?? []).sort().join('') + close) : tag).join(''); return tags(a) === tags(b); };
    const count = (xml, text) => xml.split(text).length - 1;

    {
        const editor = await open();
        const box = shape(editor, 100), oval = shape(editor, 140), triangle = shape(editor, 80);
        const before = Word.getShapeRects(editor).map(place);
        check('the fixture has its three shapes', [!!oval, !!box, !!triangle, before.length], [true, true, true, 3]);
        // ---- moved -------------------------------------------------------
        const revision = Word.getRevision(editor);
        check('moveShape(30, 20): the shape is 30 to the right and 20 down, the others where they were',
            [Word.moveShape(editor, box.Id, 30, 20), place(shape(editor, 100)), place(shape(editor, 140)), place(shape(editor, 80)), Word.getRevision(editor) > revision, Word.canUndo(editor)],
            [true, [box.X + 30, box.Y + 20, box.Width, box.Height].map((value) => Math.round(value * 100) / 100), place(oval), place(triangle), true, true]);
        const moved = await saved(editor);
        const element = elementOf(moved.xml, 'Gradient box');
        check('saved in the element it was read from: the offsets of the anchor changed (1905000 + 30 px and 0 + 20 px in EMU), its frames, gradient and name as they were',
            [/<wp:positionH relativeFrom="margin"><wp:posOffset>2190750<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>190500<\/wp:posOffset><\/wp:positionV>/.test(element),
             element.includes('<a:lin ang="5400000" scaled="0"/>'), element.includes('<wp:docPr id="2" name="Gradient box"/>'), element.includes('<wp:extent cx="952500" cy="476250"/>'), element.includes('behindDoc="1"')],
            [true, true, true, true, true]);
        check('the other shapes are the text they were (their VML fallback too); the moved one has none',
            [sameShape(elementOf(moved.xml, 'Oval'), elementOf(original, 'Oval')), elementOf(moved.xml, 'Triangle') === elementOf(original, 'Triangle'), count(moved.xml, '<mc:Fallback>'), count(original, '<mc:Fallback>'), count(moved.xml, '</mc:AlternateContent>')],
            [true, true, 2, 3, 3]);
        check('read again: the shapes are where the editor has them', moved.shapes, Word.getShapeRects(editor).map(place));
        check('one undo: where it was, and saved as the text that was read', [Word.undo(editor), place(shape(editor, 100)), (await saved(editor)).xml === original], [true, place(box), true]);
        // ---- resized -----------------------------------------------------
        check('resizeShape(120, 60): that size, its corner where it was', [Word.resizeShape(editor, triangle.Id, 120, 60), place(shape(editor, 120))], [true, [triangle.X, triangle.Y, 120, 60].map((value) => Math.round(value * 100) / 100)]);
        check('to the size it has: nothing changes', Word.resizeShape(editor, triangle.Id, 120, 60), false);
        const resized = await saved(editor);
        const triangleXml = elementOf(resized.xml, 'Triangle');
        check('saved in its element: wp:extent and the shape\'s own a:ext (120 x 60 px in EMU), its path as it was',
            [triangleXml.includes('<wp:extent cx="1143000" cy="571500"/>'), triangleXml.includes('<a:ext cx="1143000" cy="571500"/>'), /<a:path (?:w="100" h="100"|h="100" w="100")><a:moveTo><a:pt x="50" y="0"\/>/.test(triangleXml),
             /<wp:positionH relativeFrom="column"><wp:posOffset>95250<\/wp:posOffset>/.test(triangleXml), elementOf(resized.xml, 'Oval') === elementOf(original, 'Oval')],
            [true, true, true, true, true]);
        check('read again: the shapes are where the editor has them', resized.shapes, Word.getShapeRects(editor).map(place));
        check('one undo: the size it had', [Word.undo(editor), place(shape(editor, 80))], [true, place(triangle)]);
        // ---- handles -----------------------------------------------------
        const handles = Word.getShapeHandles(editor, box.Id);
        check('getShapeHandles: eight, each a square of eight pixels around its corner or the middle of its edge',
            [handles.map((handle) => handle.Name), handles.find((handle) => handle.Name === 'nw') && [handles[0].X, handles[0].Y, handles[0].Width, handles[0].Height, handles[0].Cursor],
             (() => { const se = handles.find((handle) => handle.Name === 'se'); return se && [Math.round((se.X + 4) * 100) / 100, Math.round((se.Y + 4) * 100) / 100]; })()],
            [['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'], [box.X - 4, box.Y - 4, 8, 8, 'nwse-resize'], [Math.round((box.X + box.Width) * 100) / 100, Math.round((box.Y + box.Height) * 100) / 100]]);
        check('hitTestShapeHandle: a corner, the middle of an edge, inside the shape, outside it',
            [Word.hitTestShapeHandle(editor, box.Id, box.X + 1, box.Y - 1), Word.hitTestShapeHandle(editor, box.Id, box.X + box.Width, box.Y + box.Height / 2), Word.hitTestShapeHandle(editor, box.Id, box.X + box.Width / 2, box.Y + box.Height / 2), Word.hitTestShapeHandle(editor, box.Id, box.X - 40, box.Y)],
            ['nw', 'e', 'move', '']);
        // ---- refused -----------------------------------------------------
        check('no such shape, no distance: false, and nothing changes',
            [Word.moveShape(editor, 'no-such-shape', 5, 5), Word.moveShape(editor, box.Id, 0, 0), Word.resizeShape(editor, 'no-such-shape', 10, 10), Word.deleteShape(editor, 'no-such-shape'), Word.getShapeHandles(editor, 'no-such-shape').length, (await saved(editor)).xml === original],
            [false, false, false, false, 0, true]);
        check('while the text of a box is edited the shapes are not moved', [Word.enterTextBox(editor, oval.Id), Word.moveShape(editor, box.Id, 5, 5), Word.deleteShape(editor, box.Id), Word.exitTextBox(editor)], [true, false, false, true]);
        // ---- deleted -----------------------------------------------------
        check('deleteShape: the shape and its text are gone', [Word.deleteShape(editor, oval.Id), Word.getShapeRects(editor).length, !!shape(editor, 140), Word.findText(editor, 'in the oval', true).length], [true, 2, false, 0]);
        const deleted = await saved(editor);
        check('saved without it; the other shapes are the text they were', [deleted.xml.includes('name="Oval"'), deleted.xml.includes('in the oval'), elementOf(deleted.xml, 'Triangle') === elementOf(original, 'Triangle'), deleted.shapes.length], [false, false, true, 2]);
        check('one undo: it is back, its text with it', [Word.undo(editor), place(shape(editor, 140)), Word.findText(editor, 'in the oval', true).length], [true, place(oval), 1]);
        Word.dispose(editor);
    }
    {
        // A shape that is aligned (centred between the margins): moved, it is by its offset from the page.
        const zip = await window.JSZip.loadAsync(bytes);
        zip.file('word/document.xml', original.replace('<wp:positionH relativeFrom="page"><wp:posOffset>571500</wp:posOffset></wp:positionH>', '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>'));
        const editor = await open(await zip.generateAsync({ type: 'uint8array' }));
        const oval = shape(editor, 140);
        check('a shape centred between the margins, 100 pixels down its page, is read so', [Math.round(oval.X), Math.round(oval.Y - (Word.getPageFrames(editor)[0].Top ?? 0))], [338, 100]);
        check('an aligned shape moved 10 to the right is 10 to the right', [Word.moveShape(editor, oval.Id, 10, 0), place(shape(editor, 140))], [true, [oval.X + 10, oval.Y, oval.Width, oval.Height].map((value) => Math.round(value * 100) / 100)]);
        const result = await saved(editor);
        const element = elementOf(result.xml, 'Oval');
        check('saved by its offset from the page; its vertical place as it was read',
            [element.includes(`<wp:positionH relativeFrom="page"><wp:posOffset>${Math.round((oval.X + 10) * 9525)}</wp:posOffset></wp:positionH>`), element.includes('<wp:positionV relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionV>'), element.includes('<a:xfrm rot="900000">'), result.shapes],
            [true, true, true, Word.getShapeRects(editor).map(place)]);
        Word.dispose(editor);
    }
} catch (e) { fail(e); }
finish();
