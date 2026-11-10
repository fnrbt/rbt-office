// Groups and drawing canvases (the fixture synthetic/group-shapes: a wpg group of a red text box, a picture and a
// group in it whose shape is green, with its VML fallback; a yellow wpc canvas with a blue rectangle; a VML group of a
// magenta text box and a cyan oval) in pixels: each shape of a group where the group's coordinates put it, the
// canvas's background below its shape; getShapeRects lists each group once (IsGroup) with its shapes (Children);
// enterTextBox refuses a group. Moved and resized through the API (moveShape, resizeShape) a group is drawn there and
// saved as the element it was read from with its place and size changed in it - the anchor's offsets, wp:extent and
// the group's own a:ext (its child extent as it was: its shapes scale with it), the VML group's style - and nothing
// else of it; one undo and it is saved as it was read. (Groups were drawn as empty shapes, and saved from the model:
// their shapes were lost.)
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const context = canvas.getContext('2d', { willReadFrequently: true });
    // (The picture of the group: one orange pixel, stretched.)
    const source = new OffscreenCanvas(1, 1);
    const sourceContext = source.getContext('2d');
    sourceContext.fillStyle = '#FF8000'; sourceContext.fillRect(0, 0, 1, 1);
    const png = new Uint8Array(await (await source.convertToBlob({ type: 'image/png' })).arrayBuffer());
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/group-shapes.docx')));
    zip.file('word/media/body.png', png);
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const original = await zip.file('word/document.xml').async('string');
    const xmlOf = async (data) => (await window.JSZip.loadAsync(data)).file('word/document.xml').async('string');
    const open = async (data = bytes) => {
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data));
        Word.completeLayout(editor);
        // (The picture decoded before it is painted.)
        const group = Word.exportDocumentPackage(editor).Paragraphs.find((item) => item.Children && item.Children.length === 3);
        const src = group && group.Children.map((shape) => shape.ImageSrc).find((value) => value);
        if (src) { const image = new Image(); image.src = src; await image.decode(); Word.setImageCache(editor, new Map([[src, image]])); }
        await new Promise((resolve) => setTimeout(resolve, 100));
        return editor;
    };
    const save = async (editor) => new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const colour = (x, y) => {
        const d = context.getImageData(Math.round(x), Math.round(y), 1, 1).data;
        const [r, g, b] = d;
        if (r > 200 && g < 80 && b < 80) return 'red';
        if (r < 80 && g > 200 && b < 80) return 'green';
        if (r < 80 && g < 80 && b > 200) return 'blue';
        if (r > 200 && g > 200 && b < 80) return 'yellow';
        if (r > 200 && g < 80 && b > 200) return 'magenta';
        if (r < 80 && g > 200 && b > 200) return 'cyan';
        if (r > 200 && g > 90 && g < 170 && b < 80) return 'orange';
        if (r > 240 && g > 240 && b > 240) return 'white';
        return `rgb(${r},${g},${b})`;
    };
    // The shapes on the canvas: X from the page's left edge plus the page's own place.
    const shapes = (editor) => {
        const frame = Word.getPageFrames(editor)[0];
        const left = frame.X ?? 0;
        return Word.getShapeRects(editor).map((rect) => ({ ...rect, X: rect.X + left }));
    };
    const byName = (editor) => {
        const all = shapes(editor);
        return { group: all.find((rect) => rect.IsGroup && (rect.Children ?? []).length === 3), canvasShape: all.find((rect) => rect.IsGroup && (rect.Children ?? []).length === 1), vml: all.find((rect) => rect.IsGroup && (rect.Children ?? []).length === 2), all };
    };
    const at = (rect, fx, fy) => colour(rect.X + rect.Width * fx, rect.Y + rect.Height * fy);
    const round = (value) => Math.round(value * 100) / 100;
    const box = (rect) => rect && [rect.X, rect.Y, rect.Width, rect.Height].map(round);
    // The colours of a group's quarters (and of the canvas's middle right): where its shapes are drawn.
    const groupColours = (rect) => [at(rect, 0.25, 0.25), at(rect, 0.75, 0.25), at(rect, 0.25, 0.75), at(rect, 0.75, 0.75)];
    const canvasColours = (rect) => [at(rect, 0.2, 0.2), at(rect, 0.75, 0.75)];
    // (An element compared without its namespace declarations and with its attributes in one order: the writer keeps
    // a shape as the reader has it, its namespaces declared on it.)
    const normal = (text) => (text.match(/<[^>]+>|[^<]+/g) ?? []).map((tag) => tag.startsWith('<') ? tag.replace(/ xmlns:\w+="[^"]*"/g, '').replace(/^(<[\w:]+)((?: [\w:]+="[^"]*")*)(\s*\/?>)$/, (_, open, attributes, close) => open + (attributes.match(/ [\w:]+="[^"]*"/g) ?? []).sort().join('') + close) : tag).join('');
    const between = (xml, start, end) => { const from = xml.indexOf(start); const to = xml.indexOf(end, from); return from < 0 || to < 0 ? '' : xml.slice(from, to + end.length); };
    const wpgOf = (xml) => normal(between(xml, '<wpg:wgp', '</wpg:wgp>'));
    const wpcOf = (xml) => normal(between(xml, '<wpc:wpc', '</wpc:wpc>'));
    const vmlGroupOf = (xml) => { const start = /<v:group [^>]*id="_x0000_s1200"/.exec(xml); return start ? normal(between(xml.slice(start.index), '<v:group', '</v:group>')) : ''; };
    const fallbackOf = (xml) => normal(between(xml, '<mc:Fallback>', '</mc:Fallback>'));

    const editor = await open();
    const { group, canvasShape, vml, all } = byName(editor);
    note(`group ${JSON.stringify(box(group))}, canvas ${JSON.stringify(box(canvasShape))}, VML group ${JSON.stringify(box(vml))}`);
    check('getShapeRects: the three groups, each once, as large as their anchors say', [all.length, box(group)?.slice(2), box(canvasShape)?.slice(2), box(vml)?.slice(2)], [3, [200, 100], [200, 100], [200, 100]]);
    check('a group\'s shapes: where the group\'s coordinates put them, at half their size in them; the group in it with its own shape',
        [group.Children.map((shape) => [shape.X - group.X, shape.Y - group.Y, shape.Width, shape.Height].map(round)), group.Children.map((shape) => shape.IsGroup), group.Children[0].Text,
         group.Children[2].Children.map((shape) => [shape.X - group.X, shape.Y - group.Y, shape.Width, shape.Height].map(round))],
        [[[0, 0, 100, 50], [100, 0, 100, 50], [0, 50, 200, 50]], [false, false, true], 'group text', [[100, 50, 100, 50]]]);
    check('the canvas\'s shape and the VML group\'s, in their places', [canvasShape.Children.map((shape) => [shape.X - canvasShape.X, shape.Y - canvasShape.Y, shape.Width, shape.Height].map(round)), vml.Children.map((shape) => [shape.X - vml.X, shape.Y - vml.Y, shape.Width, shape.Height].map(round)), vml.Children[0].Text],
        [[[100, 50, 100, 50]], [[0, 0, 100, 50], [100, 50, 100, 50]], 'vml text']);
    // ---- pixels ------------------------------------------------------
    check('the group in pixels: the red text box at the top left, the picture at the top right, below them the inner group\'s green shape at the right', groupColours(group), ['red', 'orange', 'white', 'green']);
    check('the canvas in pixels: yellow, its blue rectangle at its bottom right', canvasColours(canvasShape), ['yellow', 'blue']);
    check('the VML group in pixels: magenta at the top left, the cyan oval at the bottom right', [at(vml, 0.25, 0.25), at(vml, 0.75, 0.75), at(vml, 0.75, 0.25)], ['magenta', 'cyan', 'white']);
    check('enterTextBox refuses a group and a canvas', [Word.enterTextBox(editor, group.Id), Word.enterTextBox(editor, canvasShape.Id), Word.enterTextBox(editor, vml.Id)], [false, false, false]);
    check('an unedited save is the document as it was read', await xmlOf(await save(editor)) === original, true);
    // ---- moved -------------------------------------------------------
    check('moveShape: each group 30 to the right and 20 down', [Word.moveShape(editor, group.Id, 30, 20), Word.moveShape(editor, canvasShape.Id, 30, 20), Word.moveShape(editor, vml.Id, 30, 20)], [true, true, true]);
    const moved = byName(editor);
    check('moved: there', [box(moved.group), box(moved.canvasShape), box(moved.vml)],
        [box({ ...group, X: group.X + 30, Y: group.Y + 20 }), box({ ...canvasShape, X: canvasShape.X + 30, Y: canvasShape.Y + 20 }), box({ ...vml, X: vml.X + 30, Y: vml.Y + 20 })]);
    check('moved: drawn there, its shapes with it', [groupColours(moved.group), canvasColours(moved.canvasShape), at(moved.vml, 0.25, 0.25)], [['red', 'orange', 'white', 'green'], ['yellow', 'blue'], 'magenta']);
    const movedXml = await xmlOf(await save(editor));
    check('moved, saved: the anchors\' offsets (10 + 30 px, 20 + 20 px; the canvas 0 + 30 px, 300 + 20 px in EMU)',
        [movedXml.includes('<wp:positionH relativeFrom="column"><wp:posOffset>381000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>381000</wp:posOffset></wp:positionV>'),
         movedXml.includes('<wp:positionH relativeFrom="column"><wp:posOffset>285750</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>3048000</wp:posOffset></wp:positionV>')],
        [true, true]);
    check('moved, saved: the groups\' shapes are the XML they were', [wpgOf(movedXml) === wpgOf(original), wpcOf(movedXml) === wpcOf(original), wpgOf(movedXml).length > 1000], [true, true, true]);
    check('moved, saved: the VML group (alone, and the fallback) says the new place in its style',
        [/<v:group [^>]*id="_x0000_s1200"[^>]*style="position:absolute;margin-left:322.5pt;margin-top:135pt;width:150pt;height:75pt;z-index:15"/.test(movedXml),
         /<v:group [^>]*id="Group 11"[^>]*style="position:absolute;margin-left:30pt;margin-top:30pt;width:150pt;height:75pt;z-index:11;mso-position-horizontal-relative:text;mso-position-vertical-relative:text"/.test(movedXml),
         vmlGroupOf(movedXml).length > 200 && vmlGroupOf(movedXml).replace(/ style="[^"]*"/, '') === vmlGroupOf(original).replace(/ style="[^"]*"/, '')],
        [true, true, true]);
    check('one undo of each: saved as it was read', [Word.undo(editor), Word.undo(editor), Word.undo(editor), await xmlOf(await save(editor)) === original], [true, true, true, true]);
    // ---- resized -----------------------------------------------------
    check('resizeShape: the group twice as large, the canvas and the VML group too', [Word.resizeShape(editor, group.Id, 400, 200), Word.resizeShape(editor, canvasShape.Id, 400, 200), Word.resizeShape(editor, vml.Id, 300, 150)], [true, true, true]);
    const resized = byName(editor);
    check('resized: its shapes scale with the group (its child coordinates as they were); the canvas\'s stay as they were',
        [box(resized.group), resized.group.Children.map((shape) => [shape.X - resized.group.X, shape.Y - resized.group.Y, shape.Width, shape.Height].map(round)),
         resized.canvasShape.Children.map((shape) => [shape.X - resized.canvasShape.X, shape.Y - resized.canvasShape.Y, shape.Width, shape.Height].map(round))],
        [[group.X, group.Y, 400, 200].map(round), [[0, 0, 200, 100], [200, 0, 200, 100], [0, 100, 400, 100]], [[100, 50, 100, 50]]]);
    check('resized: drawn so', [groupColours(resized.group), at(resized.canvasShape, 0.3, 0.4), at(resized.canvasShape, 0.75, 0.75)], [['red', 'orange', 'white', 'green'], 'blue', 'yellow']);
    const resizedXml = await xmlOf(await save(editor));
    check('resized, saved: wp:extent and the group\'s own a:ext (400 x 200 px), its child extent as it was; the canvas\'s wp:extent',
        [resizedXml.includes('<wp:extent cx="3810000" cy="1905000"/><wp:effectExtent b="0" l="0" r="0" t="0"/><wp:wrapNone/><wp:docPr id="11"'),
         resizedXml.includes('<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/><a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm></wpg:grpSpPr>'),
         resizedXml.includes('<wp:extent cx="3810000" cy="1905000"/><wp:effectExtent b="0" l="0" r="0" t="0"/><wp:wrapNone/><wp:docPr id="14"'), wpcOf(resizedXml) === wpcOf(original)],
        [true, true, true, true]);
    check('resized, saved: the shapes of the group as they were (the outer a:ext aside); the VML groups\' sizes in their styles',
        [wpgOf(resizedXml).replace('<a:ext cx="3810000" cy="1905000"/><a:chOff', '<a:ext cx="1905000" cy="952500"/><a:chOff') === wpgOf(original),
         /<v:group [^>]*id="_x0000_s1200"[^>]*style="position:absolute;margin-left:300pt;margin-top:120pt;width:225pt;height:112.5pt;z-index:15"/.test(resizedXml),
         /<v:group [^>]*id="Group 11"[^>]*style="position:absolute;margin-left:7.5pt;margin-top:15pt;width:300pt;height:150pt;/.test(resizedXml),
         fallbackOf(resizedXml).replace(/ style="position:absolute;margin-left[^"]*"/, '') === fallbackOf(original).replace(/ style="position:absolute;margin-left[^"]*"/, '')],
        [true, true, true, true]);
    const again = await open(await save(editor));
    check('resized, read again: as large as it was saved, its shapes where the editor had them', box(byName(again).group) && byName(again).group.Children.map((shape) => [shape.Width, shape.Height]), resized.group.Children.map((shape) => [shape.Width, shape.Height]));
    Word.dispose(again);
    check('undone: saved as it was read', [Word.undo(editor), Word.undo(editor), Word.undo(editor), await xmlOf(await save(editor)) === original], [true, true, true, true]);
    Word.dispose(editor);
} catch (e) { fail(e); }
finish();
