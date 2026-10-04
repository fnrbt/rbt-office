// Resizing and selecting a picture with the pointer. Where the picture and
// its corner handle are is read from the PAINTED canvas (the picture is a
// solid colour nothing else uses), never from the engine's own hit test or
// layout: real pointer events go to those screen positions, in whole-canvas
// mode and through a viewport (100% and 150% zoom at devicePixelRatio 2,
// scrolled). The stored size must change by what was dragged, and one undo
// must restore it. Also: the cursor over the handle / the picture / the text,
// and a click on the picture selects it without resizing it.
import { check, note, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    document.getElementById('c').style.display = 'none';
    // A 120 x 80 picture of one colour.
    const source = (() => { const c = document.createElement('canvas'); c.width = 12; c.height = 8; const g = c.getContext('2d'); g.fillStyle = '#FF00FF'; g.fillRect(0, 0, 12, 8); return c.toDataURL('image/png'); })();
    const picture = new Image(); picture.src = source; await picture.decode();
    const modes = [
        { name: 'whole canvas', viewport: null, style: 'position:absolute;left:0;top:0' },
        { name: 'viewport 100%, scrolled', viewport: { scale: 1, devicePixelRatio: 1, scrollTop: 420, scrollLeft: 30, width: 900, height: 600, originX: 40, originY: 25 }, style: 'position:fixed;left:60px;top:40px' },
        { name: 'viewport 150% at 2x, scrolled', viewport: { scale: 1.5, devicePixelRatio: 2, scrollTop: 760, scrollLeft: 90, width: 900, height: 640, originX: 0, originY: 0 }, style: 'position:fixed;left:35px;top:55px' },
    ];
    const results = [];
    let current = null;
    // The picture's box on the canvas, in CSS pixels of the page, from its pixels.
    const paintedBox = (canvas) => {
        const d = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data, w = canvas.width;
        let left = 1e9, right = -1, top = 1e9, bottom = -1;
        for (let i = 0; i < d.length; i += 4) if (d[i] > 240 && d[i + 1] < 16 && d[i + 2] > 240) { const x = (i / 4) % w, y = Math.floor(i / 4 / w); if (x < left) left = x; if (x > right) right = x; if (y < top) top = y; if (y > bottom) bottom = y; }
        if (right < 0) return null;
        const rect = canvas.getBoundingClientRect(), kx = rect.width / canvas.width, ky = rect.height / canvas.height;
        return { left: rect.left + left * kx, top: rect.top + top * ky, right: rect.left + (right + 1) * kx, bottom: rect.top + (bottom + 1) * ky, width: (right - left + 1) * kx, height: (bottom - top + 1) * ky };
    };
    // The size the document stores for the picture (layout pixels).
    const storedSize = (editor) => { const run = Word.exportDocumentPackage(editor).Paragraphs.flatMap((p) => p.Runs ?? []).find((r) => r.IsImage); return [Math.round(run.Width ?? run.ImageWidth), Math.round(run.Height ?? run.ImageHeight)]; };
    const start = (mode) => {
        for (const canvas of document.querySelectorAll('canvas.mode')) canvas.remove();
        const canvas = document.createElement('canvas');
        canvas.className = 'mode'; canvas.setAttribute('style', mode.style);
        document.body.appendChild(canvas);
        const editor = Word.initializeCanvas(canvas);
        Word.setImageCache(editor, new Map([[source, picture]]));
        const paragraphs = Array.from({ length: 34 }, (_, i) => `Paragraph ${i + 1} of the text before and after the picture.`);
        Word.setParagraphs(editor, paragraphs);
        const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
        Word.setSelection(editor, ids[26], 0, ids[26], 0);
        Word.insertImage(editor, source, 120, 80, 'a picture');
        Word.setSelection(editor, ids[2], 0, ids[2], 0);
        if (mode.viewport) Word.setViewport(editor, mode.viewport);
        else { const size = Word.getDocumentSize(editor); canvas.width = size.Width + 40; canvas.height = size.Height; Word.repaint(editor); }
        const scale = mode.viewport ? mode.viewport.scale : 1;
        current = { mode, canvas, editor, scale, record: { name: mode.name, undoBefore: 0 } };
        results.push(current.record);
        const box = paintedBox(canvas);
        current.record.painted = box && [Math.round(box.width), Math.round(box.height)];
        current.record.stored = storedSize(editor);
        current.box = box;
        return box;
    };
    const undoDepth = (editor) => { let n = 0; while (Word.undo(editor)) n++; for (let i = 0; i < n; i++) Word.redo(editor); return n; };
    window.snap = (name) => {
        const { canvas, editor, record } = current;
        if (name === 'cursor-handle' || name === 'cursor-body' || name === 'cursor-text') record[name] = canvas.style.cursor;
        else if (name === 'clicked') { const s = Word.getSelection(editor); record.clicked = { size: storedSize(editor), undo: undoDepth(editor) - record.undoBefore, selected: s && s.AnchorBlockId === s.FocusBlockId ? Math.abs(s.FocusOffset - s.AnchorOffset) : -1, painted: (() => { const b = paintedBox(canvas); return [Math.round(b.width), Math.round(b.height)]; })() }; }
        else if (name === 'resized') { const b = paintedBox(canvas); record.resized = { size: storedSize(editor), painted: [Math.round(b.width), Math.round(b.height)], undo: undoDepth(editor) - record.undoBefore }; }
        else if (name === 'undone') { record.undone = [Word.undo(editor), storedSize(editor), Word.undo(editor) && false]; Word.redo(editor); }
    };
    // The actions of one mode: computed when it starts, from what is painted then.
    const dragBy = [45, 30];
    window.stepsFor = (index) => {
        const box = start(modes[index]);
        if (!box) return [];
        current.record.undoBefore = undoDepth(current.editor);
        const corner = [box.right - 2, box.bottom - 2], middle = [(box.left + box.right) / 2, (box.top + box.bottom) / 2];
        return [
            ['hover', corner[0], corner[1]], ['eval', 'snap("cursor-handle")'],
            ['hover', middle[0], middle[1]], ['eval', 'snap("cursor-body")'],
            ['hover', box.left + 20, box.top - 30 * current.scale], ['eval', 'snap("cursor-text")'],
            ['click', middle[0], middle[1]], ['eval', 'snap("clicked")'],
            ['down', corner[0], corner[1]], ['move', corner[0] + dragBy[0], corner[1] + dragBy[1]], ['up'], ['eval', 'snap("resized")'],
            ['eval', 'snap("undone")'],
        ];
    };
    window.finish = () => {
        try {
            for (const r of results) {
                const scale = modes.find((m) => m.name === r.name).viewport?.scale ?? 1;
                note(`${r.name}: ${JSON.stringify(r)}`);
                check(`${r.name}: the picture is painted at its stored size`, [r.stored, r.painted && [Math.abs(r.painted[0] - 120 * scale) <= 2, Math.abs(r.painted[1] - 80 * scale) <= 2]], [[120, 80], [true, true]]);
                check(`${r.name}: cursors over the corner handle, the picture and the text`, [r['cursor-handle'], r['cursor-body'], r['cursor-text']], ['nwse-resize', 'move', 'text']);
                // (The hover outline and its handles are drawn over the picture's edge: a few of its pixels are covered.)
                check(`${r.name}: a click on the picture selects it and changes nothing`, [r.clicked.size, r.clicked.undo, r.clicked.selected, Math.abs(r.clicked.painted[0] - r.painted[0]) <= 4 && Math.abs(r.clicked.painted[1] - r.painted[1]) <= 4], [[120, 80], 0, 1, true]);
                // The larger of the two deltas, in layout pixels, is added to both sides (the aspect is kept by the handle).
                const delta = Math.max(dragBy[0], dragBy[1]) / scale;
                check(`${r.name}: dragging the painted corner by ${dragBy} px resizes the stored picture by ${delta.toFixed(0)}`, r.resized && [Math.abs(r.resized.size[0] - (120 + delta)) <= 1, Math.abs(r.resized.size[1] - (80 + delta)) <= 1], [true, true]);
                check(`${r.name}: and the painted one with it`, r.resized && [Math.abs(r.resized.painted[0] - (120 + delta) * scale) <= 5, Math.abs(r.resized.painted[1] - (80 + delta) * scale) <= 5], [true, true]);
                check(`${r.name}: as one undo step, which restores the size`, [r.resized && r.resized.undo, r.undone && r.undone.slice(0, 2)], [1, [true, [120, 80]]]);
            }
            check('all three modes ran', results.length, 3);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [['then', 'window.stepsFor(0)'], ['then', 'window.stepsFor(1)'], ['then', 'window.stepsFor(2)']];
} catch (e) { fail(e); finish(); }
