// A turned text box that is edited in place: its caret, its selection and
// the frame around it are turned as its text is, and a point of the canvas is
// looked up in its text as the point it is before the turn.
// docx4j-drawingml-wps.docx: <a:xfrm rot="837439"> - 837439 / 60000 = 13.96
// degrees clockwise - about the middle of a box of 965607 x 819303 EMU
// (101.4 x 86.0 px). A second time with rot="5400000" (90 degrees): the line
// of text runs down the page.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d', { willReadFrequently: true });
    const original = await bytesOf('/repo/tests/test-documents-wild/docx4j-drawingml-wps.docx');
    const zip = await window.JSZip.loadAsync(original);
    const xml = await zip.file('word/document.xml').async('string');
    if (!xml.includes('<a:xfrm rot="837439">')) throw new Error('the document changed');
    zip.file('word/document.xml', xml.replace('<a:xfrm rot="837439">', '<a:xfrm rot="5400000">'));
    const quarter = await zip.generateAsync({ type: 'uint8array' });
    const open = async (bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes)); return editor; };
    const region = (rect) => { const reach = Math.ceil(Math.hypot(rect.Width, rect.Height) / 2) + 12; const cx = rect.X + rect.Width / 2, cy = rect.Y + rect.Height / 2; return { x: Math.round(cx - reach), y: Math.round(cy - reach), size: reach * 2, cx, cy }; };
    const pixels = (area) => context.getImageData(area.x, area.y, area.size, area.size).data;
    // The pixels that differ between two pictures of an area: their number, their box and the direction they lie in
    // (the angle of the main axis of their spread, degrees clockwise from the x axis as the canvas has it: y goes down).
    const spread = (area, before, after) => {
        let n = 0, sx = 0, sy = 0; const points = [];
        for (let i = 0; i < before.length; i += 4) {
            if (Math.abs(before[i] - after[i]) + Math.abs(before[i + 1] - after[i + 1]) + Math.abs(before[i + 2] - after[i + 2]) > 24) {
                const x = (i / 4) % area.size, y = Math.floor(i / 4 / area.size);
                points.push([x, y]); n++; sx += x; sy += y;
            }
        }
        if (n === 0) return { n: 0 };
        const mx = sx / n, my = sy / n; let xx = 0, yy = 0, xy = 0, left = Infinity, right = -Infinity, top = Infinity, bottom = -Infinity;
        for (const [x, y] of points) { xx += (x - mx) * (x - mx); yy += (y - my) * (y - my); xy += (x - mx) * (y - my); left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y); }
        const angle = 0.5 * Math.atan2(2 * xy, xx - yy) * 180 / Math.PI;
        // The ends of the spread along its main axis.
        const ax = Math.cos(angle * Math.PI / 180), ay = Math.sin(angle * Math.PI / 180); let low = Infinity, high = -Infinity;
        for (const [x, y] of points) { const t = (x - mx) * ax + (y - my) * ay; low = Math.min(low, t); high = Math.max(high, t); }
        return { n, angle, width: right - left + 1, height: bottom - top + 1, middle: [area.x + mx, area.y + my], start: [area.x + mx + ax * low, area.y + my + ay * low], end: [area.x + mx + ax * high, area.y + my + ay * high] };
    };
    const turned = (degrees, cx, cy, x, y) => { const a = degrees * Math.PI / 180, dx = x - cx, dy = y - cy; return [cx + dx * Math.cos(a) - dy * Math.sin(a), cy + dx * Math.sin(a) + dy * Math.cos(a)]; };

    // ---- 13.96 degrees ----
    {
        const editor = await open(original);
        const rect = Word.getShapeRects(editor).find((item) => item.Text.length > 0);
        const area = region(rect);
        const body = pixels(area);
        check('the box is entered', [Word.enterTextBox(editor, rect.Id), Word.getCaretContext(editor).Story], [true, 'textbox']);
        const entered = pixels(area);
        // (Entering the box draws its frame and its caret: nothing else changes.)
        const frame = spread(area, body, entered);
        const angle = 837439 / 60000 * Math.PI / 180;
        const wantWidth = (rect.Width + 5) * Math.cos(angle) + (rect.Height + 5) * Math.sin(angle), wantHeight = (rect.Width + 5) * Math.sin(angle) + (rect.Height + 5) * Math.cos(angle);
        note(`frame: ${frame.width} x ${frame.height} (a turned frame: ${wantWidth.toFixed(1)} x ${wantHeight.toFixed(1)}; upright: ${rect.Width + 5} x ${rect.Height + 5})`);
        check('the frame around the box is turned with it (its extent on the canvas)', [Math.abs(frame.width - wantWidth) <= 3, Math.abs(frame.height - wantHeight) <= 3], [true, true]);
        // (The text is two lines, "Funky" and "chicken": the second is selected - one line, longer than high.)
        const paragraph = Word.getSelection(editor).FocusBlockId;
        // (Compared with the caret where the selection starts: the caret alone does not count.)
        Word.setSelection(editor, paragraph, 6, paragraph, 6);
        const caretOnly = pixels(area);
        check('the text of the box', [rect.Text, Word.setSelection(editor, paragraph, 6, paragraph, 13)], ['Funky chicken', true]);
        const selected = spread(area, caretOnly, pixels(area));
        note(`selection: ${selected.n} pixels at ${selected.angle.toFixed(2)} degrees`);
        check('the selection is turned as the text is', [selected.n > 200, Math.abs(selected.angle - 837439 / 60000) < 2.5], [true, true]);
        const caret = Word.getCaretRect(editor);
        check('getCaretRect says how the caret is turned', [Math.round(caret.Rotation * 100) / 100, caret.Height > 5], [13.96, true]);
        Word.exitTextBox(editor);
        check('in the body the caret is upright', [Word.getCaretContext(editor).Story, (Word.getCaretRect(editor) ?? { Rotation: 0 }).Rotation], ['body', 0]);
    }

    // ---- a quarter turn ----
    const editor = await open(quarter);
    const rect = Word.getShapeRects(editor).find((item) => item.Text.length > 0);
    const area = region(rect);
    const length = rect.Text.length;
    check('the box is entered', [Word.enterTextBox(editor, rect.Id), length > 5], [true, true]);
    const entered = pixels(area);
    const paragraph = Word.getSelection(editor).FocusBlockId;
    Word.setSelection(editor, paragraph, 6, paragraph, 6);
    const caretOnly = pixels(area);
    Word.setSelection(editor, paragraph, 6, paragraph, 13);
    const selected = spread(area, caretOnly, pixels(area));
    note(`selection: ${selected.n} pixels, ${selected.width} x ${selected.height}, at ${selected.angle.toFixed(2)} degrees`);
    check('the selection of a line turned a quarter is higher than wide', [selected.n > 200, selected.height > 1.5 * selected.width, Math.abs(Math.abs(selected.angle) - 90) < 3], [true, true, true]);
    // (Turned a quarter clockwise the line starts at the top.)
    const [top, bottom] = selected.start[1] < selected.end[1] ? [selected.start, selected.end] : [selected.end, selected.start];
    const hitTop = Word.hitTestPosition(editor, top[0], top[1] + 2), hitBottom = Word.hitTestPosition(editor, bottom[0], bottom[1] - 2), hitMiddle = Word.hitTestPosition(editor, selected.middle[0], selected.middle[1]);
    check('a point at the top of the line is its start, one at the bottom its end, one in the middle in between', [hitTop.BlockId === paragraph, hitTop.Offset >= 6 && hitTop.Offset <= 7, hitBottom.Offset >= length - 1, hitMiddle.Offset > 7 && hitMiddle.Offset < length - 1], [true, true, true, true]);
    Word.setSelection(editor, paragraph, 6, paragraph, 6);
    const first = Word.getCaretRect(editor);
    Word.setSelection(editor, paragraph, length, paragraph, length);
    const last = Word.getCaretRect(editor);
    check('the caret goes down the page along the line', [Math.round(first.Rotation), last.Y + last.Height / 2 > first.Y + first.Height / 2 + 20, Math.abs(last.X - first.X) < 3], [90, true, true]);
    Word.exitTextBox(editor);
    // A point in the turned box that is not in the upright one (the box is wider than high: above its upright top), and one the other way round.
    const cx = rect.X + rect.Width / 2, cy = rect.Y + rect.Height / 2;
    const state = {};
    window.steps = { inside: () => { state.inside = Word.getCaretContext(editor).Story; Word.exitTextBox(editor); }, outside: () => { state.outside = Word.getCaretContext(editor).Story; } };
    const box = canvas.getBoundingClientRect();
    window.drive = [
        ['click', Math.round(box.left + cx), Math.round(box.top + cy - rect.Width / 2 + 3)], ['eval', 'window.steps.inside()'],
        ['click', Math.round(box.left + cx - rect.Width / 2 + 3), Math.round(box.top + cy)], ['eval', 'window.steps.outside()'],
    ];
    window.finish = () => {
        check('the box is wider than high', rect.Width > rect.Height + 8, true);
        check('a click in the turned box enters it; one where only the upright box would be does not', [state.inside, state.outside], ['textbox', 'body']);
        finish();
    };
} catch (e) { fail(e); finish(); }
