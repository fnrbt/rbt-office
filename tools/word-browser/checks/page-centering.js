// Through a viewport, a page narrower than the widest one (a portrait section
// among landscape ones) is centred in the document's width; getPageFrames
// says where each page is, and pointer positions, the caret and hit tests
// follow. Nothing of a page is painted in the gap between pages.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initialize('c');
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/repros/word/landscape-section.docx')));
    const wholeCanvas = Word.getPageFrames(editor);
    check('whole-canvas mode: every page at x = 0', wholeCanvas.map((f) => f.X), wholeCanvas.map(() => 0));
    const size = Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1300, height: 900, originX: 0, originY: 0 });
    const frames = Word.getPageFrames(editor);
    note('frames ' + JSON.stringify(frames.map((f) => [f.X, Math.round(f.Top), Math.round(f.Width), Math.round(f.Height)])));
    const widest = Math.max(...frames.map((f) => f.Width));
    check('the document is as wide as its widest page', size.Width, widest);
    check('each page is centred in that width', frames.map((f) => f.X), frames.map((f) => Math.floor((widest - f.Width) / 2)));
    check('a narrower page is moved right', frames.some((f) => f.X > 50), true);
    const ctx = canvas.getContext('2d');
    const px = (x, y) => Array.from(ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data.slice(0, 3));
    const white = [255, 255, 255], desk = [238, 240, 243];
    for (const [index, frame] of frames.entries()) {
        Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: frame.Top, scrollLeft: 0, width: 1300, height: 900, originX: 0, originY: 0 });
        check(`page ${index}: paper inside its centred sheet`, [px(frame.X + 6, 200), px(frame.X + frame.Width - 6, 200)], [white, white]);
        if (frame.X > 10) check(`page ${index}: desk beside it`, [px(frame.X - 6, 200), px(frame.X + frame.Width + 6, 200)], [desk, desk]);
        // A line of the page: a click on it (canvas pixels) finds its paragraph, and the caret is drawn there.
        const line = Word.getLayoutLines(editor).find((l) => l.Y >= frame.Top && l.Y < frame.Top + frame.Height && (l.Text || '').trim().length > 3);
        if (!line) continue;
        const hit = Word.hitTestViewport(editor, frame.X + line.X + 4, line.Y - frame.Top + line.Height / 2);
        check(`page ${index}: a hit test on the centred page finds its paragraph`, hit && hit.BlockId, line.BlockId);
        Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0);
        const caret = Word.getCaretRect(editor);
        check(`page ${index}: the caret is drawn on the centred page`, [Math.abs(caret.ViewportX - (caret.X + frame.X)) < 0.01, caret.ViewportX >= frame.X, caret.ViewportX <= frame.X + frame.Width], [true, true, true]);
        const there = Word.documentToViewport(editor, 10, frame.Top + 10), back = Word.viewportToDocument(editor, there.X, there.Y);
        check(`page ${index}: documentToViewport and viewportToDocument are inverse`, [Math.round(there.X), Math.round(back.X), Math.round(back.Y)], [10 + frame.X, 10, Math.round(frame.Top + 10)]);
    }
    // A real click on the centred page puts the caret under the pointer.
    const narrow = frames.findIndex((f) => f.X > 50);
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: frames[narrow].Top, scrollLeft: 0, width: 1300, height: 900, originX: 0, originY: 0 });
    const target = Word.getLayoutLines(editor).find((l) => l.Y >= frames[narrow].Top && l.Y < frames[narrow].Top + frames[narrow].Height && (l.Text || '').trim().length > 10);
    const box = canvas.getBoundingClientRect();
    window.finish = () => {
        try {
            const selection = Word.getSelection(editor);
            check('a click on the centred page lands in the line under the pointer', selection && selection.FocusBlockId, target.BlockId);
            // Nothing is painted between pages: sample-formatted with a measurer that puts its nested table across a page break.
            Word.setViewport(editor, null);
        } catch (e) { fail(e); }
        (async () => {
            try {
                const k = 0.45;
                Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents/sample-formatted.docx')));
                Word.setTextMeasurers(editor, (text, style) => text.length * style.FontSize * k, (text, style, start, length) => length * style.FontSize * k);
                const pages = Word.getPageFrames(editor);
                let stray = 0, gaps = 0;
                for (let i = 0; i + 1 < pages.length; i++) {
                    const gapTop = pages[i].Top + pages[i].Height, gapBottom = pages[i + 1].Top;
                    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: gapTop - 100, scrollLeft: 0, width: 1000, height: 400, originX: 0, originY: 0 });
                    const data = ctx.getImageData(0, 101, canvas.width, Math.round(gapBottom - gapTop) - 2).data;
                    gaps++;
                    for (let p = 0; p < data.length; p += 4) if (!(data[p] === 238 && data[p + 1] === 240 && data[p + 2] === 243)) stray++;
                }
                check(`nothing is painted in the ${gaps} gaps between pages`, stray, 0);
            } catch (e) { fail(e); }
            finish();
        })();
    };
    window.drive = [['click', box.left + frames[narrow].X + target.X + 30, box.top + target.Y - frames[narrow].Top + target.Height / 2]];
} catch (e) { fail(e); finish(); }
