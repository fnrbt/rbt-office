// What was measured is measured otherwise while most of a long document is
// being laid out again with the view far down it (everything made bold on
// page 300: the host gets the bold face and installs measurers that know it -
// setTextMeasurersForFaces - or other measurers altogether, or zooms, or its
// window gets another width). The blocks above the view stay as they were
// laid out before the change: the document keeps its height and its pages
// (it was the three pages the layout from the start had come to, and a host
// that sizes its scroller from it was thrown to the top), the view stays on
// its text, the host is told where to scroll (the ScrollTop notice) when the
// heights above the view are known, and the layout at the end is the one the
// new measurers give a document that was bold from the start.
import { check, note, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const paragraphs = Array.from({ length: 3000 }, (_, i) => `Paragraph ${i + 1}: ` + Array.from({ length: 12 + (i % 9) * 4 }, (_, w) => `word${(i * 7 + w) % 97}`).join(' '));
    // A text stack: bold text is set in the regular face (as wide as it) until it has the bold one.
    const stack = () => {
        const text = { hasBold: false, families: new Set() };
        const width = (style) => (style.Bold && text.hasBold ? 0.62 : 0.5);
        text.measureText = (string, style) => { text.families.add(style.FontFamily); return string.length * style.FontSize * width(style); };
        text.measureRange = (string, style, start, length) => Math.max(0, Math.min(length, string.length - start)) * style.FontSize * width(style);
        return text;
    };
    // Other measurers altogether: everything is wider.
    const wider = (text) => ({
        measureText: (string, style) => text.measureText(string, style) * 1.1,
        measureRange: (string, style, start, length) => text.measureRange(string, style, start, length) * 1.1,
    });
    const viewport = (scrollTop, scale = 1, width = 1000) => ({ scale, devicePixelRatio: 1, scrollTop, scrollLeft: 0, width, height: 800 });
    const make = (text, withViewport) => {
        const canvas = document.createElement('canvas'); canvas.width = 1000; canvas.height = 800;
        const editor = Word.initializeCanvas(canvas);
        if (withViewport) Word.setViewport(editor, viewport(0));
        Word.setTextMeasurers(editor, text.measureText, text.measureRange);
        Word.setParagraphs(editor, paragraphs);
        return editor;
    };
    const linesOf = (editor) => Word.getLayoutLines(editor).map((line) => [line.BlockId.length > 0, Math.round(line.Y), line.StartOffset, line.EndOffset]);

    // `change(editor, text, state)` is what the host does once the change of everything has returned
    // (state: the host's scroll position and scale); `reference(text)` gives the measurers a document
    // that was bold from the start is laid out with for the layout at the end.
    const scenario = async (name, hasBold, change, reference, sameScroll, atOnce = false) => {
        const text = stack();
        text.hasBold = hasBold;
        const editor = make(text, true);
        await Word.whenLayoutComplete(editor);
        const before = Word.getDocumentSize(editor);
        const plainLines = Word.getLayoutLines(editor);
        const firstLines = plainLines.filter((line, i) => line.BlockId.length > 0 && line.StartOffset === 0 && (i === 0 || plainLines[i - 1].BlockId !== line.BlockId));
        const topLine = firstLines[Math.floor(firstLines.length * 0.7)];
        // The host: a scroller as high as the document says it is (a scroll position beyond it is
        // not one), that scrolls where a notice says.
        // (The top of the view is in the block's first line: the view is held by that block.)
        const host = { scrollTop: topLine.Y + 3, scale: 1, width: 1000 };
        const blockAtTop = () => { const hit = Word.hitTestViewport(editor, 300, 4 * host.scale); return hit ? hit.BlockId : null; };
        Word.setViewport(editor, viewport(host.scrollTop));
        const notices = [];
        const scrolls = [];
        // (How far the document reaches below the bottom of the view, at the least.)
        let lowest = Infinity, clamped = 0;
        const below = () => Word.getDocumentSize(editor).ScaledHeight - (Word.getViewport(editor).scrollTop + 800);
        const off = Word.onChange(editor, (notice) => {
            const size = Word.getDocumentSize(editor);
            lowest = Math.min(lowest, below());
            notices.push([Math.round(size.Height), size.PageCount, Word.getRenderStats(editor).StaleBlocks, notice.ScrollTop]);
            if (typeof notice.ScrollTop === 'number') { scrolls.push(notice.ScrollTop); host.scrollTop = notice.ScrollTop; Word.setViewport(editor, viewport(host.scrollTop, host.scale, host.width)); }
            const most = Math.max(0, size.ScaledHeight - 800);
            if (host.scrollTop > most) { clamped++; host.scrollTop = most; Word.setViewport(editor, viewport(host.scrollTop, host.scale, host.width)); }
        });
        Word.selectAll(editor);
        Word.toggleBold(editor);
        const stale = Word.getRenderStats(editor).StaleBlocks;
        const then = Word.getDocumentSize(editor);
        check(`${name}: everything made bold far down the document - the blocks above the view are kept as they were`, [stale > 1000, Word.isLayoutComplete(editor), blockAtTop()], [true, false, topLine.BlockId]);
        // A slice of the layout from the start, then the host's change.
        if (!atOnce) {
            await new Promise((resolve) => setTimeout(resolve, 0));
            await new Promise((resolve) => setTimeout(resolve, 0));
            check(`${name}: (the layout from the start is under way)`, [Word.isLayoutComplete(editor), Word.getRenderStats(editor).StaleBlocks > 1000], [false, true]);
        }
        const t = performance.now();
        const result = change(editor, text, host);
        const took = performance.now() - t;
        const stats = Word.getRenderStats(editor);
        const now = Word.getDocumentSize(editor);
        note(`${name}: returned ${JSON.stringify(result)} in ${took.toFixed(1)} ms; ${stats.StaleBlocks} blocks stale, height ${Math.round(now.Height)} and ${now.PageCount} pages (${Math.round(then.Height)} and ${then.PageCount} before it, ${Math.round(before.Height)} and ${before.PageCount} before the bold)`);
        check(`${name}: the blocks above the view stay as they were laid out, the document is as high and has the pages it had`, [stats.StaleBlocks > 1000, now.Height >= then.Height * 0.98, now.PageCount >= then.PageCount - 1, below() >= 0, Word.isLayoutComplete(editor)], [true, true, true, true, false]);
        check(`${name}: the view is on the block it was on`, [blockAtTop(), sameScroll ? Word.getViewport(editor).scrollTop === topLine.Y + 3 : true], [topLine.BlockId, true]);
        await new Promise((resolve) => setTimeout(resolve, 0));
        await Word.whenLayoutComplete(editor);
        await new Promise((resolve) => setTimeout(resolve, 0));
        off();
        const end = Word.getDocumentSize(editor);
        check(`${name}: never did the document end above the bottom of the view, the host's scroll position was never beyond its end`, [lowest >= 0, clamped], [true, 0]);
        check(`${name}: at the end the host was told where to scroll, and the view is on the block it was on`, [scrolls.length >= 1, Word.getViewport(editor).scrollTop === host.scrollTop, blockAtTop(), Word.getRenderStats(editor).StaleBlocks, Word.isLayoutComplete(editor)], [true, true, topLine.BlockId, 0, true]);
        // The same document, bold from the start, with the measurers the host has now.
        const other = stack();
        const whole = make(other, false);
        Word.selectAll(whole);
        Word.toggleBold(whole);
        const measurers = reference(other);
        Word.setTextMeasurers(whole, measurers.measureText, measurers.measureRange);
        check(`${name}: the layout at the end is that of the document laid out at once with what the host has now`, [end.PageCount, JSON.stringify(linesOf(editor)) === JSON.stringify(linesOf(whole))], [Word.getPageCount(whole), true]);
        if (counts(notices)) note(`${name}: notices [height, pages, stale blocks, ScrollTop]: ${JSON.stringify(notices.slice(0, 6))}${notices.length > 6 ? ' ...' : ''}`);
        Word.dispose(editor); Word.dispose(whole);
    };
    const counts = (notices) => notices.length > 0;
    const family = (text) => [...text.families].filter((name) => typeof name === 'string' && name.length > 0);

    // The host got the bold face: measurers that know it, and which family it is of.
    await scenario('a face arrives', false,
        (editor, text) => { text.hasBold = true; const families = family(text); return Word.setTextMeasurersForFaces(editor, text.measureText, text.measureRange, (families.length ? families : ['Calibri', 'Arial', 'Times New Roman']).map((name) => ({ family: name, bold: true, italic: false }))) > 0; },
        (text) => { text.hasBold = true; return text; }, true);
    // The same before a slice of the layout from the start has run (the face is at hand when the
    // listeners are told of the bold): the document was the six blocks that layout starts with.
    await scenario('a face arrives at once', false,
        (editor, text) => { text.hasBold = true; const families = family(text); return Word.setTextMeasurersForFaces(editor, text.measureText, text.measureRange, (families.length ? families : ['Calibri', 'Arial', 'Times New Roman']).map((name) => ({ family: name, bold: true, italic: false }))) > 0; },
        (text) => { text.hasBold = true; return text; }, true, true);
    // The same without saying which (everything is measured again).
    await scenario('a face arrives, not named', false,
        (editor, text) => { text.hasBold = true; return Word.setTextMeasurersForFaces(editor, text.measureText, text.measureRange, null); },
        (text) => { text.hasBold = true; return text; }, true);
    // Other measurers.
    await scenario('other measurers', true,
        (editor, text) => { const other = wider(text); Word.setTextMeasurers(editor, other.measureText, other.measureRange); return true; },
        (text) => { text.hasBold = true; return wider(text); }, true);
    // The host zooms: its content is twice as high, its scroll position twice as far down.
    await scenario('zoom', true,
        (editor, text, host) => { host.scale = 2; host.scrollTop = Word.getViewport(editor).scrollTop * 2; Word.setViewport(editor, { scale: 2, devicePixelRatio: 1, scrollTop: host.scrollTop, scrollLeft: 0, width: 1000, height: 800 }); return true; },
        (text) => { text.hasBold = true; return text; }, false);
    // The host's window gets narrower.
    await scenario('another width', true,
        (editor, text, host) => { host.width = 700; Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: host.scrollTop, scrollLeft: 0, width: 700, height: 800 }); return true; },
        (text) => { text.hasBold = true; return text; }, true);
} catch (e) { fail(e); }
finish();
