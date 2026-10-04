// Caret movement and selection with real keyboard and mouse events (driven by
// run.mjs through Playwright), on the layout that is on screen: Up / Down move
// by visual lines and keep the column, Home / End go to the ends of the
// visual line, Ctrl+Left / Right by words, PageUp / PageDown by a screenful,
// Ctrl+Home / Ctrl+End to the ends of the document, Shift extends; a double
// click selects a word, a triple click a paragraph.
import { check, note, finish, fail, loadWord } from './common.js';
try {
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initialize('c');
    const long = 'The quick brown fox jumps over the lazy dog and keeps running through the forest until the evening comes, when it rests beside the river and watches the water go by for a while before it sets off again towards the hills in the north.';
    const text = long + ' ' + long.replace('quick', 'swift') + ' ' + long.replace('quick', 'clever');
    const paragraphs = [text, 'Short one.', text.replace('quick', 'slow'), ...Array.from({ length: 60 }, (_, i) => `Filler paragraph number ${i + 1} with a few words in it.`)];
    Word.setParagraphs(editor, paragraphs);
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 700 });
    const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
    const linesOf = (id) => Word.getLayoutLines(editor).filter((line) => line.BlockId === id);
    const first = linesOf(ids[0]);
    note(`the first paragraph wraps into ${first.length} lines`);
    check('the first paragraph wraps into several lines', first.length >= 6, true);
    const snaps = {};
    window.snap = (name) => {
        const selection = Word.getSelection(editor), caret = Word.getCaretRect(editor);
        snaps[name] = { anchor: [ids.indexOf(selection.AnchorBlockId), selection.AnchorOffset], focus: [ids.indexOf(selection.FocusBlockId), selection.FocusOffset], x: caret.X, y: caret.Y };
    };
    const box = canvas.getBoundingClientRect();
    // A point 200 px into the second visual line of the first paragraph, and a word in it.
    const second = first[1], third = first[2];
    const clickX = box.left + second.X + 200, clickY = box.top + second.Y + second.Height / 2;
    window.finish = () => {
        try {
            const s = snaps;
            note('snaps ' + JSON.stringify(Object.fromEntries(Object.entries(s).map(([name, v]) => [name, [v.focus.join(':'), Math.round(v.x), Math.round(v.y)]]))));
            note('third paragraph ' + JSON.stringify(Word.getText(editor).split('\n')[2].slice(0, 30)));
            const sameLine = (a, b) => Math.abs(a.y - b.y) < 1;
            check('a click puts the caret in the second line', [s.click.focus[0], sameLine(s.click, { y: second.Y }) || Math.abs(s.click.y - second.Y) < second.Height], [0, true]);
            // Home / End: the ends of the visual line, not of the paragraph.
            check('End stays on the visual line', [sameLine(s.end, s.click), s.end.focus[1] > s.click.focus[1], s.end.focus[1] < text.length, s.end.x > s.click.x + 100], [true, true, true, true]);
            check('Home goes to the start of the visual line', [sameLine(s.home, s.click), s.home.focus[1] > 0, s.home.focus[1] < s.click.focus[1], Math.abs(s.home.x - second.X) < 2], [true, true, true, true]);
            check('the line runs from Home to End', text.slice(s.home.focus[1], s.end.focus[1]).length > 40, true);
            // Up / Down keep the column and move one visual line within the paragraph.
            check('Down moves one visual line down, in the same paragraph', [s.down.focus[0], Math.round(s.down.y - s.back.y) === Math.round(third.Y - second.Y), Math.abs(s.down.x - s.back.x) < 12], [0, true, true]);
            check('Up comes back to the same place', [s.up.focus, Math.abs(s.up.x - s.back.x) < 12], [s.back.focus, true]);
            check('Up again: the first line, same column', [s.up2.focus[0], s.up2.y < s.up.y, Math.abs(s.up2.x - s.back.x) < 12, s.up2.focus[1] > 0], [0, true, true, true]);
            // The column is kept across a short line.
            check('Down through a short paragraph and on keeps the column', [s.shortLine.focus[0], s.shortLine.focus[1], s.pastShort.focus[0], Math.abs(s.pastShort.x - s.column.x) < 12], [1, paragraphs[1].length, 2, true]);
            // Ctrl+Left / Right: words.
            const wordStarts = [...text.matchAll(/\S+/g)].map((m) => m.index);
            check('Ctrl+Right goes to the start of the next word', [wordStarts.includes(s.wordRight.focus[1]), s.wordRight.focus[1] > s.wordFrom.focus[1], s.wordRight2.focus[1] > s.wordRight.focus[1], wordStarts.includes(s.wordRight2.focus[1])], [true, true, true, true]);
            check('Ctrl+Left goes back to the starts of words', [s.wordLeft.focus[1], wordStarts.includes(s.wordLeft2.focus[1]), s.wordLeft2.focus[1] < s.wordLeft.focus[1]], [s.wordRight.focus[1], true, true]);
            // Shift extends, the anchor stays.
            check('Shift+Down extends the selection by a visual line', [s.shiftDown.anchor, s.shiftDown.focus[0], s.shiftDown.y > s.shiftFrom.y, Math.abs(s.shiftDown.x - s.shiftFrom.x) < 12], [s.shiftFrom.focus, 0, true, true]);
            check('Shift+Up takes it back', [s.shiftUp.anchor, s.shiftUp.focus], [s.shiftFrom.focus, s.shiftFrom.focus]);
            check('Shift+End selects to the end of the visual line', [s.shiftEnd.anchor, s.shiftEnd.focus[1] > s.shiftFrom.focus[1], sameLine(s.shiftEnd, s.shiftFrom)], [s.shiftFrom.focus, true, true]);
            check('Shift+Ctrl+Right selects a word further', [s.shiftWord.anchor, wordStarts.includes(s.shiftWord.focus[1]) || s.shiftWord.focus[1] === text.length], [s.shiftFrom.focus, true]);
            // Document ends and pages.
            check('Ctrl+End goes to the end of the document', s.docEnd.focus, [ids.length - 1, paragraphs.at(-1).length]);
            check('Ctrl+Home goes to its start', s.docHome.focus, [0, 0]);
            check('Shift+Ctrl+End selects to the end', [s.selectToEnd.anchor, s.selectToEnd.focus], [[0, 0], [ids.length - 1, paragraphs.at(-1).length]]);
            check('PageDown moves about a screenful down', [s.pageDown.y - s.docHome2.y > 400, s.pageDown.y - s.docHome2.y < 760, s.pageDown.focus[0] > 3], [true, true, true]);
            check('PageDown again moves further', s.pageDown2.y > s.pageDown.y + 300, true);
            check('PageUp twice comes back to the top', [s.pageUp2.focus[0], s.pageUp2.y <= s.docHome2.y + 1], [0, true]);
            // Double and triple click.
            const word = text.slice(Math.min(...[s.double.anchor[1], s.double.focus[1]]), Math.max(s.double.anchor[1], s.double.focus[1]));
            check('a double click selects the word under the pointer', [/^\w+$/.test(word), word.length > 1, s.double.focus[0]], [true, true, 0]);
            check('a triple click selects the paragraph', [[s.triple.anchor, s.triple.focus].sort((a, b) => a[1] - b[1]), s.triple.focus[0]], [[[0, 0], [0, text.length]], 0]);
            check('typing replaces what a double click selected', Word.getText(editor).split('\n')[2].startsWith('The Z '), true);
            check('Shift+click extends the selection from the caret', [s.shiftClick.anchor, s.shiftClick.focus[0], s.shiftClick.focus[1] > 0], [s.beforeShiftClick.focus, 0, true]);
        } catch (e) { fail(e); }
        finish();
    };
    const thirdPara = linesOf(ids[2])[0];
    window.drive = [
        ['click', clickX, clickY], ['eval', 'snap("click")'],
        ['press', 'End'], ['eval', 'snap("end")'],
        ['press', 'Home'], ['eval', 'snap("home")'],
        ['click', clickX, clickY], ['eval', 'snap("back")'],
        ['press', 'ArrowDown'], ['eval', 'snap("down")'],
        ['press', 'ArrowUp'], ['eval', 'snap("up")'],
        ['press', 'ArrowUp'], ['eval', 'snap("up2")'],
        // From the last line of the first paragraph, far right: down through "Short one." into the third paragraph.
        ['click', box.left + first.at(-2).X + 420, box.top + first.at(-2).Y + first.at(-2).Height / 2], ['eval', 'snap("column")'],
        ['press', 'ArrowDown'],
        ['press', 'ArrowDown'], ['eval', 'snap("shortLine")'],
        ['press', 'ArrowDown'], ['eval', 'snap("pastShort")'],
        ['click', clickX, clickY], ['eval', 'snap("wordFrom")'],
        ['press', 'Control+ArrowRight'], ['eval', 'snap("wordRight")'],
        ['press', 'Control+ArrowRight'], ['eval', 'snap("wordRight2")'],
        ['press', 'Control+ArrowLeft'], ['eval', 'snap("wordLeft")'],
        ['press', 'Control+ArrowLeft'], ['eval', 'snap("wordLeft2")'],
        ['click', clickX, clickY], ['eval', 'snap("shiftFrom")'],
        ['press', 'Shift+ArrowDown'], ['eval', 'snap("shiftDown")'],
        ['press', 'Shift+ArrowUp'], ['eval', 'snap("shiftUp")'],
        ['press', 'Shift+End'], ['eval', 'snap("shiftEnd")'],
        ['click', clickX, clickY],
        ['press', 'Control+Shift+ArrowRight'], ['eval', 'snap("shiftWord")'],
        ['press', 'Control+End'], ['eval', 'snap("docEnd")'],
        ['press', 'Control+Home'], ['eval', 'snap("docHome")'],
        ['press', 'Control+Shift+End'], ['eval', 'snap("selectToEnd")'],
        ['press', 'Control+Home'], ['eval', 'snap("docHome2")'],
        ['press', 'PageDown'], ['eval', 'snap("pageDown")'],
        ['press', 'PageDown'], ['eval', 'snap("pageDown2")'],
        ['press', 'PageUp'], ['press', 'PageUp'], ['eval', 'snap("pageUp2")'],
        ['dblclick', clickX, clickY], ['eval', 'snap("double")'],
        ['tripleclick', clickX, clickY], ['eval', 'snap("triple")'],
        // Double click the second word of the third paragraph ("slow") and type over it.
        ['dblclick', box.left + thirdPara.X + 30, box.top + thirdPara.Y + thirdPara.Height / 2], ['type', 'Z'],
        ['click', clickX, clickY], ['eval', 'snap("beforeShiftClick")'],
        ['shiftclick', clickX + 150, clickY], ['eval', 'snap("shiftClick")'],
    ];
} catch (e) { fail(e); finish(); }
