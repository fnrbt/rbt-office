// A text box as its file says: no outline and no fill where the shape states
// none, and a line of an exact height lower than its text inside the box.
// resume-template.docx, the header's text box: 3808729 x 723900 EMU (399.9 x
// 76.0 px) at 1875955 / 847000 EMU (196.9 / 88.9 px) from the page's corner;
// no a:ln, no fill and no wps:style - Word draws neither outline nor fill;
// insets 0, a:noAutofit; "Lorem Ipsum" in 55 pt (73.3 px) in a line of
// exactly 1083 twips (72.2 px). The text's descent stays in the line: the
// baseline is a fifth of the size above the line's bottom, 88.9 + 72.2 - 14.7
// = 146.4 px down the page. (It stood 73.3 px below the line's top - below the
// line -, so the box cut the text off, and the box had a grey outline.)
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const editor = Word.initializeCanvas(canvas);
    const pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-features/resume-template.docx'));
    Word.importFromDocxPackage(editor, pkg);
    const shape = (pkg.HeaderParts ?? []).flatMap((part) => part.Paragraphs ?? []).find((item) => item.Type === 'TextBox');
    check('read: the box, without a line and without a fill', [shape !== undefined, shape && Math.round(shape.Width), shape && Math.round(shape.Height), shape && shape.NoLine, shape && shape.NoFill], [true, 400, 76, true, true]);
    const context = canvas.getContext('2d', { willReadFrequently: true });
    const left = 197, top = 89, width = 400, height = 76;
    // The rows around the box's top and bottom edge, and the columns around its left edge: no line there.
    const lineInk = (x, y, w, h) => { const data = context.getImageData(x, y, w, h).data; let count = 0; for (let i = 0; i < data.length; i += 4) if (data[i] + data[i + 1] + data[i + 2] < 660) count++; return count; };
    check('painted: no outline (the box\'s bottom edge, where no text is, and the right half of its top edge)', [lineInk(left + 2, top + height - 2, width - 4, 4), lineInk(left + width / 2, top - 2, width / 2 - 2, 4)], [0, 0]);
    // The text: the lowest row with ink of the first word (no letter of "Lorem" goes below the baseline).
    const data = context.getImageData(left, top - 20, 230, height + 60).data;
    let lowest = -1, highest = -1;
    for (let row = 0; row < height + 60; row++) { let ink = 0; for (let column = 0; column < 230; column++) { const i = (row * 230 + column) * 4; if (data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) ink++; } if (ink > 2) { if (highest < 0) highest = row; lowest = row; } }
    const baseline = top - 20 + lowest + 1, capTop = top - 20 + highest;
    note(`the first word: ink from ${capTop} to ${baseline} px down the page (the box: ${top} to ${top + height})`);
    check('the text stands on a baseline a fifth of its size above the line\'s bottom, inside the box', [Math.abs(baseline - 146.4) <= 3, capTop >= top, baseline <= top + height], [true, true, true]);
} catch (e) { fail(e); }
finish();
