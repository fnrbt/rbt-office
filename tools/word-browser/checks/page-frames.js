// Item: on-screen page frames, page count and hit-testing follow each section's page setup.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
  const OF = await loadFormats();
  const Word = await loadWord();
  const bytes = await bytesOf('/repo/tests/repros/word/landscape-section.docx');
  const pkg = await OF.processors.forEditorPackage(bytes);
  const canvas = document.getElementById('c');
  const editor = Word.initialize('c');
  Word.importFromDocxPackage(editor, pkg);
  const frames = Word.getPageFrames(editor).map((f) => ({ top: Math.round(f.Top), w: Math.round(f.Width), h: Math.round(f.Height) }));
  note('frames ' + JSON.stringify(frames));
  check('page count matches frames', Word.getPageCount(editor), frames.length);
  const portrait = frames.find((f) => f.h > f.w), landscape = frames.find((f) => f.w > f.h);
  check('has a portrait and a landscape frame', [!!portrait, !!landscape], [true, true]);
  for (let i = 1; i < frames.length; i++) check('frame ' + i + ' stacks below frame ' + (i - 1), frames[i].top, frames[i - 1].top + frames[i - 1].h + 28);
  const ctx = canvas.getContext('2d');
  const px = (x, y) => Array.from(ctx.getImageData(x, y, 1, 1).data.slice(0, 3));
  const white = [255, 255, 255], desk = [238, 240, 243];
  for (const [i, f] of frames.entries()) {
    check('frame ' + i + ' is paper just inside its right edge', px(f.w - 6, f.top + f.h - 6), white);
    check('frame ' + i + ' is desk just outside its right edge', px(f.w + 6, f.top + f.h - 6), desk);
    check('frame ' + i + ' is paper just above its bottom edge', px(8, f.top + f.h - 4), white);
    check('gap below frame ' + i + ' is desk', px(8, f.top + f.h + 8), desk);
  }
  // Hit-testing uses the same geometry: a click on a line of each page lands on that line's paragraph.
  const lines = Word.getLayoutLines(editor);
  const sample = (f) => lines.find((l) => l.Y >= f.top && l.Y < f.top + f.h && (l.Text || '').trim().length > 3);
  for (const [i, f] of frames.entries()) {
    const line = sample(f);
    if (!line) { note('frame ' + i + ' has no text line'); continue; }
    const hit = Word.hitTestPosition(editor, line.X + 4, line.Y + line.Height / 2);
    check('hit test on frame ' + i + ' finds its paragraph', hit && hit.BlockId, line.BlockId);
  }
} catch (e) { fail(e); }
finish();
