// Item: IME / composition input through the hidden text-input proxy (synthetic events).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
  const Word = await loadWord();
  const canvas = document.getElementById('c');
  const editor = Word.initialize('c');
  const proxy = Word.getInputElement(editor);
  check('proxy is a textarea next to the canvas', [proxy && proxy.tagName, proxy && proxy.previousElementSibling === canvas, document.querySelectorAll('textarea[data-rbt-word-input]').length], ['TEXTAREA', true, 1]);
  // Focusing the canvas hands the focus to the proxy; ordinary keys still work there.
  canvas.focus();
  check('focus goes to the proxy', document.activeElement === proxy, true);
  const key = (target, k, extra = {}) => target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...extra }));
  for (const ch of 'ab') key(proxy, ch);
  key(canvas, 'c');   // hosts and tests that send keys to the canvas keep working
  check('typing through keydown', Word.getText(editor), 'abc');
  const undoStepsToEmpty = () => { let n = 0; while (Word.canUndo(editor)) { Word.undo(editor); n++; } return n; };
  // --- CJK composition: "ni" -> "你" then "你hao" -> "你好"
  const comp = (type, data) => proxy.dispatchEvent(new CompositionEvent(type, { data, bubbles: true, cancelable: true }));
  const ctx = canvas.getContext('2d');
  const inkBelow = (rect) => { // count dark pixels in the caret line box
    const d = ctx.getImageData(Math.max(0, Math.round(rect.x - 60)), Math.round(rect.y), 120, Math.round(rect.h + 4)).data; let n = 0;
    for (let i = 0; i < d.length; i += 4) if (d[i] < 120 && d[i + 1] < 120 && d[i + 2] < 120) n++; return n; };
  const proxyRect = () => { const r = proxy.getBoundingClientRect(), c = canvas.getBoundingClientRect(); return { x: r.left - c.left, y: r.top - c.top, h: r.height }; };
  const before = proxyRect();
  const inkBefore = inkBelow(before);
  comp('compositionstart', '');
  comp('compositionupdate', 'n');
  key(proxy, 'Process', { isComposing: true, keyCode: 229 });
  comp('compositionupdate', 'ni');
  check('document text is untouched while composing', Word.getText(editor), 'abc');
  const during = proxyRect();
  check('composition text is drawn at the caret', inkBelow(during) > inkBefore, true);
  check('proxy follows the composition caret', during.x > before.x, true);
  key(proxy, 'Enter', { isComposing: true, keyCode: 229 });   // confirming a candidate must not split the paragraph
  key(proxy, 'Backspace', { isComposing: true, keyCode: 229 });
  check('keys during composition are left to the IME', Word.getText(editor), 'abc');
  comp('compositionupdate', '你好');
  comp('compositionend', '你好');
  proxy.dispatchEvent(new InputEvent('input', { data: '你好', inputType: 'insertCompositionText', bubbles: true }));
  check('composition committed', Word.getText(editor), 'abc你好');
  check('proxy is emptied', proxy.value, '');
  check('undo once removes the whole composition', [Word.undo(editor), Word.getText(editor)], [true, 'abc']);
  check('redo brings it back', [Word.redo(editor), Word.getText(editor)], [true, 'abc你好']);
  // --- dead key: ´ + e -> é
  key(proxy, 'Dead');
  comp('compositionstart', '');
  comp('compositionupdate', '´');
  key(proxy, 'e', { isComposing: true, keyCode: 229 });
  comp('compositionend', 'é');
  check('dead key composition', Word.getText(editor), 'abc你好é');
  // --- a cancelled composition leaves the document alone and adds no undo step
  comp('compositionstart', '');
  comp('compositionupdate', 'x');
  comp('compositionend', '');
  check('cancelled composition changes nothing', Word.getText(editor), 'abc你好é');
  // --- composition replaces the selection, still one undo step
  Word.selectAll(editor);
  comp('compositionstart', '');
  comp('compositionupdate', 'か');
  comp('compositionend', '漢');
  check('composition replaces the selection', Word.getText(editor), '漢');
  check('... as one undo step', [Word.undo(editor), Word.getText(editor)], [true, 'abc你好é']);
  // --- text from a virtual keyboard / emoji picker: beforeinput without a keydown
  key(proxy, 'End');   // undo restored the select-all selection; put the caret at the end
  const bi = new InputEvent('beforeinput', { data: '😀', inputType: 'insertText', bubbles: true, cancelable: true });
  proxy.dispatchEvent(bi);
  check('beforeinput insertText', [Word.getText(editor), bi.defaultPrevented], ['abc你好é😀', true]);
  proxy.dispatchEvent(new InputEvent('beforeinput', { inputType: 'deleteContentBackward', bubbles: true, cancelable: true }));
  check('beforeinput deleteContentBackward', Word.getText(editor), 'abc你好é');
  proxy.value = 'pasted';
  proxy.dispatchEvent(new InputEvent('input', { inputType: 'insertFromPaste', bubbles: true }));
  check('input fallback takes the proxy text', [Word.getText(editor), proxy.value], ['abc你好épasted', '']);
  // undo steps: abc (keys typed one after the other are one) | 你好 | é | 😀 | delete | pasted  = 6
  check('each input is one undo step', undoStepsToEmpty(), 6);
  check('all undone', Word.getText(editor), '');
  // --- ordinary keys after all that
  for (const ch of 'ok') key(proxy, ch);
  key(proxy, 'Backspace');
  check('ordinary keydown handling still works', Word.getText(editor), 'o');
  // --- dispose removes the proxy
  Word.dispose(editor);
  check('dispose removes the proxy', [document.querySelectorAll('textarea[data-rbt-word-input]').length, Word.getInputElement(editor)], [0, null]);
  canvas.focus();
  check('focus stays on the canvas after dispose', document.activeElement === canvas, true);
} catch (e) { fail(e); }
finish();
