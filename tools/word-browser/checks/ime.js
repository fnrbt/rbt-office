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
  // Virtual paragraph intent and physical Enter share list/table/tracked semantics.
  const paragraph = () => { const event = new InputEvent('beforeinput', { inputType: 'insertParagraph', bubbles: true, cancelable: true }); proxy.dispatchEvent(event); return event; };
  Word.setParagraphs(editor, ['ab', 'tail']);
  let first = Word.getLayoutLines(editor)[0].BlockId;
  Word.setSelection(editor, first, 1, first, 1);
  check('virtual paragraph cancels native textarea insertion', paragraph().defaultPrevented, true);
  check('virtual paragraph splits the document', [Word.getText(editor), Word.undo(editor), Word.getText(editor)], ['a\nb\ntail', true, 'ab\ntail']);
  await new Promise(resolve => setTimeout(resolve, 0));
  Word.setSelection(editor, first, 1, first, 1);
  proxy.value = '\n';
  proxy.dispatchEvent(new InputEvent('input', { inputType: 'insertParagraph', bubbles: true }));
  check('input-only virtual paragraph uses paragraph edit, not newline text', [Word.getText(editor), Word.undo(editor), Word.getText(editor)], ['a\nb\ntail', true, 'ab\ntail']);
  Word.setSelection(editor, first, 1, first, 1);
  key(proxy, 'Enter');
  paragraph(); // browser should suppress this after a prevented keydown
  proxy.dispatchEvent(new InputEvent('input', { inputType: 'insertParagraph', bubbles: true }));
  check('physical Enter plus stray beforeinput/input inserts once', [Word.getText(editor), Word.undo(editor), Word.getText(editor)], ['a\nb\ntail', true, 'ab\ntail']);
  Word.setParagraphs(editor, ['ab']);
  first = Word.getLayoutLines(editor)[0].BlockId;
  Word.setSelection(editor, first, 1, first, 1);
  paragraph();
  paragraph();
  check('two virtual Enter intents create two paragraphs', [Word.getText(editor), Word.undo(editor), Word.getText(editor)], ['a\n\nb', true, 'a\nb']);
  Word.setParagraphs(editor, ['ab']);
  first = Word.getLayoutLines(editor)[0].BlockId;
  Word.setSelection(editor, first, 1, first, 1);
  key(proxy, 'Enter', { shiftKey: true });
  check('Shift+Enter remains a line break in one paragraph', [Word.getText(editor), Word.exportDocumentPackage(editor).Paragraphs.length, Word.undo(editor), Word.getText(editor)], ['a\nb', 1, true, 'ab']);
  Word.setSelection(editor, first, 1, first, 1);
  proxy.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertLineBreak', bubbles: true, cancelable: true }));
  check('virtual insertLineBreak remains in one paragraph', [Word.getText(editor), Word.exportDocumentPackage(editor).Paragraphs.length, Word.undo(editor)], ['a\nb', 1, true]);
  Word.setParagraphs(editor, ['x']);
  first = Word.getLayoutLines(editor)[0].BlockId;
  Word.setSelection(editor, first, 0, first, 1);
  Word.toggleBulletList(editor);
  key(proxy, 'Backspace');
  const listBefore = Word.getSelectionFormat(editor).Paragraph.ListKind;
  paragraph();
  check('virtual Enter exits an empty list without a new paragraph', [listBefore, Word.getText(editor), Word.getSelectionFormat(editor).Paragraph.ListKind === 'bullet'], ['bullet', '', false]);
  check('list exit has one undo step', [Word.undo(editor), Word.getSelectionFormat(editor).Paragraph.ListKind], [true, 'bullet']);
  Word.setParagraphs(editor, ['body']);
  Word.insertTable(editor, 1, 1);
  Word.insertText(editor, 'ab');
  const cell = Word.getSelection(editor).FocusBlockId;
  Word.setSelection(editor, cell, 1, cell, 1);
  const cellParagraphs = () => Word.exportDocumentPackage(editor).Paragraphs.find(p => p.Type === 'Table').Rows[0].Cells[0].Paragraphs.length;
  paragraph();
  check('virtual Enter splits inside the table cell', [cellParagraphs(), Word.getCaretContext(editor).InTable], [2, true]);
  check('table split undoes once', [Word.undo(editor), cellParagraphs()], [true, 1]);
  Word.setParagraphs(editor, ['ab']);
  first = Word.getLayoutLines(editor)[0].BlockId;
  Word.setTrackChangesOptions(editor, 'Input Test', '2026-11-11T00:00:00Z');
  Word.setTrackChanges(editor, true);
  Word.setSelection(editor, first, 1, first, 1);
  paragraph();
  check('virtual tracked paragraph splits and undoes', [Word.getText(editor), Word.undo(editor), Word.getText(editor)], ['a\nb', true, 'ab']);
  Word.setSelection(editor, first, 1, first, 1);
  paragraph();
  Word.rejectAllRevisions(editor);
  check('virtual Enter records a rejectable tracked paragraph mark', Word.getText(editor), 'ab');
  Word.setTrackChanges(editor, false);
  // Asynchronous canvas Cut must wait for the actual write and its captured selection.
  Word.setParagraphs(editor, ['cut here']);
  first = Word.getLayoutLines(editor)[0].BlockId;
  const selectCut = () => Word.setSelection(editor, first, 0, first, 3);
  const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const originalError = console.error;
  const failures = [];
  console.error = (...parts) => failures.push(parts[0]);
  const clipboard = value => Object.defineProperty(navigator, 'clipboard', { configurable: true, value });
  const settle = () => new Promise(resolve => setTimeout(resolve, 0));
  try {
    clipboard({ writeText: () => Promise.reject(new Error('denied')) });
    selectCut();
    key(canvas, 'x', { ctrlKey: true });
    await settle();
    check('denied canvas Cut retains text and reports failure', [Word.getText(editor), failures.some(m => String(m).includes('write failed'))], ['cut here', true]);
    clipboard(undefined);
    selectCut();
    key(canvas, 'x', { ctrlKey: true });
    check('unavailable canvas Cut retains text and reports failure', [Word.getText(editor), failures.length >= 2], ['cut here', true]);
    let completeSelectionWrite;
    clipboard({ writeText: () => new Promise(resolve => { completeSelectionWrite = resolve; }) });
    selectCut();
    key(canvas, 'x', { ctrlKey: true });
    Word.setSelection(editor, first, 3, first, 3);
    completeSelectionWrite();
    await settle();
    check('moving selection while Cut is pending retains text', [Word.getText(editor), failures.some(m => String(m).includes('cut cancelled'))], ['cut here', true]);
    let completeWrite;
    clipboard({ writeText: () => new Promise(resolve => { completeWrite = resolve; }) });
    selectCut();
    key(canvas, 'x', { ctrlKey: true });
    check('pending canvas Cut does not delete before transfer', Word.getText(editor), 'cut here');
    Word.insertText(editor, 'edited');
    completeWrite();
    await settle();
    check('intervening edit invalidates pending Cut', [Word.getText(editor), failures.some(m => String(m).includes('cut cancelled'))], ['edited here', true]);
    Word.setParagraphs(editor, ['cut here']);
    first = Word.getLayoutLines(editor)[0].BlockId;
    selectCut();
    key(canvas, 'x', { ctrlKey: true });
    completeWrite();
    await settle();
    check('successful canvas Cut deletes after write and undoes', [Word.getText(editor), Word.undo(editor), Word.getText(editor)], [' here', true, 'cut here']);
  } finally {
    console.error = originalError;
    if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
    else delete navigator.clipboard;
  }
  // Disposal invalidates retained proxy callbacks and all canvas listeners.
  const oldText = Word.getText(editor);
  selectCut();
  key(proxy, 'c', { ctrlKey: true });
  check('copy key leaves temporary selection in the proxy', proxy.value, 'cut');
  const pixelsBeforeDispose = canvas.toDataURL();
  // --- dispose removes the proxy
  Word.dispose(editor);
  check('dispose removes the proxy', [document.querySelectorAll('textarea[data-rbt-word-input]').length, Word.getInputElement(editor)], [0, null]);
  canvas.focus();
  check('focus stays on the canvas after dispose', document.activeElement === canvas, true);
  key(proxy, 'Q');
  key(canvas, 'X');
  canvas.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 300, clientY: 120, button: 0 }));
  canvas.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: 320, clientY: 120, button: 0 }));
  window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: 300, clientY: 120, button: 0 }));
  canvas.focus();
  await settle();
  check('disposed editor ignores retained proxy/canvas events and timer', [Word.getText(editor), canvas.toDataURL(), proxy.value], [oldText, pixelsBeforeDispose, 'cut']);
  const replacement = Word.initialize('c');
  Word.setParagraphs(replacement, ['one']);
  const replacementId = Word.getLayoutLines(replacement)[0].BlockId;
  Word.setSelection(replacement, replacementId, 0, replacementId, 0);
  key(canvas, 'Z');
  check('reused canvas has exactly one keyboard handler', [Word.getText(replacement), Word.getText(editor)], ['Zone', oldText]);
  Word.dispose(replacement);
} catch (e) { fail(e); }
finish();
