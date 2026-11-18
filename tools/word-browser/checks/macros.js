// Word macros in the browser: each showcase .docm (tests/macros/word-showcase/docm) opened, its macros
// known without the VBA engine (hasMacros), enabled (enableMacros loads WordMacroProject.js on use), Main
// run as one undo step (undo / redo), the Immediate window (runMacroLine), a module edited
// (setMacroModule) and saved with the document: read again, the project has the edited code. The pages
// each macro made are painted and given back as PNGs (window.result.images; run.mjs writes them to
// $SHOT_DIR) for the showcase.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';

const names = ['rainbow-text', 'create-new-table', 'find-and-replace', 'highlight-word-list', 'month-calendar',
    'multiplication-table', 'contents-list', 'sales-report', 'form-letters'];
const faces = [['Calibri', 'calibri.ttf', 'normal', 'normal'], ['Calibri', 'calibrib.ttf', 'bold', 'normal'],
    ['Calibri', 'calibrii.ttf', 'normal', 'italic'], ['Calibri', 'calibriz.ttf', 'bold', 'italic'],
    ['Calibri Light', 'calibril.ttf', 'normal', 'normal'], ['Calibri Light', 'calibrili.ttf', 'normal', 'italic'],
    ['Arial', 'arial.ttf', 'normal', 'normal']];

/// The pages side by side on a grey ground, as one PNG.
const picture = (Word, editor, maxPages, scale) => {
    const count = Math.min(Word.getPageCount(editor), maxPages);
    const pages = [];
    for (let i = 0; i < count; i++) {
        const c = document.createElement('canvas');
        Word.renderPage(editor, c, i, scale);
        pages.push(c);
    }
    const gap = 24;
    const out = document.createElement('canvas');
    out.width = pages.reduce((w, p) => w + p.width + gap, gap);
    out.height = Math.max(...pages.map((p) => p.height)) + 2 * gap;
    const g = out.getContext('2d');
    g.fillStyle = '#d9dde3';
    g.fillRect(0, 0, out.width, out.height);
    let x = gap;
    for (const p of pages) {
        g.fillStyle = 'rgba(0,0,0,0.18)';
        g.fillRect(x + 3, gap + 3, p.width, p.height);
        g.drawImage(p, x, gap);
        x += p.width + gap;
    }
    return out.toDataURL('image/png');
};

try {
    for (const [family, file, weight, style] of faces) {
        try {
            const face = new FontFace(family, `url(/font?name=${file})`, { weight, style });
            await face.load();
            document.fonts.add(face);
        } catch { note(`font ${file} not loaded`); }
    }
    const Word = await loadWord();
    const OF = await loadFormats();
    const engineLoaded = () => performance.getEntriesByType('resource').some((r) => r.name.includes('WordMacroProject.js'));
    window.images = {};
    let first = true;
    for (const name of names) {
        const bytes = await bytesOf(`/repo/tests/macros/word-showcase/docm/${name}.docm`);
        const canvas = document.createElement('canvas');
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        check(`${name}: hasMacros without the engine`, Word.hasMacros(editor), true);
        if (first) check('the VBA engine is not loaded before macros are used', engineLoaded(), false);
        const output = [];
        Word.setMacroHooks(editor, { output: (text) => output.push(text), messageBox: () => 1 });
        let threw = false;
        try { await Word.runMacro(editor, 'Main'); } catch { threw = true; }
        check(`${name}: runMacro before enableMacros throws`, threw, true);
        check(`${name}: enableMacros`, (await Word.enableMacros(editor)).ok, true);
        if (first) check('enableMacros loaded the engine (WordMacroProject.js)', engineLoaded(), true);
        check(`${name}: Main is a macro`, Word.getMacros(editor).includes('Main'), true);
        check(`${name}: no compile problems`, Word.getMacroProblems(editor), []);
        const before = Word.getText(editor);
        const started = performance.now();
        const result = await Word.runMacro(editor, 'Main');
        const ms = Math.round(performance.now() - started);
        check(`${name}: Main runs`, result.ok ? true : `${result.number} ${result.description} (${result.procedure}, line ${result.line})`, true);
        const after = Word.getText(editor);
        check(`${name}: the document changed`, after !== before, true);
        check(`${name}: modified`, Word.isModified(editor), true);
        Word.undo(editor);
        check(`${name}: one undo step back`, Word.getText(editor), before);
        Word.redo(editor);
        check(`${name}: redo`, Word.getText(editor), after);
        Word.completeLayout(editor);
        note(`${name}: ran in ${ms} ms, ${Word.getPageCount(editor)} page(s)`);
        window.images[name] = picture(Word, editor, name === 'month-calendar' || name === 'form-letters' ? 3 : 1, 1);
        if (first) {
            // The Immediate window and an edited module, saved with the document and read again.
            await Word.runMacroLine(editor, '?ActiveDocument.Paragraphs.Count');
            check('runMacroLine prints the value', Word.getMacroOutput(editor).some((line) => /^ ?\d+ ?$/.test(line)), true);
            const problems = await Word.setMacroModule(editor, 'Module2', 'Sub Shout()\n    ActiveDocument.Content.Font.AllCaps = True\nEnd Sub');
            check('setMacroModule: compiles', problems, []);
            check('setMacroModule: the new macro is listed', Word.getMacros(editor).includes('Shout'), true);
            check('runMacro of the edited module', (await Word.runMacro(editor, 'Shout')).ok, true);
            const bad = await Word.setMacroModule(editor, 'Module3', 'Sub Broken(\nEnd Sub');
            check('setMacroModule: a problem is reported', bad.length > 0 && bad[0].module === 'Module3', true);
            await Word.setMacroModule(editor, 'Module3', 'Sub Fixed()\nEnd Sub');
            const saved = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
            const again = Word.initializeCanvas(document.createElement('canvas'));
            Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(saved));
            check('saved: the document keeps its macros', Word.hasMacros(again), true);
            await Word.loadMacros(again);
            const modules = Word.getMacroModules(again).map((m) => m.name);
            check('saved: the edited project', ['Module2', 'Module3'].every((m) => modules.includes(m)) && modules.includes('RainbowText'), true);
            check('saved: the edited code', Word.getMacroModules(again).find((m) => m.name === 'Module2').code.includes('AllCaps = True'), true);
            check('a newly opened document: macros off', Word.getMacrosEnabled(again), false);
            first = false;
        }
    }
} catch (e) { fail(e); }
finish();
window.result.images = window.images;
