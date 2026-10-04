// The viewer has an entry module of its own (word/WordViewer.js) that does not
// pull in the editor: a preview bundle is the document import, the layout, the
// paint commands and the renderer. Checked on the import graph of the Fable
// output, and by using the viewer from that module alone.
import { check, note, finish, fail, loadFormats, bytesOf } from './common.js';
try {
    // Every module WordViewer.js imports, directly or not.
    const seen = new Map();
    const visit = async (url) => {
        if (seen.has(url)) return;
        const source = await (await fetch(url)).text();
        seen.set(url, source.length);
        for (const match of source.matchAll(/(?:from|import)\s*["']([^"']+)["']/g)) {
            if (!match[1].startsWith('.') && !match[1].startsWith('/')) continue;
            await visit(new URL(match[1], new URL(url, location.href)).pathname);
        }
    };
    await visit('/word/WordViewer.js');
    const own = [...seen.keys()].filter((url) => !url.includes('/fable_modules/')).map((url) => url.replace('/word/', '')).sort();
    const editorModules = ['Word.js', 'Input.js', 'Commands.js', 'Clipboard.js', 'WordUi.js', 'WordDialogs.js', 'WordHost.js', 'WordShell.js', 'WordVba.js'];
    note('WordViewer.js imports ' + own.join(', '));
    check('the viewer module imports no editor module', own.filter((name) => editorModules.includes(name.split('/').pop())), []);
    check('it has what a viewer needs', ['Docx.js', 'Layout.js', 'Paint.js', 'Renderer.js'].every((name) => own.includes(name)), true);
    const bytes = (pick) => [...seen].filter(([url]) => pick(url)).reduce((n, [, size]) => n + size, 0);
    const editorSeen = new Map(seen);
    seen.clear();
    await visit('/word/Word.js');
    const editorBytes = bytes(() => true);
    seen.clear(); for (const [k, v] of editorSeen) seen.set(k, v);
    note(`viewer graph ${(bytes(() => true) / 1024).toFixed(0)} KB of source, editor graph ${(editorBytes / 1024).toFixed(0)} KB`);
    check('the viewer graph is smaller than the editor graph', bytes(() => true) < editorBytes, true);

    // Used from that module alone (the editor module is not loaded on this page before this).
    // (Modules the page loaded, not the sources fetched above to read their imports.)
    const loadedModules = () => performance.getEntriesByType('resource').filter((entry) => entry.initiatorType !== 'fetch').map((entry) => new URL(entry.name).pathname);
    const loadedBefore = loadedModules();
    check('this page has not loaded the editor', loadedBefore.some((path) => path === '/word/Word.js' || path === '/word/Input.js'), false);
    const { default: createViewer, createViewer: named } = await import('/word/WordViewer.js');
    check('default export and named export', [typeof createViewer, typeof named], ['function', 'function']);
    const OF = await loadFormats();
    const pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-features/test-multi-headers.docx'));
    const measure = document.createElement('canvas');
    const viewer = createViewer(measure);
    check('the whole viewer API', ['load', 'setImageCache', 'setTextMeasurers', 'setTextMeasurer', 'setTextRangeMeasurer', 'clearTextMeasurer', 'setTextPathPainter', 'getPageFrames', 'getPageCount', 'getDocumentSize', 'renderPage', 'render'].filter((name) => typeof viewer[name] !== 'function'), []);
    viewer.load(pkg);
    viewer.setImageCache(new Map());
    const pages = viewer.getPageCount();
    check('pages, frames and size', [pages > 1, viewer.getPageFrames().length, viewer.getDocumentSize().PageCount], [true, pages, pages]);
    const target = document.createElement('canvas');
    document.body.appendChild(target);
    check('renderPage paints a page', [viewer.renderPage(target, 0, 1), target.width > 100, target.height > 100], [true, true, true]);
    const data = target.getContext('2d').getImageData(0, 0, target.width, target.height).data;
    let dark = 0;
    for (let i = 0; i < data.length; i += 4) if (data[i + 3] > 0 && data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) dark++;
    check('with text on it', dark > 500, true);
    let painted = 0;
    viewer.setTextPathPainter((context) => (text, style, x, baseline, width) => { painted++; return false; });
    viewer.renderPage(target, 1, 1);
    check('a plain-function painter factory is used', painted > 5, true);
    viewer.setTextMeasurers((text, style) => text.length * style.FontSize * 0.9, (text, style, start, length) => Math.max(0, Math.min(length, text.length - start)) * style.FontSize * 0.9);
    check('plain-function measurers change the pagination', viewer.getPageCount() !== pages, true);
    viewer.clearTextMeasurer();
    check('and clearing them brings it back', viewer.getPageCount(), pages);
    const loadedAfter = loadedModules();
    check('using the viewer loaded no editor module', loadedAfter.filter((path) => editorModules.some((name) => path === '/word/' + name)), []);
    // Word.createViewer is the same viewer.
    const Word = (await import('/word/Word.js')).default;
    const fromEditor = Word.createViewer(document.createElement('canvas'));
    check('Word.createViewer gives the same API', Object.keys(fromEditor).sort(), Object.keys(viewer).sort());
    fromEditor.load(pkg);
    check('and the same pages', fromEditor.getPageCount(), pages);
} catch (e) { fail(e); }
finish();
