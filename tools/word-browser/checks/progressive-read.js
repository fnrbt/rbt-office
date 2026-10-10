// A DOCX opened front first (officeFormats.word.openProgressive): the package
// is handed over when the parts around the body and the first blocks are read,
// the rest is read in slices; read to its end it is, field for field, the
// package forEditorPackage makes - for every test document, whatever the
// size of the steps.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    // (A package as text: bytes by their length and a sum, the reading state left out.)
    const textOf = (pkg) => JSON.stringify(pkg, (key, value) => {
        if (key === 'Progressive') return undefined;
        if (value instanceof Uint8Array) { let sum = 0; for (let i = 0; i < value.length; i++) sum = (sum * 31 + value[i]) >>> 0; return `bytes:${value.length}:${sum}`; }
        return value;
    });
    const firstDifference = (a, b) => { let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++; return `at ${i}: ...${a.slice(Math.max(0, i - 60), i + 80)}... against ...${b.slice(Math.max(0, i - 60), i + 80)}...`; };
    const corpus = await (await fetch('/corpus-list')).json();
    let same = 0, frontFirst = 0, differing = [];
    for (const name of corpus) {
        const bytes = await bytesOf('/repo/' + name);
        const whole = textOf(await OF.processors.forEditorPackage(bytes));
        for (const [label, read] of [
            ['in steps', async (pkg) => { while (!pkg.Progressive.IsComplete) pkg.Progressive.readMore(0); }],
            ['at once', async (pkg) => { pkg.Progressive.complete(); }],
            ['in slices', async (pkg) => { pkg.Progressive.start(); await OF.word.whenLoadComplete(pkg); }],
        ]) {
            const pkg = await OF.word.openProgressive(bytes, { autoStart: false, firstBlocks: label === 'in steps' ? 3 : 64 });
            if (label === 'in steps' && !pkg.Progressive.AtOnce) frontFirst++;
            // What is there at first is the start of what is there at the end.
            const firstItems = pkg.Paragraphs.length, firstText = JSON.stringify(pkg.Paragraphs);
            await read(pkg);
            const text = textOf(pkg);
            if (text === whole && OF.word.isLoadComplete(pkg) && JSON.stringify(pkg.Paragraphs.slice(0, firstItems)) === firstText) same++;
            else differing.push(`${name} (${label}): ${text === whole ? 'the first items changed, or the package is not complete' : firstDifference(whole, text)}`);
        }
    }
    check(`the package read front first is the package read at once (${corpus.length} documents, three ways each)`, [same, differing.slice(0, 5)], [corpus.length * 3, []]);
    check(`most documents are read front first (${frontFirst} of ${corpus.length})`, frontFirst >= corpus.length * 0.7, true);
} catch (e) { fail(e); }
finish();
