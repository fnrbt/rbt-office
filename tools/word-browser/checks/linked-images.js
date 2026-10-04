// A picture that is linked, not embedded, has an address the document chose. The
// editor must not request it when the document is opened; only after the host
// allows linked pictures. run.mjs counts the requests for /external-probe.png.
import { check, finish, fail, loadWord } from './common.js';
const probes = async () => Number(await (await fetch('/external-probe-count')).text());
const settle = () => new Promise((resolve) => setTimeout(resolve, 400));
try {
    const Word = await loadWord();
    const editor = Word.initialize('c');
    const run = { Text: 'before ', FontSize: 14.667, FontFamily: 'Calibri' };
    const image = (linked) => ({ IsImage: true, ImageSrc: location.origin + '/external-probe.png?linked=' + linked, ImageLinked: linked, ImageWidth: 40, ImageHeight: 40 });
    Word.importFromDocxPackage(editor, { Paragraphs: [{ Runs: [run, image(true)], Alignment: 'left' }, { Runs: [{ ...run, Text: 'unsafe link', Hyperlink: 'javascript:alert(1)' }], Alignment: 'left' }] });
    await settle();
    check('opening a document does not request its linked picture', await probes(), 0);
    check('linked pictures are off by default', Word.getAllowLinkedImages(editor), false);
    const exported = Word.exportDocumentPackage(editor);
    check('the linked picture keeps its address for saving', exported.Paragraphs[0].Runs[1].ImageSrc, location.origin + '/external-probe.png?linked=true');
    check('an unsafe link offered by a package is inactive', [exported.Paragraphs[1].Runs[0].Hyperlink ?? null, exported.Paragraphs[1].Runs[0].InactiveHyperlink], [null, 'javascript:alert(1)']);
    Word.setHyperlink(editor, 'javascript:alert(2)');
    Word.setAllowLinkedImages(editor, true);
    await settle();
    check('the host can allow linked pictures', (await probes()) >= 1, true);
    Word.setAllowLinkedImages(editor, false);
} catch (e) { fail(e); }
finish();
