import { check, finish, fail, loadWord } from './common.js';
const wait = async predicate => {
    for (let i = 0; i < 1000; i++) { if (predicate()) return; await new Promise(resolve => setTimeout(resolve, 10)); }
    throw new Error('PDF export did not settle');
};
let dispose;
const open = window.open, createUrl = URL.createObjectURL;
try {
    const [Word, Host, Shell, Dom] = await Promise.all([
        loadWord(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'pdf-shell'; document.body.appendChild(root);
    const host = Host.create('PDF SNAPSHOT', 'word-shell');
    dispose = Dom.mountApp('#pdf-shell', Shell.view(host));
    const editor = host.Editor;
    const blobs = [], popups = [];
    URL.createObjectURL = blob => { blobs.push(blob); return createUrl.call(URL, blob); };
    window.open = () => {
        const popup = {closed:false, opener:null, location:{replace(url) { popup.url=url; }}, close() {this.closed=true;}};
        popups.push(popup); return popup;
    };
    const print = async () => {
        root.querySelector('[data-testid="word-file-menu"]').click();
        await wait(() => root.querySelector('[data-testid="word-menu-print-pdf"]'));
        root.querySelector('[data-testid="word-menu-print-pdf"]').click();
    };
    await print(); await wait(() => !host.Printing);
    check('missing fonts reject before opening a blank tab', [popups.length, blobs.length, /font/i.test(host.Status.Value)], [0,0,true]);
    let response;
    for (const name of ['LiberationSans-Regular.ttf','calibri.ttf']) {
        response = await fetch('/font?name='+name); if (response.ok) break;
    }
    if (!response.ok) throw new Error('Configure RB_WORD_FONT_DIR with a PDF TrueType font');
    const fonts = [{family:'Calibri',data:new Uint8Array(await response.arrayBuffer()),bold:false,italic:false}];
    Host.setPdfFonts(host, fonts);
    const popupOpen = window.open;
    window.open = () => null;
    await print(); await wait(() => !host.Printing);
    check('blocked popup reports failure without claiming a PDF', [blobs.length, /block/i.test(host.Status.Value)], [0,true]);
    window.open = popupOpen;
    await print(); Host.printPdf(host);
    Word.insertText(editor,' LATER EDIT');
    await wait(() => !host.Printing);
    check('overlapping print requests produce one PDF tab', [popups.length, blobs.length, popups[0]?.closed, popups[0]?.url?.startsWith('blob:')], [1,1,false,true]);
    const bytes = new Uint8Array(await blobs[0].arrayBuffer());
    const pdf = new TextDecoder('latin1').decode(bytes);
    check('real PDF contains one page and embedded font', [pdf.startsWith('%PDF-'), (pdf.match(/\/Type\s*\/Page\b/g)||[]).length, /\/FontFile2\b/.test(pdf)], [true,1,true]);
    check('PDF export does not mark document edits saved', Word.isModified(editor), true);
    const before = blobs.length;
    await print();
    host.Status.Value = 'Unmounted PDF host';
    dispose(); dispose = null;
    await wait(() => !host.Printing);
    check('unmounted export closes its reserved tab without publishing', [blobs.length, popups.at(-1).closed, host.Status.Value], [before,true,'Unmounted PDF host']);
} catch(error) { fail(error); }
finally { window.open=open; URL.createObjectURL=createUrl; if(dispose) dispose(); finish(); }
