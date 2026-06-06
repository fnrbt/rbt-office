import { renderAsync } from '@eigenpal/docx-editor-react';
import '@eigenpal/docx-editor-react/styles.css';

const statusEl = document.getElementById('status');
const setStatus = (t) => { statusEl.textContent = t; };

// Inject @font-face for the Windows fonts so docx-editor rasterizes with the
// same faces as MS Word and our renderer (else headless Chromium substitutes
// and the comparison is unfair). ?fonts=/@fs/<dir> points at a folder of TTFs.
function injectFonts(fontsBase) {
  if (!fontsBase) return;
  // EXACT mirror of the @font-face set our browser-runner registers
  // (tools/browser-runner/Program.fs) so both renderers rasterize with
  // identical Windows faces and the comparison is font-fair.
  const faces = [
    ['Calibri', 'calibri.ttf', 400, 'normal'],
    ['Calibri', 'calibrib.ttf', 700, 'normal'],
    ['Calibri', 'calibrii.ttf', 400, 'italic'],
    ['Calibri', 'calibriz.ttf', 700, 'italic'],
    ['Cambria', 'cambria.ttc', 400, 'normal'],
    ['Cambria', 'cambriab.ttf', 700, 'normal'],
    ['Cambria', 'cambriai.ttf', 400, 'italic'],
    ['Arial', 'arial.ttf', 400, 'normal'],
    ['Arial', 'arialbd.ttf', 700, 'normal'],
    ['Arial', 'ariali.ttf', 400, 'italic'],
    ['Arial', 'arialbi.ttf', 700, 'italic'],
    ['Times New Roman', 'times.ttf', 400, 'normal'],
    ['Times New Roman', 'timesbd.ttf', 700, 'normal'],
    ['Times New Roman', 'timesi.ttf', 400, 'italic'],
    ['Times New Roman', 'timesbi.ttf', 700, 'italic'],
  ];
  const css = faces
    .map(([fam, file, weight, style]) =>
      `@font-face{font-family:'${fam}';src:url('${fontsBase}/${file}') format('truetype');font-weight:${weight};font-style:${style};font-display:block;}`)
    .join('\n');
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  document.head.appendChild(styleEl);
}

async function boot() {
  const params = new URLSearchParams(location.search);
  const docUrl = params.get('doc');
  const fontsBase = params.get('fonts');
  const zoom = Number(params.get('zoom') || '1') || 1;
  injectFonts(fontsBase);
  if (!docUrl) { setStatus('no ?doc='); return; }

  try {
    setStatus('fetching…');
    const res = await fetch(docUrl);
    if (!res.ok) throw new Error('fetch ' + res.status);
    const buf = await res.arrayBuffer();
    setStatus('rendering…');
    const container = document.getElementById('container');
    // NOTE: renderAsync's promise only resolves on the first onChange, which
    // never fires in readOnly mode — so we DON'T await it for readiness. We
    // fire it (capturing errors / the handle) and poll the DOM for painted
    // pages instead.
    renderAsync(buf, container, {
      readOnly: true,
      showToolbar: false,
      onError: (e) => { window.__docxEditorError = String(e && e.stack || e); setStatus('error: ' + e); },
    })
      .then((editor) => {
        window.__docxEditor = editor;
        if (editor && typeof editor.setZoom === 'function' && zoom !== 1) editor.setZoom(zoom);
      })
      .catch((e) => { window.__docxEditorError = String(e && e.stack || e); });

    // Poll for the painter's first page.
    const deadline = Date.now() + 40000;
    while (Date.now() < deadline) {
      if (window.__docxEditorError) return;
      if (document.querySelector('.layout-page')) break;
      await new Promise((r) => setTimeout(r, 150));
    }
    if (!document.querySelector('.layout-page')) throw new Error('no .layout-page rendered');
    if (window.__docxEditor && typeof window.__docxEditor.setZoom === 'function' && zoom !== 1) {
      window.__docxEditor.setZoom(zoom);
      await new Promise((r) => setTimeout(r, 200));
    }
    // Let fonts settle + layout flush. Bound fonts.ready with a race so a
    // missing @font-face URL (404) can't hang us forever.
    if (document.fonts && document.fonts.ready) {
      await Promise.race([document.fonts.ready, new Promise((r) => setTimeout(r, 1500))]);
    }
    await new Promise((r) => setTimeout(r, 400));
    window.__docxEditorReady = true;
    setStatus('ready');
  } catch (err) {
    window.__docxEditorError = String(err && err.stack || err);
    setStatus('error: ' + (err && err.message || err));
    console.error(err);
  }
}

boot();
