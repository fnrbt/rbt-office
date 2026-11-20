// Side-by-side recording: serves rbt-office's Word shell opening a showcase document, and relays Word's
// scroll position (POSTed by drive.ps1 on the VM) to the page. Reached from the VM through `ssh -R`.
import http from 'node:http'; import fs from 'node:fs'; import path from 'node:path';
const here = path.dirname(new URL(import.meta.url).pathname);
const repo = path.resolve(here, '../..');
const dist = path.resolve(process.env.DIST ?? path.join(repo, 'dist'));
const docs = path.join(repo, 'tests/test-documents-showcase');
const jszip = path.join(repo, 'node_modules/jszip/dist/jszip.min.js');
let pos = '';
http.createServer((req, res) => {
  const url = new URL(req.url, 'http://x');
  const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
  if (url.pathname === '/') return send(200, 'text/html', fs.readFileSync(path.join(here, 'page.html')));
  if (url.pathname === '/pos') {
    if (req.method === 'POST') { let b = ''; req.on('data', (c) => (b += c)); req.on('end', () => { pos = b; send(200, 'text/plain', 'ok'); }); return; }
    return send(200, 'text/plain', pos);
  }
  if (url.pathname.startsWith('/doc/')) { const f = path.join(docs, path.basename(decodeURIComponent(url.pathname))); return fs.existsSync(f) ? send(200, 'application/octet-stream', fs.readFileSync(f)) : send(404, 'text/plain', 'no doc'); }
  if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(jszip));
  if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
  const file = path.join(dist, path.normalize(url.pathname));
  const type = file.endsWith('.css') ? 'text/css' : /\.m?js$/.test(file) ? 'text/javascript' : file.endsWith('.woff2') ? 'font/woff2' : 'application/octet-stream';
  if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, type, fs.readFileSync(file));
  send(404, 'text/plain', 'not found');
}).listen(Number(process.env.PORT ?? 16990), '127.0.0.1');
