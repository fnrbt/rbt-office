// Dump our paragraph layout for a fixture: block Y/heights, to debug
// missing/mis-placed content. Usage: node dump-layout.mjs <fixture.docx>
import { spawn } from 'node:child_process';
import path from 'node:path';
import { createRequire } from 'node:module';
const here = path.dirname(new URL(import.meta.url).pathname);
const runnerRoot = path.resolve(here, '../browser-runner');
const fixture = path.resolve(process.argv[2]);
const port = 5191;
const req = createRequire(path.join(runnerRoot, 'package.json'));
const { chromium } = req('playwright');
const server = spawn(path.join(runnerRoot, 'node_modules/.bin/vite'), ['--host','127.0.0.1','--port',String(port)], { cwd: runnerRoot, stdio:['ignore','pipe','pipe'] });
let ready=false; server.stdout.on('data',c=>{ if(c.toString().includes('Local:')) ready=true; });
const t0=Date.now(); while(!ready && Date.now()-t0<30000) await new Promise(r=>setTimeout(r,150));
const browser = await chromium.launch();
const page = await browser.newPage({ viewport:{width:1100,height:1800} });
page.on('pageerror',e=>console.error('PAGEERR',e?.stack||e?.message));
await page.goto(`http://127.0.0.1:${port}/?doc=/@fs/${fixture}`);
await page.waitForFunction(()=>{ const s=document.querySelector('#status')?.textContent||''; return window.rbWordRunner && (s.startsWith('Loaded ')||s.startsWith('Load failed:')); }, null, {timeout:60000});
console.log('status:', await page.locator('#status').textContent());
const segJson = await page.evaluate(()=> window.rbWordRunner.getLayoutJson());
const arr = JSON.parse(segJson);
const items = arr.map(s=>({y:Math.round(s.Y),h:Math.round(s.Height||0),x:Math.round(s.X),kind:s.Kind,block:s.BlockId,text:(s.Text||'').slice(0,24)}));
items.sort((a,b)=>a.y-b.y);
console.log('paragraph segments:', items.length);
for (const it of items) console.log(`y=${it.y} h=${it.h} x=${it.x} ${it.kind} blk=${it.block} "${it.text}"`);
await browser.close(); server.kill(); process.exit(0);
