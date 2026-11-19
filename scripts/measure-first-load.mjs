// node scripts/measure-first-load.mjs <site> [--runs N] [--net slow4g|fast3g|none] [--cpu N] [sample ...]
//
// A demo site's first load (the rbt-office demo, built with Vite against a dist: see
// docs/continuation-roadmap.md) in Chromium as a visitor on a slow network and a slow machine sees it:
// the site served gzipped (as GitHub Pages serves it), the cache empty, the network throttled (CDP,
// DevTools' "Slow 4G": 150 ms, 1.6 Mbit/s down, 750 kbit/s up; "Fast 3G": 562.5 ms, 1.44 Mbit/s,
// 675 kbit/s) and the CPU slowed (4x). For each sample (opened by its link, `#budget.xlsx`) it gives:
//   bytes   what was transferred (compressed), and of it JavaScript; requests
//   drawn   ms from navigation until the sample is on screen: a sheet's grid has its cells, a
//           document's first page has its text
//   input   ms until it has taken input: a key typed into it (a cell's value, a character of the
//           document) has changed it - the page was busy until then
// Each figure is the median of the runs (default 3). PLAYWRIGHT=<playwright's index.mjs> (default
// tools/browser-runner's).
import fs from "node:fs";
import path from "node:path";
import http from "node:http";
import zlib from "node:zlib";
import { pathToFileURL } from "node:url";

const args = process.argv.slice(2);
const option = (name, fallback) => {
  const at = args.indexOf(name);
  if (at < 0) return fallback;
  const value = args[at + 1];
  args.splice(at, 2);
  return value;
};
const runs = Number(option("--runs", 3));
const net = option("--net", "slow4g");
const cpu = Number(option("--cpu", 4));
const [siteArg, ...sampleArgs] = args;
if (!siteArg) {
  console.error("usage: measure-first-load.mjs <site> [--runs N] [--net slow4g|fast3g|none] [--cpu N] [sample ...]");
  process.exit(2);
}
const site = path.resolve(siteArg);
const samples = sampleArgs.length > 0 ? sampleArgs : ["budget.xlsx", "dashboard.xlsx", "annual-review.docx"];
const here = path.dirname(new URL(import.meta.url).pathname);
const playwrightPath = process.env.PLAYWRIGHT ?? path.join(here, "../tools/browser-runner/node_modules/playwright/index.mjs");
const { chromium } = await import(pathToFileURL(path.resolve(playwrightPath)).href);

const networks = {
  slow4g: { latency: 150, downloadThroughput: (1.6e6 / 8) * 0.9, uploadThroughput: (750e3 / 8) * 0.9 },
  fast3g: { latency: 562.5, downloadThroughput: (1.44e6 / 8) * 0.9, uploadThroughput: (675e3 / 8) * 0.9 },
};
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".woff2": "font/woff2", ".png": "image/png", ".svg": "image/svg+xml", ".json": "application/json" };
const compressible = new Set([".html", ".js", ".css", ".svg", ".json", ".dat", ".xml"]);
const gzipped = new Map();
const server = http.createServer((req, res) => {
  let file = path.join(site, decodeURIComponent(new URL(req.url, "http://x").pathname));
  if (fs.existsSync(file) && fs.statSync(file).isDirectory()) file = path.join(file, "index.html");
  if (!file.startsWith(site) || !fs.existsSync(file)) {
    res.statusCode = 404;
    return res.end();
  }
  const ext = path.extname(file);
  res.setHeader("Content-Type", types[ext] ?? "application/octet-stream");
  res.setHeader("Cache-Control", "no-store");
  if (compressible.has(ext) && /\bgzip\b/.test(req.headers["accept-encoding"] ?? "")) {
    if (!gzipped.has(file)) gzipped.set(file, zlib.gzipSync(fs.readFileSync(file), { level: 9 }));
    res.setHeader("Content-Encoding", "gzip");
    return res.end(gzipped.get(file));
  }
  res.end(fs.readFileSync(file));
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const base = `http://127.0.0.1:${server.address().port}/`;
const browser = await chromium.launch();

// When the sample is on screen, and how to give it a key and see that it took it.
const isSheet = (sample) => /\.xls[xm]$/.test(sample);
const drawn = (sheet) =>
  sheet
    ? () => {
        const grid = document.getElementById("sheetgrid");
        return !!grid && !document.getElementById("sheet").hidden && grid.offsetHeight > 0 && (grid.innerText ?? "").trim().length > 20;
      }
    : () => {
        const word = document.getElementById("word");
        return !!word && !word.hidden && !!window.demo?.word && !/^Opening/.test(window.demo.word.Status.Value) && (word.innerText ?? "").trim().length > 200;
      };
const changed = (sheet) =>
  sheet ? (before) => window.demo.sheet.grid.GetRevision() !== before : () => window.demo.word.Dirty.Value === true;
const stateOf = (sheet) => (sheet ? () => window.demo.sheet.grid.GetRevision() : () => null);

async function measure(sample) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  await cdp.send("Network.enable");
  await cdp.send("Network.setCacheDisabled", { cacheDisabled: true });
  if (networks[net]) await cdp.send("Network.emulateNetworkConditions", { offline: false, ...networks[net] });
  if (cpu > 1) await cdp.send("Emulation.setCPUThrottlingRate", { rate: cpu });
  const responses = new Map();
  let bytes = 0;
  let jsBytes = 0;
  const byKind = {};
  const kindOf = (mime, url) => (/javascript/.test(mime) ? "js" : /font/.test(mime) ? "fonts" : /\.dat$/.test(url) ? "data" : /image/.test(mime) ? "images" : /\.(xlsx|xlsm|docx|docm)$/.test(url) ? "sample" : "other");
  cdp.on("Network.responseReceived", (e) => responses.set(e.requestId, kindOf(e.response.mimeType, new URL(e.response.url).pathname)));
  cdp.on("Network.loadingFinished", (e) => {
    bytes += e.encodedDataLength;
    const kind = responses.get(e.requestId) ?? "other";
    byKind[kind] = (byKind[kind] ?? 0) + e.encodedDataLength;
    if (kind === "js") jsBytes += e.encodedDataLength;
  });
  const errors = [];
  page.on("pageerror", (e) => errors.push(String(e).slice(0, 200)));
  const sheet = isSheet(sample);
  await page.goto(base + "#" + encodeURIComponent(sample));
  await page.waitForFunction(drawn(sheet), null, { polling: "raf", timeout: 180000 });
  const drawnAt = await page.evaluate(() => performance.now());
  const requests = responses.size;
  const bytesDrawn = bytes;
  const jsDrawn = jsBytes;
  const kinds = { ...byKind };
  // A key: into a cell of the sheet (one below what the samples fill), or into the document's text.
  const before = await page.evaluate(stateOf(sheet));
  const box = await page.locator(sheet ? "#sheetgrid" : "#word").boundingBox();
  await page.mouse.click(box.x + (sheet ? 120 : box.width / 2), box.y + (sheet ? box.height - 40 : Math.min(400, box.height / 2)));
  await page.keyboard.type(sheet ? "7" : "Z");
  if (sheet) await page.keyboard.press("Enter");
  await page.waitForFunction(changed(sheet), before, { polling: "raf", timeout: 180000 });
  const inputAt = await page.evaluate(() => performance.now());
  await context.close();
  if (errors.length > 0) console.error(`${sample}: page errors: ${errors.join(" | ")}`);
  return { drawn: drawnAt, input: inputAt, bytes: bytesDrawn, js: jsDrawn, requests, kinds };
}

const median = (values) => values.slice().sort((a, b) => a - b)[Math.floor(values.length / 2)];
console.log(`first load of ${path.basename(site)}: Chromium ${browser.version()}, network ${net}, CPU ${cpu}x, gzip, median of ${runs}`);
for (const sample of samples) {
  const results = [];
  for (let run = 0; run < runs; run++) results.push(await measure(sample));
  const kb = (n) => (n / 1e3).toFixed(0) + " kB";
  const s = (n) => (n / 1e3).toFixed(2) + " s";
  const m = (key) => median(results.map((r) => r[key]));
  console.log(
    `${sample.padEnd(22)} bytes ${kb(m("bytes"))} (JS ${kb(m("js"))}, ${m("requests")} requests)  drawn ${s(m("drawn"))}  input ${s(m("input"))}` +
      `   [drawn ${results.map((r) => s(r.drawn)).join(", ")}; input ${results.map((r) => s(r.input)).join(", ")}]`,
  );
  console.log(`${"".padEnd(22)} of it: ${Object.entries(results[0].kinds).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${kb(v)}`).join(", ")}`);
}
await browser.close();
server.close();
