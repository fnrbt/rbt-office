// Shared playwright helpers for the harness.
import { createRequire } from "node:module";
import fs from "node:fs";
const require = createRequire("/home/dave/Dev/redbat-office-viewers/mega/frontend/package.json");
export const { chromium } = require("playwright");
export const PORT = process.env.PORT || 16139;
export const BASE = `http://127.0.0.1:${PORT}`;
export const T = "/home/dave/Dev/rbt-office/tests";
export const S = "/tmp/redbat-adoption";
export const SHOTS = `${S}/shots`;
export const saved = new Map(); // id -> Buffer

export async function launch(opts = {}) {
  const browser = await chromium.launch({ headless: true, args: ["--enable-precise-memory-info"] });
  const context = await browser.newContext({ viewport: opts.viewport ?? { width: 1366, height: 768 }, permissions: ["clipboard-read", "clipboard-write"] });
  return { browser, context };
}

export const attested = { path: null };
export const fontStats = { bytes: 0, files: [] };

/**
 * page.uploads: every body POSTed to /api/upload, in order.
 * page.upload: { delay (ms before answering), fail (status to answer with), hold (a promise to await) }.
 * page.requests: every request path the page made.
 */
export async function newPage(context, opts = {}) {
  const page = await context.newPage();
  page.uploads = [];
  page.upload = { delay: 0, fail: 0, hold: null };
  page.requests = [];
  page.on("console", (m) => { if (m.type() === "error" || opts.verbose) console.log("  [console." + m.type() + "]", m.text().slice(0, 400)); });
  page.on("pageerror", (e) => console.log("  [pageerror]", String(e).slice(0, 600)));
  page.on("request", (r) => {
    const path = new URL(r.url()).pathname;
    page.requests.push(path);
    if (path.startsWith("/word-text-stack/")) {
      const size = fs.statSync("/home/dave/Dev/redbat-office-viewers/mega/frontend/public" + path).size;
      fontStats.bytes += size; fontStats.files.push(path.split("/").pop());
    }
  });
  await page.route("**/api/**", async (route) => {
    const u = new URL(route.request().url());
    if (u.pathname.startsWith("/api/fixture/")) {
      const p = decodeURIComponent(u.pathname.slice("/api/fixture/".length));
      if (!fs.existsSync(p)) return route.fulfill({ status: 404, body: "missing" });
      return route.fulfill({ status: 200, body: fs.readFileSync(p), contentType: "application/octet-stream" });
    }
    if (u.pathname.startsWith("/api/files/attested/") && attested.path) {
      return route.fulfill({ status: 200, body: fs.readFileSync(attested.path), contentType: "application/octet-stream" });
    }
    if (u.pathname.startsWith("/api/saved/")) {
      const b = saved.get(u.pathname.slice("/api/saved/".length));
      return b ? route.fulfill({ status: 200, body: b, contentType: "application/octet-stream" }) : route.fulfill({ status: 404, body: "" });
    }
    if (u.pathname === "/api/customer/profile" && page.profile) {
      const { buencode } = await import("/home/dave/Dev/redbat-office-viewers/mega/frontend/node_modules/@rbt/client/dist/index.js");
      return route.fulfill({ status: 200, body: Buffer.from(buencode(page.profile)), contentType: "application/octet-stream" });
    }
    if (u.pathname === "/api/upload") {
      const body = route.request().postDataBuffer();
      const { delay, fail, hold } = page.upload;
      page.uploadsStarted = (page.uploadsStarted ?? 0) + 1;
      if (hold) await hold;
      if (delay) await new Promise((r) => setTimeout(r, delay));
      if (fail) return route.fulfill({ status: fail, body: "no" });
      page.uploads.push(body);
      return route.fulfill({ status: 200, body: "ok" });
    }
    return route.fulfill({ status: 404, body: "" });
  });
  return page;
}

export function fixtureUrl(path) { return "/api/fixture/" + encodeURIComponent(path); }

export async function open(page, component, path, extra = {}) {
  const q = new URLSearchParams({ c: component, url: path.startsWith("/api/") ? path : fixtureUrl(path), name: extra.name ?? path.split("/").pop(), ...(extra.params ?? {}) });
  await page.goto(`${BASE}/office-harness.html?${q}`);
  await page.waitForFunction(() => window.__mounted === true, null, { timeout: 120000 });
}

/** Wait until the page has completed `n` uploads; the last body. */
export async function waitUploads(page, n, timeout = 120000) {
  const t = Date.now();
  while (page.uploads.length < n) { if (Date.now() - t > timeout) throw new Error(`upload ${n} did not arrive`); await new Promise((r) => setTimeout(r, 50)); }
  await page.waitForFunction((n) => window.__saved >= n, n, { timeout });
  return page.uploads[n - 1];
}

let failures = 0;
export function check(name, cond, detail = "") {
  if (cond) console.log("PASS", name, detail === "" ? "" : `(${typeof detail === "string" ? detail : JSON.stringify(detail)})`);
  else { failures++; console.log("FAIL", name, typeof detail === "string" ? detail : JSON.stringify(detail)); }
  return !!cond;
}
export function summary() { console.log(failures === 0 ? "ALL PASS" : `${failures} FAILED`); return failures; }

const requireOffice = createRequire("/home/dave/Dev/redbat-office-viewers/mega/frontend/node_modules/@rbt/office/package.json");
export const JSZip = requireOffice("jszip");
export async function unzip(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const out = new Map();
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) out.set(name, zip.files[name]);
  return out;
}
export const listenerProbe = () => {
  const live = new Map();
  const add = EventTarget.prototype.addEventListener, remove = EventTarget.prototype.removeEventListener;
  const cap = (o) => (typeof o === "boolean" ? o : !!(o && o.capture));
  const name = (t) => (t === window ? "window" : t === document ? "document" : t instanceof Element ? (t.isConnected ? "" : "detached:") + t.tagName.toLowerCase() + "." + String(t.className).slice(0, 40) : Object.prototype.toString.call(t));
  EventTarget.prototype.addEventListener = function (type, fn, o) {
    if (fn) { let m = live.get(this); if (!m) live.set(this, (m = new Map())); const k = type + "|" + cap(o); let s = m.get(k); if (!s) m.set(k, (s = new Set())); s.add(fn); }
    return add.call(this, type, fn, o);
  };
  EventTarget.prototype.removeEventListener = function (type, fn, o) {
    const s = live.get(this)?.get(type + "|" + cap(o)); if (s) s.delete(fn);
    return remove.call(this, type, fn, o);
  };
  window.__listeners = () => { const out = {}; for (const [t, m] of live) { if (t instanceof Element && t.isConnected && t !== document.body && t !== document.documentElement) continue; for (const [k, s] of m) if (s.size) { const key = name(t) + " " + k; out[key] = (out[key] || 0) + s.size; } } return out; };
};
