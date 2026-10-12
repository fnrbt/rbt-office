// Serves vers/ and, under /pnpm/, the pnpm store of the worktree (every installed @rbt/office version).
import http from "node:http"; import fs from "node:fs"; import path from "node:path";
const root = new URL("./vers/", import.meta.url).pathname, pnpm = "/home/dave/Dev/redbat-office-viewers/node_modules/.pnpm/";
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css" };
http.createServer((req, res) => {
  const u = decodeURIComponent(new URL(req.url, "http://x").pathname);
  if (u.startsWith("/api/fixture/")) { const p = decodeURIComponent(u.slice(13)); return fs.existsSync(p) ? (res.writeHead(200), fs.createReadStream(p).pipe(res)) : (res.writeHead(404), res.end()); }
  const p = u.startsWith("/pnpm/") ? path.join(pnpm, u.slice(6)) : path.join(root, u === "/" ? "index.html" : u);
  if (!(p.startsWith(pnpm) || p.startsWith(root)) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end("no"); }
  res.writeHead(200, { "content-type": types[path.extname(p)] ?? "application/octet-stream" }); fs.createReadStream(p).pipe(res);
}).listen(16142, "127.0.0.1");
