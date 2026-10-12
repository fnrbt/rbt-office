// A static file server for a build directory: node static.mjs dir port
import http from "node:http"; import fs from "node:fs"; import path from "node:path";
const [dir, port] = process.argv.slice(2);
const types = { ".html": "text/html", ".js": "text/javascript", ".mjs": "text/javascript", ".css": "text/css", ".woff2": "font/woff2", ".woff": "font/woff", ".ttf": "font/ttf", ".json": "application/json", ".dat": "application/octet-stream" };
http.createServer((req, res) => {
  const p = path.join(dir, decodeURIComponent(new URL(req.url, "http://x").pathname));
  if (!p.startsWith(dir) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end("no"); }
  res.writeHead(200, { "content-type": types[path.extname(p)] ?? "application/octet-stream" });
  fs.createReadStream(p).pipe(res);
}).listen(Number(port), "127.0.0.1");
