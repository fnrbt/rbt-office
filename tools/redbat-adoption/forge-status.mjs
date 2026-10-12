import fs from "fs";
const C = await import("/home/dave/Dev/redbat-office-viewers/node_modules/.pnpm/@rbt+client@0.1.0/node_modules/@rbt/client/dist/index.js");
const BASE = "http://10.34.110.90:18082";
const dev = fs.readFileSync("/home/dave/instances/mega-source/run/keys/dev-cookie", "utf8").trim();
let cookies = "";
const realFetch = globalThis.fetch;
globalThis.fetch = (url, init = {}) => {
  const u = typeof url === "string" && url.startsWith("/") ? BASE + url : url;
  const headers = new Headers(init.headers || {}); if (cookies) headers.set("Cookie", cookies);
  return realFetch(u, { ...init, headers });
};
const login = await realFetch(BASE + "/api/auth/dev-login", { method: "POST", headers: { "Content-Type": "application/octet-stream", Accept: "application/octet-stream", Cookie: "RedBat.Dev=" + dev }, body: C.buencode({ username: "will@cardillion.com" }) });
cookies = (login.headers.getSetCookie?.() || []).map(c => c.split(";")[0]).join("; ");
console.log("login", login.status, cookies ? "cookie ok" : "no cookie");
globalThis.__C = C;
const show = v => JSON.stringify(v, (k, x) => typeof x === "bigint" ? String(x) : (x && x.constructor && x.constructor.name === "Id" ? "id:" + globalThis.__C.Id.key(x) : x));
// args: repo commit ref action ...
const args = process.argv.slice(2);
for (let i = 0; i + 3 < args.length + 1; i += 4) {
  const [repo, commit, ref, action] = args.slice(i, i + 4);
  for (const kind of ["status", "log"]) {
    try {
      const res = await fetch(`/api/forge/repos/${repo}/actions/${kind}?commit=${commit}&ref=${encodeURIComponent(ref)}&action=${action}`);
      const text = await res.text();
      console.log(repo, action, kind, res.status, text.replace(/\s+/g, " ").slice(kind === "log" ? -1800 : 0).slice(0, 6000));
    } catch (e) { console.log(repo, action, kind, "ERR", String(e.message).slice(0, 300)); }
  }
}
