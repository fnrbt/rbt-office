import fs from "node:fs"; import zlib from "node:zlib"; import path from "node:path";
for (const d of ["before", "after"]) {
  const dir = `build-${d}/assets`; const files = fs.readdirSync(dir).filter((f) => f.endsWith(".js"));
  const src = Object.fromEntries(files.map((f) => [f, fs.readFileSync(path.join(dir, f), "utf8")]));
  const stat = (f) => { const m = [...src[f].matchAll(/(?:import|export)\s*(?:[^"'()]*?from\s*)?["']\.\/([^"']+\.js)["']/g)].map((x) => x[1]); return [...new Set(m)].filter((x) => src[x]); };
  const dyn = (f) => [...new Set([...src[f].matchAll(/import\(\s*["']\.\/([^"']+\.js)["']\s*\)/g)].map((x) => x[1]))].filter((x) => src[x]);
  const closure = (roots) => { const seen = new Set(); const go = (f) => { if (seen.has(f)) return; seen.add(f); stat(f).forEach(go); }; roots.forEach(go); return seen; };
  const size = (set) => [...set].reduce((n, f) => n + src[f].length, 0);
  const gz = (set) => [...set].reduce((n, f) => n + zlib.gzipSync(src[f]).length, 0);
  const main = fs.readFileSync(`build-${d}/index.html`, "utf8").match(/assets\/(index-[^"]+\.js)/)[1];
  const shell = closure([main]);
  console.log(`== ${d}: app shell entry ${main} ${src[main].length} B; with its static imports ${size(shell)} B (${shell.size} chunks)`);
  const find = (prefix) => files.filter((f) => f.startsWith(prefix + "-"));
  for (const name of ["WordDocumentContent", "SpreadsheetContent", "WordEditor", "SpreadsheetEditor", "ChatRoute", "FilesApp"]) {
    for (const f of find(name)) {
      const c = closure([f]); const extra = new Set([...c].filter((x) => !shell.has(x)));
      // dynamic imports reachable from the extra chunks (text stack)
      const dyns = new Set(); for (const x of extra) dyn(x).forEach((y) => { if (!shell.has(y)) dyns.add(y); });
      const dynClosure = new Set([...closure([...dyns])].filter((x) => !shell.has(x) && !extra.has(x)));
      console.log(`  ${name}: +${(size(extra) / 1000).toFixed(0)} kB (${(gz(extra) / 1000).toFixed(0)} kB gzip) in ${extra.size} chunks beyond the shell: ${[...extra].map((x) => x.replace(/-[^-]+\.js$/, "") + ":" + Math.round(src[x].length / 1000)).join(", ")}${name.includes("Route") || name.includes("App") ? "" : `; on demand: ${[...dynClosure].map((x) => x.replace(/-[^-]+\.js$/, "") + ":" + Math.round(src[x].length / 1000)).join(", ")}`}`);
    }
  }
}
