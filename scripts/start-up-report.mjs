// node scripts/start-up-report.mjs <bundle> <entry.js,...> [--top N] [--min]
//
// What a bundle's entry modules load before the page can use them, biggest first: each module of
// the static import graph (the same walk as check-start-up-modules.mjs: `import ... from`,
// `export ... from`, `import "..."`, transitively; a dynamic import() loads only when it runs) with
// its size and the import chain that pulls it in. With --min it also gives each module's size
// minified and gzipped (esbuild's minifier, as Vite's build uses it; esbuild comes from
// tools/browser-runner/node_modules or ESBUILD=<path to esbuild's main module>), and the totals.
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { pathToFileURL } from "node:url";

const args = process.argv.slice(2);
const flag = (name) => args.includes(name);
const option = (name, fallback) => (args.indexOf(name) >= 0 ? args[args.indexOf(name) + 1] : fallback);
const [bundleArg, entriesArg] = args.filter((a, i) => !a.startsWith("--") && !(i > 0 && args[i - 1] === "--top"));
if (!bundleArg || !entriesArg) {
  console.error("usage: start-up-report.mjs <bundle> <entry.js,...> [--top N] [--min]");
  process.exit(2);
}
const top = Number(option("--top", 25));
const bundle = path.resolve(bundleArg);
const entries = entriesArg.split(",").map((entry) => path.join(bundle, entry));
const staticPattern = /(?:^|[;\s}])(?:import|export)\s+(?:[^'"()]*?\sfrom\s*)?(["'])(\.{1,2}\/[^"']+)\1/g;
const dynamicPattern = /\bimport\((["'])(\.{1,2}\/[^"']+)\1\)/g;

const reachedFrom = new Map();
const pending = entries.map((file) => [file, null]);
while (pending.length > 0) {
  const [file, from] = pending.shift();
  if (reachedFrom.has(file) || !fs.existsSync(file)) continue;
  reachedFrom.set(file, from);
  for (const m of fs.readFileSync(file, "utf8").matchAll(staticPattern)) pending.push([path.resolve(path.dirname(file), m[2]), file]);
}
const rel = (file) => path.relative(bundle, file);
const chain = (file) => {
  const steps = [];
  for (let at = file; at; at = reachedFrom.get(at)) steps.unshift(path.basename(at));
  return steps.join(" > ");
};

let minify = null;
if (flag("--min")) {
  const here = path.dirname(new URL(import.meta.url).pathname);
  const esbuildPath = process.env.ESBUILD ?? path.join(here, "../tools/browser-runner/node_modules/esbuild/lib/main.js");
  const esbuild = await import(pathToFileURL(esbuildPath).href);
  minify = (code) => (esbuild.transformSync ?? esbuild.default.transformSync)(code, { minify: true, format: "esm", target: "es2022", loader: "js" }).code;
}
const rows = [...reachedFrom.keys()].map((file) => {
  const code = fs.readFileSync(file, "utf8");
  const row = { file, bytes: Buffer.byteLength(code) };
  if (minify) {
    const small = minify(code);
    row.min = Buffer.byteLength(small);
    row.gz = zlib.gzipSync(small, { level: 9 }).length;
  }
  return row;
});
rows.sort((a, b) => b.bytes - a.bytes);
const mb = (n) => (n / 1e6).toFixed(2) + " MB";
const kb = (n) => (n / 1e3).toFixed(0).padStart(6) + " kB";
const sum = (key) => rows.reduce((s, r) => s + (r[key] ?? 0), 0);
console.log(`${path.basename(bundle)}: ${rows.length} modules, ${mb(sum("bytes"))}${minify ? `, ${mb(sum("min"))} minified, ${mb(sum("gz"))} minified+gzipped (per module)` : ""} load with ${entries.map(rel).join(", ")}`);
for (const r of rows.slice(0, top)) {
  console.log(`${kb(r.bytes)}${minify ? ` ${kb(r.min)} min ${kb(r.gz)} gz` : ""}  ${rel(r.file)}\n          ${chain(r.file)}`);
}
// What the start-up modules load on use (their dynamic imports), for the record.
const onUse = new Set();
for (const file of reachedFrom.keys()) {
  for (const m of fs.readFileSync(file, "utf8").matchAll(dynamicPattern)) {
    const target = path.resolve(path.dirname(file), m[2]);
    if (!reachedFrom.has(target)) onUse.add(rel(target));
  }
}
if (onUse.size > 0) console.log(`on use (dynamic import): ${[...onUse].sort().join(", ")}`);
