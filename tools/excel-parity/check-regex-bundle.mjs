// node tools/excel-parity/check-regex-bundle.mjs DIST
// The regular expression engine stays out of the spreadsheet's start-up bundle: walks the static
// imports from Spreadsheet.js, Toolbar.js and XlsxImporter.js and fails when that graph holds
// Functions/FnRegex.js, Pcre.js or RegexUnicode.js, or when any of its modules holds a string of the
// engine's Unicode tables; prints the graph's module count and size, and who imports the engine.
import fs from "node:fs";
import path from "node:path";

const root = path.resolve(process.argv[2] ?? "dist", "spreadsheet");
const staticPattern = /(?:^|[;\s}])(?:import|export)\s+(?:[^'"()]*?\sfrom\s*)?(["'])(\.{1,2}\/[^"']+)\1/g;
const seen = new Set();
const pending = ["Spreadsheet.js", "Toolbar.js", "XlsxImporter.js"].map((f) => path.join(root, f));
while (pending.length > 0) {
  const file = pending.pop();
  if (seen.has(file) || !fs.existsSync(file)) continue;
  seen.add(file);
  for (const m of fs.readFileSync(file, "utf8").matchAll(staticPattern)) pending.push(path.resolve(path.dirname(file), m[2]));
}
const engine = ["Functions/FnRegex.js", "Functions/Pcre.js", "Functions/RegexUnicode.js"].map((f) => path.join(root, f));
const table = fs.readFileSync(engine[2], "utf8").match(/categoryData = "([^"]{40})/)[1];
let bytes = 0;
const bad = [];
for (const file of seen) {
  const text = fs.readFileSync(file, "utf8");
  bytes += text.length;
  if (engine.includes(file) || text.includes(table) || text.includes("compileRegex")) bad.push(path.relative(root, file));
}
console.log(`start-up graph: ${seen.size} modules, ${(bytes / 1e6).toFixed(3)} MB of JavaScript`);
for (const file of engine) console.log(`${path.relative(root, file)}: ${(fs.statSync(file).size / 1024).toFixed(1)} KB, in the start-up graph: ${seen.has(file)}`);
if (bad.length > 0) {
  console.log(`FAIL: the start-up graph holds the engine: ${bad.join(", ")}`);
  process.exit(1);
}
console.log("the engine is not in the start-up graph (no module of it, no string of its tables)");
