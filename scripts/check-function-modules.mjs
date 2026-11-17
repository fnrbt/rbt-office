// node scripts/check-function-modules.mjs [dist/spreadsheet]
//
// The function modules (src/Spreadsheet/Functions, Formula.fs "The function
// table") are loaded by the browser only when a workbook or a formula needs
// them, through the dynamic import()s of FunctionCatalog.js. This walks the
// static imports (`import ... from`, `export ... from`, `import "..."`) from
// every module of the bundle outside Functions/ and fails when one of them
// reaches a module of Functions/ - which would load it with the grid. It also
// checks that each module the catalog imports is there and has its `install`.
import fs from "node:fs";
import path from "node:path";

const bundle = path.resolve(process.argv[2] ?? "dist/spreadsheet");
const functionsDir = path.join(bundle, "Functions");
const fail = (message) => {
  console.error(`check-function-modules: ${message}`);
  process.exit(1);
};
if (!fs.existsSync(functionsDir)) fail(`${functionsDir} does not exist (build with pnpm run build:fable first)`);

function* modules(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* modules(full);
    else if (entry.name.endsWith(".js")) yield full;
  }
}

const staticPattern = /(?:^|[;\s}])(?:import|export)\s+(?:[^'"()]*?\sfrom\s*)?(["'])(\.{1,2}\/[^"']+)\1/g;
const dynamicPattern = /import\s*\(\s*(["'])(\.{1,2}\/[^"']+)\1\s*\)/g;
const staticImports = (file) =>
  [...fs.readFileSync(file, "utf8").matchAll(staticPattern)].map((m) => path.resolve(path.dirname(file), m[2]));

const inFunctions = (file) => file.startsWith(functionsDir + path.sep);
const entries = [...modules(bundle)].filter((file) => !inFunctions(file));
const reachedFrom = new Map();
const pending = entries.map((file) => [file, file]);
while (pending.length > 0) {
  const [file, entry] = pending.pop();
  if (reachedFrom.has(file) || !fs.existsSync(file)) continue;
  reachedFrom.set(file, entry);
  for (const target of staticImports(file)) pending.push([target, entry]);
}
const leaked = [...reachedFrom.keys()].filter(inFunctions);
if (leaked.length > 0) {
  fail(
    "function modules reached through static imports:\n" +
      leaked.map((file) => `  ${path.relative(bundle, file)} (from ${path.relative(bundle, reachedFrom.get(file))})`).join("\n"),
  );
}

const catalogFile = path.join(bundle, "FunctionCatalog.js");
if (!fs.existsSync(catalogFile)) fail("FunctionCatalog.js is missing");
const loaded = [...fs.readFileSync(catalogFile, "utf8").matchAll(dynamicPattern)].map((m) => path.resolve(bundle, m[2]));
if (loaded.length === 0) fail("FunctionCatalog.js imports no function module");
for (const file of loaded) {
  if (!fs.existsSync(file)) fail(`FunctionCatalog.js imports ${path.relative(bundle, file)}, which is not there`);
  if (!/export function install\b/.test(fs.readFileSync(file, "utf8"))) fail(`${path.relative(bundle, file)} exports no install()`);
}
console.log(
  `function modules: ${loaded.length} loaded on demand (${loaded.map((f) => path.basename(f, ".js")).join(", ")}); ` +
    `none of ${reachedFrom.size - leaked.length} statically imported modules of ${entries.length} entries reaches them`,
);
