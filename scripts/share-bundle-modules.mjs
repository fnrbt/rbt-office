// node scripts/share-bundle-modules.mjs <bundle> <base> <entry...>
//
// Fable compiles a project with its project references, so a bundle of a project
// that builds on another one (spreadsheet-pdf on spreadsheet) carries its own copy
// of the other one's modules and of fable-library. Loaded beside the base bundle,
// that is a second copy of every module - and of its classes, which values made by
// the first are then handed to. This points every import of a module the base
// bundle has too (the same path inside it: `Spreadsheet/Model.js` is the base's
// `Model.js`, `fable_modules/...` its `fable_modules/...`) at the base's, and
// deletes what the entries no longer reach: the bundle keeps only its own modules
// and those of packages the base does not have (Rbt.Pdf).
import fs from "node:fs";
import path from "node:path";

const [bundleArg, baseArg, ...entryArgs] = process.argv.slice(2);
if (!bundleArg || !baseArg || entryArgs.length === 0) {
  console.error("usage: share-bundle-modules.mjs <bundle> <base> <entry...>");
  process.exit(2);
}
const bundle = path.resolve(bundleArg);
const base = path.resolve(baseArg);
// Where a module of the bundle is in the base: the referenced project's folder
// (named as the project's folder: Spreadsheet/) is the base's root.
const ownFolder = path.basename(base).toLowerCase();
const inBase = (relative) => {
  const parts = relative.split("/");
  const candidates = [relative];
  if (parts.length > 1 && parts[0].toLowerCase() === ownFolder) candidates.unshift(parts.slice(1).join("/"));
  return candidates.map((candidate) => path.join(base, candidate)).find((file) => fs.existsSync(file));
};

function* modules(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* modules(full);
    else if (entry.name.endsWith(".js")) yield full;
  }
}

const importPattern = /((?:import|export)\s[^'"]*?from\s*|import\s*\(\s*|import\s+)(["'])(\.{1,2}\/[^"']+)\2/g;
let rewritten = 0;
for (const file of modules(bundle)) {
  const text = fs.readFileSync(file, "utf8");
  const next = text.replace(importPattern, (whole, lead, quote, specifier) => {
    const target = path.resolve(path.dirname(file), specifier);
    const relative = path.relative(bundle, target).split(path.sep).join("/");
    if (relative.startsWith("..")) return whole;
    const shared = inBase(relative);
    if (!shared) return whole;
    let pointer = path.relative(path.dirname(file), shared).split(path.sep).join("/");
    if (!pointer.startsWith(".")) pointer = "./" + pointer;
    rewritten++;
    return lead + quote + pointer + quote;
  });
  if (next !== text) fs.writeFileSync(file, next);
}

// What the entries reach inside the bundle; the rest goes.
const reached = new Set();
const visit = (file) => {
  if (reached.has(file) || !file.startsWith(bundle) || !fs.existsSync(file)) return;
  reached.add(file);
  for (const match of fs.readFileSync(file, "utf8").matchAll(importPattern)) visit(path.resolve(path.dirname(file), match[3]));
};
for (const entry of entryArgs) visit(path.resolve(bundle, entry));
let removed = 0;
for (const file of [...modules(bundle)]) {
  if (!reached.has(file)) { fs.rmSync(file); removed++; }
}
const prune = (dir) => {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) if (entry.isDirectory()) prune(path.join(dir, entry.name));
  if (dir !== bundle && fs.readdirSync(dir).length === 0) fs.rmdirSync(dir);
};
prune(bundle);
console.log(`shared ${rewritten} imports of ${path.basename(bundle)} with ${path.basename(base)}; ${reached.size} modules kept, ${removed} removed`);
