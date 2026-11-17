// node scripts/check-start-up-modules.mjs <bundle> <entry.js,...> <forbidden-pattern,...>
//
// What the browser fetches when a bundle's entry modules load is what they import statically
// (`import ... from`, `export ... from`, `import "..."`), transitively; a dynamic import() is fetched
// only when it runs. This walks those static imports from the entries and fails when a module whose
// path matches one of the forbidden patterns is reached - code that must load only on use (the VBA
// engine, MacroProject.js) would then load with the grid. It prints what the entries load.
import fs from "node:fs";
import path from "node:path";

const [bundleArg, entriesArg, forbiddenArg] = process.argv.slice(2);
if (!bundleArg || !entriesArg || !forbiddenArg) {
  console.error("usage: check-start-up-modules.mjs <bundle> <entry.js,...> <forbidden-pattern,...>");
  process.exit(2);
}
const bundle = path.resolve(bundleArg);
const entries = entriesArg.split(",").map((entry) => path.join(bundle, entry));
const forbidden = forbiddenArg.split(",").map((pattern) => new RegExp(pattern));
const staticPattern = /(?:^|[;\s}])(?:import|export)\s+(?:[^'"()]*?\sfrom\s*)?(["'])(\.{1,2}\/[^"']+)\1/g;

const reachedFrom = new Map();
const pending = entries.map((file) => [file, null]);
while (pending.length > 0) {
  const [file, from] = pending.pop();
  if (reachedFrom.has(file) || !fs.existsSync(file)) continue;
  reachedFrom.set(file, from);
  for (const m of fs.readFileSync(file, "utf8").matchAll(staticPattern)) pending.push([path.resolve(path.dirname(file), m[2]), file]);
}
const missing = entries.filter((file) => !fs.existsSync(file));
if (missing.length > 0) {
  console.error(`check-start-up-modules: missing entries: ${missing.map((f) => path.relative(bundle, f)).join(", ")}`);
  process.exit(1);
}
const chain = (file) => {
  const steps = [];
  for (let at = file; at; at = reachedFrom.get(at)) steps.unshift(path.relative(bundle, at));
  return steps.join(" -> ");
};
const leaked = [...reachedFrom.keys()].filter((file) => forbidden.some((pattern) => pattern.test(path.relative(bundle, file))));
if (leaked.length > 0) {
  console.error(`check-start-up-modules: loaded at start but meant to load on use:\n${leaked.map((file) => "  " + chain(file)).join("\n")}`);
  process.exit(1);
}
const bytes = [...reachedFrom.keys()].reduce((sum, file) => sum + fs.statSync(file).size, 0);
console.log(`check-start-up-modules: ${path.basename(bundle)}: ${reachedFrom.size} modules (${(bytes / 1e6).toFixed(2)} MB) load with ${entries.map((f) => path.relative(bundle, f)).join(", ")}; none of ${forbidden.length} on-use patterns among them`);
