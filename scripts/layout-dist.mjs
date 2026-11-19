// node scripts/layout-dist.mjs <fable-out> <dist>
//
// build:fable compiles every browser project once, through src/Dist/Rbt.Dist.fsproj, so the browser
// loads one copy of each module and of fable-library whichever entry point a host imports, and
// values (lists, records, a workbook read by office-formats and installed by spreadsheet) cross
// entry points freely. Fable puts a referenced project's modules under its source folder's name
// (src/Office.Formats/src/Main.fs -> <fable-out>/Office.Formats/src/Main.js) and the packages under
// <fable-out>/fable_modules/; this moves them to dist's public folders (dist/office-formats/src/Main.js,
// dist/word/WordHost.js, dist/fable_modules/...), rewriting the relative imports between them.
//
// It then writes, as re-exporting modules (`export * from` the one module), the paths the
// per-project builds gave consumers too:
// - each bundle's copy of a project it references (dist/word/FableUI/Dom.js -> dist/fable-ui/Dom.js),
//   for the bundles that had them (not the PDF bundles, which imported their base bundle's);
// - the libraries of the other Rbt repositories at the paths a build from the neighbouring
//   checkouts gave them, under each bundle whose project uses the package
//   (dist/shaper-text/rbt-text/unicode/src/Rbt.Unicode/Ucd.js);
// - fable-library without its version in the path (dist/fable-library/List.js).
//
// A dynamic import() written in F# against the public layout (WordHost's
// `import('../office-formats/src/Main.js')`) names no module of <fable-out> and is kept as written.
import fs from "node:fs";
import path from "node:path";

const [outArg, distArg] = process.argv.slice(2);
if (!outArg || !distArg) {
  console.error("usage: layout-dist.mjs <fable-out> <dist>");
  process.exit(2);
}
const out = path.resolve(outArg);
const dist = path.resolve(distArg);
const repository = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const fail = (message) => {
  console.error(`layout-dist: ${message}`);
  process.exit(1);
};

// The browser projects: source folder (under src/) -> public folder (under dist/); a bundle had
// its own build before the umbrella and gets the re-exported copies of its references.
const projects = [
  { source: "Office.Formats", folder: "office-formats", bundle: true },
  { source: "OfficeUi", folder: "office-ui", bundle: false },
  { source: "FableUI", folder: "fable-ui", bundle: true },
  { source: "Word", folder: "word", bundle: true },
  { source: "ShaperText", folder: "shaper-text", bundle: true },
  { source: "Word.Pdf", folder: "word-pdf", bundle: true, copies: false },
  { source: "Spreadsheet", folder: "spreadsheet", bundle: true },
  { source: "Spreadsheet.Pdf", folder: "spreadsheet-pdf", bundle: true, copies: false },
];
const folders = new Map(projects.map((project) => [project.source, project.folder]));
folders.set("fable_modules", "fable_modules");
const dropped = new Set(["Dist.js"]);

// The libraries of the other Rbt repositories at their checkout paths (inside a bundle).
const siblingLayout = {
  "Rbt.Vb": "rbt-vb",
  "Rbt.Unicode.Core": "rbt-text/unicode/src/Rbt.Unicode",
  "Rbt.Bidi.Core": "rbt-text/bidi",
  "Rbt.OpenType.Core": "rbt-text/opentype",
  "Rbt.Shaper.Core": "rbt-text/shaper",
  "Rbt.Text.Engine": "rbt-text/engine",
  "Rbt.Pdf": "rbt-pdf/Rbt.Pdf",
};

function* files(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* files(full);
    else yield full;
  }
}
const posix = (p) => p.split(path.sep).join("/");
const specifier = (fromFile, toFile) => {
  const relative = posix(path.relative(path.dirname(fromFile), toFile));
  return relative.startsWith(".") ? relative : "./" + relative;
};

if (!fs.existsSync(out)) fail(`${out} does not exist`);
for (const entry of fs.readdirSync(out)) {
  if (!folders.has(entry) && !dropped.has(entry)) fail(`unexpected ${entry} in ${out} (a new project? add it to layout-dist.mjs)`);
}

// Where each module of <fable-out> goes.
const moved = new Map();
for (const file of files(out)) {
  const [top, ...rest] = posix(path.relative(out, file)).split("/");
  if (dropped.has(top) || path.basename(file) === ".gitignore") continue;
  moved.set(file, path.join(dist, folders.get(top), ...rest));
}

const importPattern = /((?:import|export)\s[^'"]*?from\s*|import\s*\(\s*|import\s+)(["'])(\.{1,2}\/[^"']+)\2/g;
let rewritten = 0;
for (const [from, to] of moved) {
  fs.mkdirSync(path.dirname(to), { recursive: true });
  if (!from.endsWith(".js")) {
    fs.copyFileSync(from, to);
    continue;
  }
  const text = fs.readFileSync(from, "utf8").replace(importPattern, (whole, lead, quote, target) => {
    const targetTo = moved.get(path.resolve(path.dirname(from), target));
    if (!targetTo) return whole;
    rewritten++;
    return lead + quote + specifier(to, targetTo) + quote;
  });
  fs.writeFileSync(to, text);
}

let reexports = 0;
const reexport = (alias, module) => {
  if (fs.existsSync(alias)) fail(`${posix(path.relative(dist, alias))} would hide a module`);
  const from = specifier(alias, module);
  const hasDefault = /^export default\b/m.test(fs.readFileSync(module, "utf8"));
  fs.mkdirSync(path.dirname(alias), { recursive: true });
  fs.writeFileSync(alias, `export * from "${from}";\n` + (hasDefault ? `export { default } from "${from}";\n` : ""));
  reexports++;
};
const modulesOf = (folder) => [...files(path.join(dist, folder))].filter((file) => file.endsWith(".js"));

// The project references of a project, transitively, read from its .fsproj.
const projectFile = (source) => {
  const dir = path.join(repository, "src", source);
  const name = fs.readdirSync(dir).find((entry) => entry.endsWith(".fsproj"));
  if (!name) fail(`no .fsproj in src/${source}`);
  return path.join(dir, name);
};
const references = (source, seen = new Set()) => {
  const fsproj = projectFile(source);
  for (const match of fs.readFileSync(fsproj, "utf8").matchAll(/<ProjectReference\s+Include="([^"]+)"/g)) {
    const referenced = path.basename(path.dirname(path.resolve(path.dirname(fsproj), match[1].replace(/\\/g, "/"))));
    if (!folders.has(referenced)) fail(`${source} references ${referenced}, which layout-dist.mjs does not know`);
    if (!seen.has(referenced)) {
      seen.add(referenced);
      references(referenced, seen);
    }
  }
  return seen;
};

// A bundle's copies of the projects it references: dist/<bundle>/<Source>/<module>.
for (const project of projects.filter((p) => p.bundle && p.copies !== false)) {
  for (const referenced of references(project.source)) {
    const folder = folders.get(referenced);
    for (const module of modulesOf(folder)) {
      reexport(path.join(dist, project.folder, referenced, path.relative(path.join(dist, folder), module)), module);
    }
  }
}

// The packages, at their checkout paths, under each bundle whose project uses them (its restore's
// project.assets.json lists the packages it uses, transitively).
const packages = path.join(dist, "fable_modules");
const packageDirs = fs.readdirSync(packages);
for (const project of projects.filter((p) => p.bundle)) {
  const assetsFile = path.join(path.dirname(projectFile(project.source)), "obj", "project.assets.json");
  if (!fs.existsSync(assetsFile)) fail(`${assetsFile} is missing (restore src/${project.source})`);
  const used = new Set(Object.keys(JSON.parse(fs.readFileSync(assetsFile, "utf8")).libraries).map((key) => key.split("/")[0].toLowerCase()));
  for (const dirName of packageDirs) {
    // "<PackageId>.<major>.<minor>.<patch>[-suffix]"
    const id = dirName.replace(/\.\d+\.\d+\.\d+(-.+)?$/, "");
    const target = siblingLayout[id];
    if (!target || !used.has(id.toLowerCase())) continue;
    const source = path.join(packages, dirName);
    for (const module of [...files(source)].filter((file) => file.endsWith(".js"))) {
      // Package sources are emitted as <name>.fs.js; checkout builds use <name>.js.
      const relative = path.relative(source, module).replace(/\.fs\.js$/, ".js");
      reexport(path.join(dist, project.folder, target, relative), module);
    }
  }
}

// fable-library at a path without its version.
const libraries = packageDirs.filter((name) => name.startsWith("fable-library-js."));
if (libraries.length !== 1) fail(`expected one fable-library in ${packages}, found ${libraries.join(", ") || "none"}`);
const library = path.join(packages, libraries[0]);
for (const module of [...files(library)].filter((file) => file.endsWith(".js"))) {
  reexport(path.join(dist, "fable-library", path.relative(library, module)), module);
}

console.log(
  `layout-dist: ${moved.size} files of one compile moved to ${posix(path.relative(process.cwd(), dist)) || "."} ` +
    `(${rewritten} imports rewritten); ${reexports} re-exporting modules at the bundles' other paths`,
);
