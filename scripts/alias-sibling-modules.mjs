// The libraries of the other Rbt repositories are compiled from their NuGet
// packages, so Fable emits them under fable_modules/<PackageId>.<version>/.
// Consumers import a few of those modules by the path a build from the
// neighbouring checkouts used to give them
// (e.g. shaper-text/rbt-text/unicode/src/Rbt.Unicode/Ucd.js), so recreate
// that layout as re-exporting modules.
import fs from "node:fs";
import path from "node:path";

const dist = path.resolve(process.argv[2] ?? "dist");
const layout = {
  "Rbt.Vb": "rbt-vb",
  "Rbt.Unicode.Core": "rbt-text/unicode/src/Rbt.Unicode",
  "Rbt.Bidi.Core": "rbt-text/bidi",
  "Rbt.OpenType.Core": "rbt-text/opentype",
  "Rbt.Shaper.Core": "rbt-text/shaper",
  "Rbt.Text.Engine": "rbt-text/engine",
  "Rbt.Pdf": "rbt-pdf/Rbt.Pdf",
};

function* modules(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) yield* modules(full);
    else if (entry.name.endsWith(".js")) yield full;
  }
}

let written = 0;
for (const bundle of fs.readdirSync(dist, { withFileTypes: true })) {
  const packages = path.join(dist, bundle.name, "fable_modules");
  if (!bundle.isDirectory() || !fs.existsSync(packages)) continue;
  for (const dirName of fs.readdirSync(packages)) {
    // "<PackageId>.<major>.<minor>.<patch>[-suffix]"
    const id = dirName.replace(/\.\d+\.\d+\.\d+(-.+)?$/, "");
    const target = layout[id];
    if (!target) continue;
    const source = path.join(packages, dirName);
    for (const file of modules(source)) {
      // Package sources are emitted as <name>.fs.js; checkout builds use <name>.js.
      const relative = path.relative(source, file).replace(/\.fs\.js$/, ".js");
      const alias = path.join(dist, bundle.name, target, relative);
      let from = path.relative(path.dirname(alias), file).split(path.sep).join("/");
      if (!from.startsWith(".")) from = "./" + from;
      const hasDefault = /^export default\b/m.test(fs.readFileSync(file, "utf8"));
      fs.mkdirSync(path.dirname(alias), { recursive: true });
      fs.writeFileSync(
        alias,
        `export * from "${from}";\n` + (hasDefault ? `export { default } from "${from}";\n` : ""),
      );
      written++;
    }
  }
}
if (written > 0) console.log(`aliased ${written} sibling modules to their checkout paths`);
