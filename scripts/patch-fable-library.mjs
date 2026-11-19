// node scripts/patch-fable-library.mjs <dist>
//
// Each bundle of dist has its own copy of fable-library (fable_modules/fable-library-js.<version>/),
// and a host may hand values of one bundle to another: the documented flow of reading a workbook with
// the office-formats bundle and converting it with the spreadsheet bundle's XlsxImporter
// (`InstallImportedWorkbook(importWorkbook(...))`) puts the reader's lists into the grid's model.
// Fable 5's FSharpList.Equals and CompareTo test `other instanceof FSharpList`, so a list of one copy
// is never equal to a list of the other (Fable 4 walked any list): `[] = []` was false, and the
// export took an unchanged page setup for a changed one. This makes the two tests accept a list of
// any copy (an object with `head`, `tail` and an iterator, as FSharpList is), as Fable 4 did.
// It fails when the code it replaces is not there, so a Fable upgrade that changes it is noticed.
import fs from "node:fs";
import path from "node:path";

const dist = path.resolve(process.argv[2] ?? "dist");
const helper =
  "\n// (patch-fable-library.mjs: a list of another bundle's copy of this library is a list too.)\n" +
  "const isAnyFSharpList = (x) => x instanceof FSharpList || (x != null && typeof x === \"object\" && \"head\" in x && \"tail\" in x && typeof x[Symbol.iterator] === \"function\");\n";
const tests = [
  ["        else if (other instanceof FSharpList) {\n            const ys = other;", "        else if (isAnyFSharpList(other)) {\n            const ys = other;"],
  ["        if (other instanceof FSharpList) {\n            const ys = other;", "        if (isAnyFSharpList(other)) {\n            const ys = other;"],
];
let patched = 0;
for (const bundle of fs.readdirSync(dist, { withFileTypes: true })) {
  const modules = path.join(dist, bundle.name, "fable_modules");
  if (!bundle.isDirectory() || !fs.existsSync(modules)) continue;
  for (const library of fs.readdirSync(modules).filter((name) => name.startsWith("fable-library-js."))) {
    const file = path.join(modules, library, "List.js");
    // (A bundle that shares its base's library, after share-bundle-modules.mjs, has none of its own.)
    if (!fs.existsSync(file)) continue;
    let text = fs.readFileSync(file, "utf8");
    if (text.includes("isAnyFSharpList")) continue;
    for (const [from, to] of tests) {
      const at = text.indexOf(from);
      if (at < 0 || text.indexOf(from, at + 1) >= 0) {
        console.error(`patch-fable-library: ${file}: expected one ${JSON.stringify(from.trim().split("\n")[0])}`);
        process.exit(1);
      }
      text = text.slice(0, at) + to + text.slice(at + from.length);
    }
    const classAt = text.indexOf("export class FSharpList {");
    if (classAt < 0) { console.error(`patch-fable-library: ${file}: no FSharpList class`); process.exit(1); }
    // (A function declared with const before the class is used only when a method runs.)
    text = text.slice(0, classAt) + helper.trimStart() + text.slice(classAt);
    fs.writeFileSync(file, text);
    patched++;
  }
}
console.log(`patch-fable-library: list equality across bundles in ${patched} copies of fable-library`);
