// node tools/excel-parity/regex-js-check.mjs DIST NET.jsonl
// The browser build (DIST, as scripts/build-dist.sh makes it) against .NET on the regular
// expression functions: the formulas of regex-formulas.jsonl and the engine's answers on the cases
// of regex-perl.jsonl, as regex-formulas-net.fsx wrote them for .NET in NET.jsonl. Prints every
// difference and the count; exits 1 when there is one.
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

const [dist, netFile] = process.argv.slice(2);
const sheet = path.resolve(dist, "spreadsheet");
const load = (file) => import(pathToFileURL(path.join(sheet, file)).href);
const Formula = await load("Formula.js");
const Model = await load("Model.js");
const Pcre = await load("Functions/Pcre.js");
(await load("Functions/FnRegex.js")).install();
(await load("Functions/FnText.js")).install();

const model = Model.createEmpty(40, 12);
const valueText = (v) => {
  const u = Formula.unwrapValue(v);
  switch (u.tag) {
    case 1: return "n:" + String(u.fields[0]);
    case 2: return "s:" + u.fields[0];
    case 3: return "b:" + (u.fields[0] ? "1" : "0");
    case 4: return "e:" + u.fields[0];
    default: return "x";
  }
};
const net = fs.readFileSync(netFile, "utf8").split("\n").filter((l) => l.length > 0).map((l) => JSON.parse(l));
let compared = 0;
let differ = 0;
const report = (what, js, dotnet) => {
  differ++;
  if (differ <= 20) console.log(`DIFF ${what}\n  js   ${JSON.stringify(js)}\n  .NET ${JSON.stringify(dotnet)}`);
};
for (const item of net) {
  compared++;
  if ("formula" in item) {
    const parsed = Formula.parseFormula(item.formula);
    const values = parsed.tag === 0
      ? Array.from(Formula.evalArray(parsed.fields[0], Formula.contextOf(model)), (row) => Array.from(row, valueText))
      : [["parse:" + parsed.fields[0]]];
    if (JSON.stringify(values) !== JSON.stringify(item.values)) report(item.formula, values, item.values);
  } else {
    let answer;
    try {
      const re = Pcre.compileRegex(item.pattern, item.i);
      const cps = Pcre.codePoints(item.subject);
      const first = Pcre.search(re, cps, 0);
      answer = { first: first == null ? null : Array.from(first), all: Array.from(Pcre.allMatches(re, cps), (c) => [c[0], c[1]]) };
    } catch (e) {
      answer = e instanceof Pcre.MatchLimitExceeded ? "limit" : e instanceof Pcre.PatternError ? "error" : "throw " + e;
    }
    if (JSON.stringify(answer) !== JSON.stringify(item.answer)) report(`/${item.pattern}/${item.i ? "i" : ""} on ${JSON.stringify(item.subject)}`, answer, item.answer);
  }
}
console.log(`${compared} compared: ${differ} differ between the browser build and .NET`);
process.exit(differ > 0 ? 1 : 0);
