// Element / attribute inventory of a part before and after: what a save dropped or added.
import fs from "node:fs";
import { unzip } from "./lib.mjs";
const [a, b, ...parts] = process.argv.slice(2);
const za = await unzip(fs.readFileSync(a)), zb = await unzip(fs.readFileSync(b));
const inv = (xml) => { const m = new Map(); for (const t of xml.matchAll(/<([A-Za-z0-9:]+)((?:\s+[A-Za-z0-9:]+="[^"]*")*)\s*\/?>/g)) { m.set(t[1], (m.get(t[1]) ?? 0) + 1); for (const at of t[2].matchAll(/\s([A-Za-z0-9:]+)="/g)) { const k = `${t[1]}@${at[1]}`; m.set(k, (m.get(k) ?? 0) + 1); } } return m; };
for (const p of parts) {
  const x = inv(await za.get(p).async("string")), y = inv(await zb.get(p).async("string"));
  const lost = [], gained = [];
  for (const [k, n] of x) if ((y.get(k) ?? 0) < n) lost.push(`${k} ${n}->${y.get(k) ?? 0}`);
  for (const [k, n] of y) if ((x.get(k) ?? 0) < n) gained.push(`${k} ${x.get(k) ?? 0}->${n}`);
  console.log(`${p}\n  fewer: ${lost.join(", ")}\n  more: ${gained.join(", ")}`);
}
