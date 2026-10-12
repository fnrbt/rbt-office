// element and attribute histogram differences of one part in two packages
import { unzip } from "./lib.mjs"; import fs from "node:fs";
const [fa, fb, n] = process.argv.slice(2);
const a = await (await unzip(fs.readFileSync(fa))).get(n).async("string"), b = await (await unzip(fs.readFileSync(fb))).get(n).async("string");
const hist = (s) => { const h = new Map(); for (const m of s.matchAll(/<([A-Za-z0-9:]+)((?:\s+[A-Za-z0-9:]+="[^"]*")*)\s*\/?>/g)) { h.set("<" + m[1] + ">", (h.get("<" + m[1] + ">") ?? 0) + 1); for (const at of m[2].matchAll(/\s([A-Za-z0-9:]+)="/g)) { const k = m[1] + "@" + at[1]; h.set(k, (h.get(k) ?? 0) + 1); } } return h; };
const ha = hist(a), hb = hist(b); const out = [];
for (const k of new Set([...ha.keys(), ...hb.keys()])) if ((ha.get(k) ?? 0) !== (hb.get(k) ?? 0)) out.push(`${k} ${ha.get(k) ?? 0}->${hb.get(k) ?? 0}`);
console.log(n, a.length, "->", b.length, "\n ", out.join("; "));
