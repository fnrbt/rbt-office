import { unzip } from "./lib.mjs"; import fs from "node:fs";
const [f, n, re, max] = process.argv.slice(2);
const s = await (await unzip(fs.readFileSync(f))).get(n).async("string");
const m = [...s.matchAll(new RegExp(re, "g"))].map((x) => x[0]); console.log(m.length, JSON.stringify(m.slice(-(Number(max) || 8))));
