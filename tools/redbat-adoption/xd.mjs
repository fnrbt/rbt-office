import { unzip } from "./lib.mjs"; import fs from "node:fs";
const [fa, fb, ...names] = process.argv.slice(2);
const a = await unzip(fs.readFileSync(fa)), b = await unzip(fs.readFileSync(fb));
for (const n of names) { console.log("=====", n); console.log("--- opened:\n" + (await a.get(n).async("string")).slice(0, 6000)); console.log("--- saved:\n" + (await b.get(n).async("string")).slice(0, 6000)); }
