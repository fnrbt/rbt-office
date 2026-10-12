import { unzip } from "./lib.mjs"; import fs from "node:fs";
const [fa, fb, n] = process.argv.slice(2);
const a = await (await unzip(fs.readFileSync(fa))).get(n).async("string"), b = await (await unzip(fs.readFileSync(fb))).get(n).async("string");
let pre = 0; while (pre < Math.min(a.length, b.length) && a[pre] === b[pre]) pre++;
let suf = 0; while (suf < Math.min(a.length, b.length) - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
console.log(n, a.length, "->", b.length, "prefix", pre, "suffix", suf); console.log("opened:", JSON.stringify(a.slice(Math.max(0, pre - 80), a.length - suf + 40).slice(0, 1500))); console.log("saved: ", JSON.stringify(b.slice(Math.max(0, pre - 80), b.length - suf + 40).slice(0, 1500)));
