#!/usr/bin/env node
// Where a CPU profile spends its time in Fable's library (fable_modules/fable-library-js.*), by
// library function and by the nearest caller outside the library - the F# function whose
// spelling costs it (see docs/portable-fsharp.md).
//
// Usage: node tools/fable-cost/fable-cost.mjs profile.json [rows]
// The profile is what the browser harnesses save with PROFILE_OUT=profile.json and --profile PHASE
// (tools/word-browser/perf.mjs, tools/xlsx-browser/perf.mjs, tools/xlsx-browser/bigsheet.mjs).
import fs from 'node:fs';

const cpu = JSON.parse(fs.readFileSync(process.argv[2]));
const rows = Number(process.argv[3] ?? 25);
const byId = new Map(cpu.nodes.map((n) => [n.id, n]));
const parent = new Map();
for (const n of cpu.nodes) for (const c of n.children ?? []) parent.set(c, n.id);
const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
const isLib = (n) => /fable-library/.test(n.callFrame.url);
const isApp = (n) => /\/(spreadsheet|word|office-formats|fable-ui|shaper-text|src)\//.test(n.callFrame.url) && !isLib(n);
const name = (n) => `${n.callFrame.functionName || '(anonymous)'}  ${n.callFrame.url.split('/').slice(-1)[0]}:${n.callFrame.lineNumber + 1}`;
let total = 0, lib = 0, app = 0;
const byLib = new Map(), byModule = new Map(), byPair = new Map();
const add = (map, key) => map.set(key, (map.get(key) ?? 0) + interval);
for (const id of cpu.samples) {
    total += interval;
    const n = byId.get(id);
    if (isApp(n)) app += interval;
    if (!isLib(n)) continue;
    lib += interval;
    let p = parent.get(id);
    while (p != null && !isApp(byId.get(p))) p = parent.get(p);
    const caller = p != null ? name(byId.get(p)) : '?';
    add(byLib, name(n));
    add(byModule, n.callFrame.url.split('/').slice(-1)[0]);
    add(byPair, name(n) + '  <-  ' + caller);
}
const top = (map, n) => [...map].sort((a, b) => b[1] - a[1]).slice(0, n);
console.log(`${total.toFixed(0)} ms sampled: ${app.toFixed(0)} ms in the application's own code, ${lib.toFixed(0)} ms in Fable's library`);
console.log('-- by library module');
for (const [k, v] of top(byModule, 12)) console.log(`${v.toFixed(1).padStart(8)} ms  ${k}`);
console.log('-- by library function');
for (const [k, v] of top(byLib, 15)) console.log(`${v.toFixed(1).padStart(8)} ms  ${k}`);
console.log('-- by library function and the caller outside the library');
for (const [k, v] of top(byPair, rows)) console.log(`${v.toFixed(1).padStart(8)} ms  ${k}`);
