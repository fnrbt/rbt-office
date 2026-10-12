// The package's synthetic DOCX fixtures (tools/docx-roundtrip/fixtures.mjs), written to ./synth
import fs from "node:fs"; import path from "node:path";
import { JSZip, S } from "./lib.mjs";
const { syntheticFixtures } = await import("/home/dave/Dev/rbt-office/tools/docx-roundtrip/fixtures.mjs");
fs.mkdirSync(`${S}/synth`, { recursive: true });
for (const f of await syntheticFixtures(JSZip)) { const p = `${S}/synth/${f.name.replace(/^synthetic\//, "")}`; fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, f.data); }
console.log(fs.readdirSync(`${S}/synth`).join(" "));
