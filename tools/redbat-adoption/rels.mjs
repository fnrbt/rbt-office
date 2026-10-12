// node rels.mjs file.zip : relationships whose target part is missing, parts nothing refers to, content-type overrides of missing parts
import { unzip } from "./lib.mjs"; import fs from "node:fs"; import path from "node:path";
for (const file of process.argv.slice(2)) {
  const zip = await unzip(fs.readFileSync(file));
  const names = new Set(zip.keys()); const referenced = new Set(); const dangling = [];
  for (const name of names) if (name.endsWith(".rels")) {
    const xml = await zip.get(name).async("string");
    const base = path.posix.dirname(path.posix.dirname(name));
    for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
      if (/TargetMode="External"/.test(m[0])) continue;
      const t = /Target="([^"]*)"/.exec(m[0])?.[1]; if (!t) continue;
      const target = t.startsWith("/") ? t.slice(1) : path.posix.normalize(path.posix.join(base === "." ? "" : base, t));
      referenced.add(target);
      if (!names.has(target)) dangling.push(`${name} -> ${target}`);
    }
  }
  const ct = await zip.get("[Content_Types].xml").async("string");
  const overrides = [...ct.matchAll(/<Override\b[^>]*PartName="\/([^"]*)"/g)].map((m) => m[1]).filter((n) => !names.has(n));
  const orphans = [...names].filter((n) => !n.endsWith(".rels") && n !== "[Content_Types].xml" && !referenced.has(n));
  // r:id used in a part without a relationship of that id
  const missingIds = [];
  for (const name of names) if (/\.xml$|\.vml$/.test(name) && !name.endsWith(".rels") && zip.get(name)._data?.uncompressedSize < 3e6) {
    const xml = await zip.get(name).async("string");
    const ids = new Set([...xml.matchAll(/\br:(?:id|embed|link|pict)="([^"]+)"/g)].map((m) => m[1]));
    if (ids.size === 0) continue;
    const relName = path.posix.join(path.posix.dirname(name), "_rels", path.posix.basename(name) + ".rels");
    const rel = names.has(relName) ? await zip.get(relName).async("string") : "";
    const have = new Set([...rel.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
    const missing = [...ids].filter((id) => !have.has(id));
    if (missing.length) missingIds.push(`${name}: ${missing.join(",")}`);
  }
  console.log(path.basename(file), JSON.stringify({ parts: names.size, danglingRelationships: dangling, overridesOfMissingParts: overrides, unreferencedParts: orphans, relationshipIdsWithoutRelationship: missingIds }, null, 1));
}
