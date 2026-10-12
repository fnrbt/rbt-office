import fs from "node:fs";
import { unzip } from "./lib.mjs";
for (const f of process.argv.slice(2)) {
  const z = await unzip(fs.readFileSync(f));
  const styles = await z.get("xl/styles.xml").async("string");
  const dxfs = Number(/<dxfs count="(\d+)"/.exec(styles)?.[1] ?? (styles.match(/<dxf>|<dxf /g) ?? []).length);
  const tableStyles = [...styles.matchAll(/<tableStyle [^>]*name="([^"]+)"/g)].map((m) => m[1]);
  const slicerStyles = [...styles.matchAll(/<x14:slicerStyle [^>]*name="([^"]+)"/g)].map((m) => m[1]);
  let maxDxf = -1, where = ""; const usedTableStyles = new Set(), usedSlicerStyles = new Set();
  for (const [name, e] of z) {
    if (!/^xl\/(pivotTables|tables|slicers|worksheets)\/[^/]+\.xml$/.test(name)) continue;
    const x = await e.async("string");
    if (/pivotTables|tables/.test(name)) for (const m of x.matchAll(/(?:dxfId|headerRowDxfId|dataDxfId|totalsRowDxfId)="(\d+)"/g)) if (Number(m[1]) > maxDxf) { maxDxf = Number(m[1]); where = name; }
    if (/worksheets/.test(name)) for (const m of x.matchAll(/<cfRule [^>]*dxfId="(\d+)"/g)) if (Number(m[1]) > maxDxf) { maxDxf = Number(m[1]); where = name; }
    for (const m of x.matchAll(/<(?:pivotTableStyleInfo|tableStyleInfo) [^>]*name="([^"]+)"/g)) usedTableStyles.add(m[1]);
    if (/slicers/.test(name)) for (const m of x.matchAll(/<slicer [^>]*style="([^"]+)"/g)) usedSlicerStyles.add(m[1]);
  }
  const builtin = (n) => /^(TableStyle|PivotStyle)(Light|Medium|Dark)\d+$/.test(n) || /^SlicerStyle(Light|Dark|Other)\d+$/.test(n);
  console.log(`${f.split("/").pop()}: dxfs ${dxfs}; highest dxfId referred to by pivot tables / tables / rules: ${maxDxf} (${where}); table styles defined [${tableStyles.join(", ")}], used [${[...usedTableStyles].join(", ")}] -> missing [${[...usedTableStyles].filter((n) => !builtin(n) && !tableStyles.includes(n)).join(", ")}]; slicer styles defined ${slicerStyles.length}, used [${[...usedSlicerStyles].join(", ")}] -> missing [${[...usedSlicerStyles].filter((n) => !builtin(n) && !slicerStyles.includes(n)).join(", ")}]`);
}
