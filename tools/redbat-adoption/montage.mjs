// usage: node montage.mjs out.png a.png b.png c.png d.png  (2 x 2, each scaled to 683 x 384)
import { chromium } from "./lib.mjs";
import fs from "node:fs";
const [out, ...files] = process.argv.slice(2);
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1366, height: 768 } });
await page.setContent(`<body style="margin:0;display:grid;grid-template-columns:683px 683px">${files.map((f) => `<img src="data:image/png;base64,${fs.readFileSync(f).toString("base64")}" style="width:683px;height:384px">`).join("")}</body>`);
await page.screenshot({ path: out });
await browser.close();
