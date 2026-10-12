// usage: node zoom.mjs in.png out.png x y w h scale
import { chromium } from "./lib.mjs";
import fs from "node:fs";
const [inp, out, x, y, w, h, k] = process.argv.slice(2);
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: Math.round(w * k), height: Math.round(h * k) } });
const b64 = fs.readFileSync(inp).toString("base64");
await page.setContent(`<body style="margin:0;overflow:hidden"><img src="data:image/png;base64,${b64}" style="image-rendering:pixelated;transform-origin:0 0;transform:scale(${k}) translate(${-x}px,${-y}px)"></body>`);
await page.screenshot({ path: out });
await browser.close();
