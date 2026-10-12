// TEMPORARY - not committed: the production build with the office harness page.
import { mergeConfig } from "vite";
import path from "node:path";
import base from "./vite.config";
export default mergeConfig(base, { build: { rollupOptions: { input: { index: path.resolve(__dirname, "index.html"), harness: path.resolve(__dirname, "office-harness.html") } } } });
