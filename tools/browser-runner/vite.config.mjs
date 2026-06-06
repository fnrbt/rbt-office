import { defineConfig } from 'vite';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const officeRoot = path.resolve(root, '../..');
const browserRoot = path.resolve(officeRoot, '..'); // sibling ../text (ucd.dat, fonts)

export default defineConfig({
  server: {
    fs: {
      allow: [root, officeRoot, browserRoot]
    }
  }
});
