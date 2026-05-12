import { defineConfig } from 'vite';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const officeRoot = path.resolve(root, '../..');

export default defineConfig({
  server: {
    fs: {
      allow: [root, officeRoot]
    }
  }
});
