import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const officeRoot = path.resolve(root, '../..');
const output = path.resolve(process.env.SHELL_RUNNER_OUTDIR ?? path.join(root, 'src'));
const entry = `/@fs/${output.replaceAll(path.sep, '/')}/Program.js`;

export default {
  // Fable's output can live in a private /tmp directory, avoiding collisions
  // with builds in other sessions; Vite serves that module and its imports.
  plugins: [{
    name: 'shell-runner-entry',
    transformIndexHtml: {
      order: 'pre',
      handler(html) {
        return html.replace('/src/Program.js', entry);
      }
    }
  }],
  server: {
    fs: {
      allow: [root, officeRoot, output]
    }
  }
};
