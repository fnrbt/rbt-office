import { defineConfig } from 'vite';

// Allow /@fs access to fixtures and Windows fonts anywhere under the repo (and
// its sibling worktrees), mirroring tools/browser-runner's setup.
export default defineConfig({
  server: {
    fs: { allow: ['..', '../..', '../../..', '../../../..'] },
  },
  optimizeDeps: {
    include: ['@eigenpal/docx-editor-react', '@eigenpal/docx-editor-core'],
  },
});
