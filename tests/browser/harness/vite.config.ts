import { resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

const projectRoot = resolve(import.meta.dirname, '../../..');

export default defineConfig({
  root: import.meta.dirname,
  plugins: [react()],
  resolve: {
    alias: [
      {
        find: 'tego-sheet/output/pdf',
        replacement: resolve(projectRoot, 'src/output/pdf/index.ts'),
      },
      {
        find: 'tego-sheet/output/xlsx',
        replacement: resolve(projectRoot, 'src/output/xlsx/index.ts'),
      },
      {
        find: 'tego-sheet/styles.css',
        replacement: resolve(projectRoot, 'src/ui/styles/index.less'),
      },
      { find: 'tego-sheet', replacement: resolve(projectRoot, 'src/index.ts') },
    ],
  },
  optimizeDeps: {
    include: ['pdfkit/js/pdfkit.standalone.js'],
  },
  server: {
    host: '127.0.0.1',
    port: 4173,
    strictPort: true,
  },
});
