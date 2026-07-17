import { resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vitest/config';
import VitestParityEvidenceReporter from './scripts/reporters/vitest-parity-evidence.ts';

const repositoryRoot = import.meta.dirname;

const sharedExcludes = [
  'dist/**',
  'coverage/**',
  'node_modules/**',
  'playwright-report/**',
  'test-results/**',
  'tests/**/*.check.*',
  'tests/package/**/*.test.*',
  'tests/ssr/public-entrypoints.test.mjs',
  'tests/browser/**',
  'tests/visual/**',
];

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: [
      {
        find: 'tego-sheet/locales/en',
        replacement: resolve(repositoryRoot, 'src/locales/en.ts'),
      },
      {
        find: 'tego-sheet/locales/zh-cn',
        replacement: resolve(repositoryRoot, 'src/locales/zh-cn.ts'),
      },
      {
        find: 'tego-sheet/locales/de',
        replacement: resolve(repositoryRoot, 'src/locales/de.ts'),
      },
      {
        find: 'tego-sheet/locales/nl',
        replacement: resolve(repositoryRoot, 'src/locales/nl.ts'),
      },
      {
        find: 'tego-sheet/styles.css',
        replacement: resolve(repositoryRoot, 'src/ui/styles/index.less'),
      },
      {
        find: /^tego-sheet$/,
        replacement: resolve(repositoryRoot, 'src/index.ts'),
      },
    ],
  },
  test: {
    exclude: sharedExcludes,
    reporters: ['default', new VitestParityEvidenceReporter({ releaseOnly: true })],
    projects: [
      {
        extends: true,
        test: {
          name: 'unit',
          environment: 'node',
          exclude: sharedExcludes,
          include: ['tests/unit/**/*.test.{ts,tsx}'],
        },
      },
      {
        extends: true,
        test: {
          name: 'component',
          environment: 'jsdom',
          exclude: sharedExcludes,
          include: ['tests/component/**/*.test.{ts,tsx}'],
        },
      },
      {
        extends: true,
        test: {
          name: 'architecture',
          environment: 'node',
          exclude: sharedExcludes,
          include: ['tests/architecture/**/*.test.{ts,tsx}'],
        },
      },
      {
        extends: true,
        test: {
          name: 'ssr',
          environment: 'node',
          exclude: sharedExcludes,
          include: ['tests/ssr/**/*.test.{ts,tsx}'],
        },
      },
      {
        extends: true,
        test: {
          name: 'parity',
          environment: 'node',
          exclude: sharedExcludes,
          include: ['tests/parity/**/*.test.ts'],
        },
      },
    ],
  },
});
