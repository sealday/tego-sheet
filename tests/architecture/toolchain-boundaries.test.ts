import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, it } from 'vitest';
import viteConfig from '../../vite.config';

it('externalizes every React and React DOM runtime subpath', () => {
  const external = viteConfig.build?.rollupOptions?.external;

  expect(external).toBeTypeOf('function');
  if (typeof external !== 'function') {
    throw new TypeError('Vite must use a React external predicate');
  }

  for (const id of [
    'react',
    'react/jsx-runtime',
    'react/jsx-dev-runtime',
    'react-dom',
    'react-dom/client',
    'react-dom/server',
  ]) {
    expect(external(id, undefined, false), id).toBe(true);
  }

  for (const id of ['reactive', 'react-domestic', '@scope/react']) {
    expect(external(id, undefined, false), id).toBe(false);
  }
});

it('suppresses only the non-diagnostic slow-plugin build advisory', () => {
  expect(viteConfig.build?.rollupOptions?.checks).toEqual({
    pluginTimings: false,
  });
});

it('builds package exports before typechecking fresh CI checkouts', () => {
  const workflow = readFileSync(
    resolve(import.meta.dirname, '../../.github/workflows/ci.yml'),
    'utf8',
  );
  const qualityJob = workflow.slice(workflow.indexOf('  quality:'), workflow.indexOf('  vitest:'));

  expect(qualityJob.indexOf('- run: npm run build')).toBeGreaterThan(-1);
  expect(qualityJob.indexOf('- run: npm run typecheck')).toBeGreaterThan(
    qualityJob.indexOf('- run: npm run build'),
  );
});

it('resolves every source-backed package subpath during clean Vitest runs', () => {
  const vitestConfig = readFileSync(resolve(import.meta.dirname, '../../vitest.config.ts'), 'utf8');

  for (const subpath of [
    'tego-sheet/analysis',
    'tego-sheet/integrations',
    'tego-sheet/interchange',
    'tego-sheet/sdk',
    'tego-sheet/output/pdf',
    'tego-sheet/output/image',
    'tego-sheet/output/xlsx',
  ]) {
    expect(vitestConfig, subpath).toContain(`find: '${subpath}'`);
  }
});

it('keeps the core contract independent of React and browser globals', () => {
  const coreFiles = [
    'src/core/index.ts',
    'src/core/types/json.ts',
    'src/core/types/workbook.ts',
    'src/core/types/coordinates.ts',
    'src/core/types/changes.ts',
    'src/core/types/validation.ts',
    'src/core/types/options.ts',
    'src/core/errors/tego-sheet-error.ts',
    'src/core/errors/tego-sheet-exception.ts',
    'src/core/coordinates/a1.ts',
    'src/core/coordinates/ranges.ts',
    'src/core/formulas/evaluator.ts',
    'src/core/formulas/functions.ts',
    'src/core/formulas/parser.ts',
    'src/core/formulas/rendered-value.ts',
    'src/core/formulas/tokenizer.ts',
    'src/core/model/styles.ts',
  ];

  for (const file of coreFiles) {
    const source = readFileSync(resolve(import.meta.dirname, '../..', file), 'utf8');

    expect(source, file).not.toMatch(/from\s+['"]react(?:\/[^'"]*)?['"]/);
    expect(source, file).not.toMatch(/\b(?:window|document|navigator)\b/);
  }
});

it('keeps pure operations independent from the controller mutation boundary', () => {
  const directory = resolve(import.meta.dirname, '../../src/core/operations');
  const operationFiles = readdirSync(directory).filter((file) => file.endsWith('.ts'));

  expect(operationFiles).toEqual(
    expect.arrayContaining(['cell.ts', 'merge.ts', 'sheet.ts', 'structure.ts', 'style.ts']),
  );
  for (const file of operationFiles) {
    const source = readFileSync(resolve(directory, file), 'utf8');
    expect(source, file).not.toMatch(/(?:workbook-controller|controller\/)/);
    expect(source, file).not.toMatch(/\b(?:window|document|navigator)\b/);
  }
});
