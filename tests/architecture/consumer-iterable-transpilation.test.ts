import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { transformSync } = require('@babel/core') as {
  transformSync(
    source: string,
    options: Readonly<Record<string, unknown>>,
  ): Readonly<{ code?: string | null }> | null;
};
const docusaurusPreset = require.resolve('@docusaurus/babel/preset');

function read(relativePath: string): string {
  return readFileSync(join(process.cwd(), relativePath), 'utf8');
}

function expressionAfter(source: string, marker: string, terminator: string): string {
  const start = source.indexOf(marker);
  if (start < 0) throw new Error(`missing source marker: ${marker}`);
  const valueStart = start + marker.length;
  const end = source.indexOf(terminator, valueStart);
  if (end < 0) throw new Error(`missing source terminator after: ${marker}`);
  return source.slice(valueStart, end).trim();
}

function executeClientTransform<Result>(source: string): Result {
  const transformed = transformSync(source, {
    babelrc: false,
    caller: { name: 'client' },
    configFile: false,
    filename: 'consumer-iterable-fixture.ts',
    presets: [docusaurusPreset],
  })?.code;
  if (typeof transformed !== 'string') {
    throw new Error('Docusaurus Babel transform produced no code');
  }
  const context: { result?: Result } = {};
  runInNewContext(transformed, context);
  if (!('result' in context)) throw new Error('transformed fixture produced no result');
  return context.result as Result;
}

describe('Docusaurus client iterable transpilation', () => {
  it('preserves numeric grid boundaries from the real painter function', () => {
    const source = read('src/engine/canvas/grid-painter.ts');
    const start = source.indexOf('function boundaries(');
    const end = source.indexOf('\n\nexport function paintGrid', start);
    if (start < 0 || end < 0) throw new Error('grid boundary function not found');
    const boundaryFunction = source.slice(start, end);
    const expression = expressionAfter(boundaryFunction, '  return ', ';');

    const result = executeClientTransform<unknown>(`const values = new Set([0, 10, 20]);
globalThis.result = ${expression};`);

    expect(result).toEqual([0, 10, 20]);
  });

  it('preserves selected filter strings from the real apply expression', () => {
    const source = read('src/ui/menus/filter-menu.tsx');
    const expression = expressionAfter(source, '            value: ', ',\n');

    const result = executeClientTransform<unknown>(`const selected = new Set(['Keyboard', 'Mouse']);
globalThis.result = ${expression};`);

    expect(result).toEqual(['Keyboard', 'Mouse']);
  });

  it('preserves disabled toolbar actions from the real readonly-set materialization', () => {
    const source = read('src/react/tego-sheet.tsx');
    const expression = expressionAfter(source, '  const values = ', ';');

    const result = executeClientTransform<unknown>(`const source = new Set(['undo', 'redo']);
const values = ${expression};
globalThis.result = { size: values.length, values };`);

    expect(result).toEqual({ size: 2, values: ['undo', 'redo'] });
  });

  it('preserves listener entry snapshots from the real subscription expression', () => {
    const source = read('src/core/controller/subscription-store.ts');
    const expression = expressionAfter(source, '        const current = ', ';').replace(
      'this.listeners',
      'listeners',
    );

    const result =
      executeClientTransform<unknown>(`const listeners = new Map([[1, 'first'], [2, 'second']]);
globalThis.result = ${expression};`);

    expect(result).toEqual([
      [1, 'first'],
      [2, 'second'],
    ]);
  });
});
