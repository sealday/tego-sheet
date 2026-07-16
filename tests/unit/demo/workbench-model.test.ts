import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  PREVIEW_EVENT_LIMIT,
  appendPreviewEvent,
  cloneExampleWorkbook,
  formatWorkbookJson,
  parseWorkbookJson,
  type PreviewEvent,
} from '../../../demo/src/workbench-model';

interface MutableTestRow {
  cells?: Record<number, { text?: string }>;
}

describe('fullscreen demo workbench model', () => {
  it('depends only on the public tego-sheet surface', () => {
    const modelSource = readFileSync(
      fileURLToPath(new URL('../../../demo/src/workbench-model.ts', import.meta.url)),
      'utf8',
    );

    expect(modelSource).not.toMatch(/from\s+['"]\.\.\/\.\.\/src(?:\/|['"])/);
  });

  it('round-trips the legacy workbook through readable deterministic JSON', () => {
    const workbook = cloneExampleWorkbook();
    const formatted = formatWorkbookJson(workbook);

    expect(formatted).toBe(formatWorkbookJson(workbook));
    expect(formatted).toContain('\n  {\n    "name": "Budget"');
    expect(parseWorkbookJson(formatted)).toEqual(workbook);
  });

  it('accepts and formats a single-sheet object without changing its JSON shape', () => {
    const source = '{"name":"Solo","rows":{"len":1}}';
    const parsed = parseWorkbookJson(source);

    expect(parsed).toEqual({ name: 'Solo', rows: { len: 1 } });
    expect(formatWorkbookJson(parsed)).toBe('{\n  "name": "Solo",\n  "rows": {\n    "len": 1\n  }\n}');
  });

  it('returns deeply independent example workbooks', () => {
    const first = cloneExampleWorkbook();
    const second = cloneExampleWorkbook();
    const firstSheet = first[0];
    const firstRow = firstSheet?.rows?.[1] as MutableTestRow | undefined;
    const firstCell = firstRow?.cells?.[0];
    const secondRow = second[0]?.rows?.[1] as MutableTestRow | undefined;

    expect(firstSheet).toBeDefined();
    expect(firstCell).toBeDefined();
    Reflect.set(firstSheet!, 'name', 'Changed');
    Reflect.set(firstCell!, 'text', 'Changed item');

    expect(second[0]?.name).toBe('Budget');
    expect(secondRow?.cells?.[0]?.text).toBe('Hosting');
  });

  it('rejects invalid JSON without changing an existing workbook', () => {
    const existing = cloneExampleWorkbook();
    const snapshot = formatWorkbookJson(existing);

    expect(() => parseWorkbookJson('{broken')).toThrow(SyntaxError);
    expect(formatWorkbookJson(existing)).toBe(snapshot);
  });

  it.each([
    ['null', 'null'],
    ['a number', '42'],
    ['a string', '"sheet"'],
    ['an array containing a primitive', '[1]'],
  ])('rejects %s as invalid top-level workbook data', (_description, source) => {
    expect(() => parseWorkbookJson(source)).toThrow('Workbook JSON must be a sheet object or an array of sheet objects.');
  });

  it.each([
    ['a numeric sheet name', '[{"name":42}]'],
    ['array-shaped rows', '[{"rows":[]}]'],
    ['array-shaped columns', '[{"cols":[]}]'],
    ['a non-string freeze point', '[{"freeze":false}]'],
    ['a numeric name in a single-sheet object', '{"name":42}'],
  ])('rejects %s through known-field validation', (_description, source) => {
    expect(() => parseWorkbookJson(source)).toThrow('Workbook data is invalid');
  });

  it('keeps useful event entries newest-first within the configured limit', () => {
    const entries = Array.from({ length: PREVIEW_EVENT_LIMIT + 3 }, (_, index) => index)
      .reduce((logs, index) => appendPreviewEvent(logs, {
        timestamp: `2026-07-16T00:00:${String(index).padStart(2, '0')}.000Z`,
        label: `event-${index}`,
        details: `row ${index}`,
      }), [] as ReturnType<typeof appendPreviewEvent>);

    expect(entries).toHaveLength(PREVIEW_EVENT_LIMIT);
    expect(entries[0]).toMatchObject({
      timestamp: '2026-07-16T00:00:14.000Z',
      label: 'event-14',
      details: 'row 14',
    });
    expect(entries[0]?.id).toBeTruthy();
    expect(entries.at(-1)?.label).toBe('event-3');
    expect(new Set(entries.map(entry => entry.id))).toHaveLength(PREVIEW_EVENT_LIMIT);
  });

  it('keeps an event id stable when newer entries are appended', () => {
    const first = appendPreviewEvent([], {
      timestamp: '2026-07-16T00:00:00.000Z',
      label: 'selection changed',
    });
    const firstId = first[0]?.id;
    const second = appendPreviewEvent(first, {
      timestamp: '2026-07-16T00:00:01.000Z',
      label: 'cell changed',
    });

    expect(second[1]?.id).toBe(firstId);
  });

  it('protects generated event ids and resolves repeated-timestamp collisions', () => {
    const structurallyAssignableInput: PreviewEvent = {
      id: 'caller-controlled-id',
      timestamp: '2026-07-16T00:00:00.000Z',
      label: 'selection changed',
    };
    const first = appendPreviewEvent([], structurallyAssignableInput);
    const second = appendPreviewEvent(first, {
      ...structurallyAssignableInput,
      label: 'cell changed',
    });

    expect(second.map(entry => entry.id)).not.toContain('caller-controlled-id');
    expect(new Set(second.map(entry => entry.id))).toHaveLength(2);
  });
});
