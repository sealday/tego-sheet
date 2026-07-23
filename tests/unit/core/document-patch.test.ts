import { describe, expect, it } from 'vitest';
import {
  applyDocumentPatch,
  createDocumentPatch,
} from '../../../src/core/controller/document-patch';

describe('internal document patches', () => {
  it('stores only changed paths and round-trips through an inverse patch', () => {
    const before = {
      workbook: {
        sheets: [{ id: 'sheet-1', cells: [{ row: 0, text: 'before' }], name: 'Data' }],
        settings: { locale: 'en' },
      },
      untouched: { nested: true },
    };
    const after = {
      ...before,
      workbook: {
        ...before.workbook,
        sheets: [{ id: 'sheet-1', cells: [{ row: 0, text: 'after' }], name: 'Data' }],
      },
    };

    const forward = createDocumentPatch(before, after);
    const inverse = createDocumentPatch(after, before);

    expect(forward.operations).toEqual([
      {
        op: 'set',
        path: ['workbook', 'sheets', 0, 'cells', 0, 'text'],
        value: 'after',
      },
    ]);
    expect(forward.operations.some((operation) => operation.path.length === 0)).toBe(false);
    expect(applyDocumentPatch(before, forward)).toEqual(after);
    expect(applyDocumentPatch(after, inverse)).toEqual(before);
    expect(Object.isFrozen(forward)).toBe(true);
    expect(Object.isFrozen(forward.operations[0]?.path)).toBe(true);
  });

  it('uses one collection patch for structural array length changes without storing the root', () => {
    const before = { sheets: [{ cells: [] }], metadata: { kept: true } };
    const after = {
      sheets: [{ cells: [{ row: 0, column: 0, text: 'added' }] }],
      metadata: before.metadata,
    };
    const patch = createDocumentPatch(before, after);

    expect(patch.operations).toEqual([
      {
        op: 'splice',
        path: ['sheets', 0, 'cells'],
        index: 0,
        deleteCount: 0,
        values: [{ row: 0, column: 0, text: 'added' }],
      },
    ]);
    expect(applyDocumentPatch(before, patch)).toEqual(after);
  });

  it('applies dangerous object keys as data without changing prototypes', () => {
    const before = { extensions: {} };
    const after = JSON.parse('{"extensions":{"__proto__":{"safe":true}}}') as typeof before;
    const patched = applyDocumentPatch(before, createDocumentPatch(before, after));

    expect(Object.hasOwn(patched.extensions, '__proto__')).toBe(true);
    expect(({} as { safe?: boolean }).safe).toBeUndefined();
  });

  it('applies a root array splice for workbook sheet additions and deletions', () => {
    const before = [{ name: 'A' }];
    const after = [{ name: 'A' }, { name: 'B' }];

    expect(applyDocumentPatch(before, createDocumentPatch(before, after))).toEqual(after);
    expect(applyDocumentPatch(after, createDocumentPatch(after, before))).toEqual(before);
  });

  it('encodes a large same-length left shift as one deletion and one necessary insertion', () => {
    const before = Array.from({ length: 1_000 }, (_, id) => ({
      id,
      value: `value-${id}`,
    }));
    const after = [...before.slice(1), { id: 1_000, value: 'value-1000' }];
    const patch = createDocumentPatch(before, after);
    const splices = patch.operations.filter((operation) => operation.op === 'splice');

    expect(splices).toHaveLength(2);
    expect(splices.reduce((total, operation) => total + operation.deleteCount, 0)).toBe(1);
    expect(splices.reduce((total, operation) => total + operation.values.length, 0)).toBe(1);
    expect(applyDocumentPatch(before, patch)).toEqual(after);
  });

  it('keeps adversarial duplicate-array shift detection linear and compact', () => {
    const before = [
      ...Array.from({ length: 20_000 }, () => ({ value: 'duplicate' })),
      { value: 'tail' },
    ];
    const after = [...before.slice(1), { value: 'replacement' }];
    const startedAt = performance.now();

    const patch = createDocumentPatch(before, after);

    expect(performance.now() - startedAt).toBeLessThan(1_000);
    expect(patch.operations).toHaveLength(2);
    expect(applyDocumentPatch(before, patch)).toEqual(after);
  });
});
