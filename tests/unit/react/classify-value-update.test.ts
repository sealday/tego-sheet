import { describe, expect, it } from 'vitest';
import { classifyValueUpdate } from '../../../src/react/control/classify-value-update';

describe('classifyValueUpdate', () => {
  it('preserves the canonicalization exception and its original cause for invalid values', () => {
    const update = classifyValueUpdate(
      {
        observedValue: null,
        acknowledgedKey: '[]',
        pending: [],
      },
      { schemaVersion: 2, workbook: null } as never,
    );

    expect(update.kind).toBe('invalid');
    if (update.kind !== 'invalid') return;
    expect(update.error).toMatchObject({
      code: 'INVALID_DATA',
      recoverable: true,
      cause: expect.any(Array),
    });
  });
});
