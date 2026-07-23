import { describe, expect, it } from 'vitest';
import { paginateTemplateTargets } from '../../../src/print/layout';

describe('deterministic template pagination', () => {
  it('starts every non-contiguous target on a new page in declaration order', () => {
    const result = paginateTemplateTargets({
      targets: [
        { id: 'first', rows: [30, 30], width: 100 },
        { id: 'second', rows: [20], width: 100 },
      ],
      paper: { width: 140, height: 100 },
      margins: { top: 10, right: 10, bottom: 10, left: 10 },
      scale: { type: 'fixed', value: 1 },
      manualBreaks: [],
      maxPages: 10,
    });
    expect(result.pages.map(({ targetId }) => targetId)).toEqual(['first', 'second']);
    expect(result.pages.map(({ id }) => id)).toEqual(['first:0', 'second:0']);
  });

  it('honors manual breaks and terminates oversized rows with a diagnostic', () => {
    const result = paginateTemplateTargets({
      targets: [{ id: 'target', rows: [20, 20, 90], width: 100 }],
      paper: { width: 140, height: 100 },
      margins: { top: 10, right: 10, bottom: 10, left: 10 },
      scale: { type: 'fixed', value: 1 },
      manualBreaks: [{ targetId: 'target', beforeRow: 1 }],
      maxPages: 10,
    });
    expect(result.pages).toHaveLength(2);
    expect(result.diagnostics).toEqual([expect.objectContaining({ code: 'ROW_EXCEEDS_PAGE' })]);
  });

  it('enforces the page limit before returning partial pages', () => {
    const result = paginateTemplateTargets({
      targets: [{ id: 'target', rows: [60, 60, 60], width: 100 }],
      paper: { width: 140, height: 100 },
      margins: { top: 10, right: 10, bottom: 10, left: 10 },
      scale: { type: 'fixed', value: 1 },
      manualBreaks: [],
      maxPages: 2,
    });
    expect(result.pages).toEqual([]);
    expect(result.diagnostics).toEqual([expect.objectContaining({ code: 'PAGE_LIMIT_EXCEEDED' })]);
  });
});
