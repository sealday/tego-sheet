import { describe, expect, it } from 'vitest';
import { createSheetGridModel, createViewportMetrics } from '../../../src/engine';
import { projectSheetObjectsToViewport } from '../../../src/react/adapters/object-adapter';

describe('React screen object adapter', () => {
  it('maps worksheet anchors through viewport headers and scrolling before shared projection', () => {
    const model = createSheetGridModel({ rows: { len: 20 }, cols: { len: 10 } });
    const viewport = createViewportMetrics(model, {
      width: 320,
      height: 180,
      rowHeaderWidth: 40,
      columnHeaderHeight: 25,
      scroll: { x: 80, y: 20 },
      freeze: { row: 1, column: 1 },
    });
    const projections = projectSheetObjectsToViewport(
      [
        {
          id: 'cell-object',
          kind: 'shape',
          anchor: {
            type: 'one-cell',
            cell: { sheetId: 'sheet', row: 2, column: 2 },
            offset: { x: 5, y: 3 },
            size: { width: 30, height: 20 },
          },
          zIndex: 1,
          locked: false,
          templateRepeat: 'shared',
          shape: 'rectangle',
          style: { fill: '#123456' },
          accessibility: { name: 'Cell object' },
        },
      ] as never,
      [],
      viewport,
    );

    expect(projections[0]).toMatchObject({
      bounds: { x: 165, y: 58, width: 30, height: 20 },
      commands: [
        {
          kind: 'fill-rect',
          rect: { x: 165, y: 58, width: 30, height: 20 },
        },
        {
          kind: 'stroke-rect',
          rect: { x: 165, y: 58, width: 30, height: 20 },
        },
      ],
    });
  });
});
