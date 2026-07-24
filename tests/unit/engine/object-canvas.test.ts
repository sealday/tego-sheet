import { describe, expect, it, vi } from 'vitest';
import { CanvasEngine, createSheetGridModel, createViewportMetrics } from '../../../src/engine';
import type { ScreenObjectProjection } from '../../../src/objects';
import { createCanvasHarness } from '../../helpers/canvas-harness';

describe('Canvas object layer', () => {
  it('paints shared display commands in projection order with rotation and selection', () => {
    const sheet = { rows: { len: 2 }, cols: { len: 2 } };
    const viewport = createViewportMetrics(createSheetGridModel(sheet), {
      width: 320,
      height: 180,
    });
    const harness = createCanvasHarness();
    const onRenderError = vi.fn();
    const engine = new CanvasEngine(harness.canvas, {
      animationFrame: harness.animationFrame,
      measurement: harness.measurement,
      onRenderError,
    });
    const objects = [
      {
        object: { id: 'back' },
        bounds: { x: 45, y: 30, width: 40, height: 30 },
        commands: [
          {
            kind: 'fill-rect',
            rect: { x: 45, y: 30, width: 40, height: 30 },
            color: '#123456',
          },
        ],
        diagnostics: [],
      },
      {
        object: { id: 'front' },
        bounds: { x: 90, y: 50, width: 40, height: 30 },
        commands: [
          {
            kind: 'group',
            rotation: 45,
            origin: { x: 110, y: 65 },
            commands: [
              {
                kind: 'stroke-rect',
                rect: { x: 90, y: 50, width: 40, height: 30 },
                color: '#654321',
                width: 2,
              },
            ],
          },
        ],
        diagnostics: [],
      },
    ] as unknown as readonly ScreenObjectProjection[];

    engine.render({ sheet, viewport, objects, selectedObjectId: 'front' });
    harness.animationFrame.flush();

    expect(onRenderError).not.toHaveBeenCalled();
    const firstObjectFill = harness.operations.findIndex(
      ({ name, args }) => name === 'set:fillStyle' && args[0] === '#123456',
    );
    const frontStroke = harness.operations.findIndex(
      ({ name, args }) => name === 'set:strokeStyle' && args[0] === '#654321',
    );
    expect(firstObjectFill).toBeGreaterThan(-1);
    expect(frontStroke).toBeGreaterThan(firstObjectFill);
    expect(harness.operations).toContainEqual({
      name: 'rotate',
      args: [Math.PI / 4],
    });
    expect(harness.operations).toContainEqual({
      name: 'set:strokeStyle',
      args: ['#2563eb'],
    });
  });
});
