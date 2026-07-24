import { describe, expect, it } from 'vitest';
import { createFontMetrics } from '../../../src/presentation';
import { createPrintDisplayList, validatePrintDisplayCommands } from '../../../src/print';

describe('renderer-neutral display command groups', () => {
  it('validates nested commands and deeply freezes grouped transforms', () => {
    const command = {
      kind: 'group',
      rotation: 90,
      origin: { x: 30, y: 35 },
      commands: [
        {
          kind: 'clip',
          rect: { x: 10, y: 20, width: 40, height: 30 },
          commands: [{ kind: 'script', source: 'alert(1)' }],
        },
      ],
    } as const;
    expect(validatePrintDisplayCommands([command])).toEqual([
      expect.objectContaining({ code: 'DRAW_COMMAND_UNSUPPORTED' }),
    ]);

    const safe = {
      ...command,
      commands: [
        {
          kind: 'clip',
          rect: { x: 10, y: 20, width: 40, height: 30 },
          commands: [
            {
              kind: 'fill-rect',
              rect: { x: 10, y: 20, width: 40, height: 30 },
              color: '#ffffff',
            },
          ],
        },
      ],
    } as const;
    const display = createPrintDisplayList({
      pages: [{ width: 100, height: 100, cells: [], overlays: [safe] }],
      fontMetrics: createFontMetrics({
        fonts: {},
        fallbackFont: 'Arial',
        fallback: { averageAdvance: 6, lineHeight: 12 },
      }),
    });
    const frozen = display.pages[0]!.commands[0]!;

    expect(validatePrintDisplayCommands([safe])).toEqual([]);
    expect(Object.isFrozen(frozen)).toBe(true);
    expect(frozen.kind).toBe('group');
    if (frozen.kind !== 'group') throw new Error('expected group command');
    expect(Object.isFrozen(frozen.origin)).toBe(true);
    expect(Object.isFrozen(frozen.commands)).toBe(true);
    expect(Object.isFrozen(frozen.commands[0])).toBe(true);
  });
});
