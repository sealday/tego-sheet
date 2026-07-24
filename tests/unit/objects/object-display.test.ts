import { describe, expect, it } from 'vitest';
import { objectToDisplayCommands } from '../../../src/objects';

const context = {
  resources: { image: new Uint8Array([1]) },
  geometry: {
    rowOffset: (row: number) => row * 20,
    columnOffset: (column: number) => column * 80,
  },
};
const common = {
  id: 'object',
  anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 40, height: 30 } },
  zIndex: 1,
  locked: false,
  templateRepeat: 'shared',
  accessibility: { name: 'Object' },
} as const;

describe('renderer-neutral persistent object commands', () => {
  it.each([
    [
      'rectangle',
      [
        {
          kind: 'fill-rect',
          rect: { x: 10, y: 20, width: 40, height: 30 },
          color: '#ffeecc',
        },
        {
          kind: 'stroke-rect',
          rect: { x: 10, y: 20, width: 40, height: 30 },
          color: '#112233',
          width: 2,
        },
      ],
    ],
    [
      'ellipse',
      [
        {
          kind: 'path',
          data: 'M 30 20 A 20 15 0 1 0 30 50 A 20 15 0 1 0 30 20 Z',
          fill: '#ffeecc',
          stroke: '#112233',
          width: 2,
        },
      ],
    ],
    [
      'line',
      [
        {
          kind: 'line',
          x1: 10,
          y1: 20,
          x2: 50,
          y2: 50,
          color: '#112233',
          width: 2,
        },
      ],
    ],
  ])('emits exact %s geometry inside one rotation group', (shape, commands) => {
    expect(
      objectToDisplayCommands(
        {
          ...common,
          kind: 'shape',
          shape,
          rotation: 90,
          style: { fill: '#ffeecc', stroke: '#112233', strokeWidth: 2 },
        } as never,
        context,
      ),
    ).toEqual([
      {
        kind: 'group',
        rotation: 90,
        origin: { x: 30, y: 35 },
        commands,
      },
    ]);
  });

  it('clips text boxes before applying rotation and preserves unrotated compatibility', () => {
    const object = {
      ...common,
      kind: 'text-box',
      rotation: 45,
      text: 'Clipped text',
      style: { color: '#111111', fontFamily: 'Arial', fontSize: 12 },
    } as const;
    expect(objectToDisplayCommands(object as never, context)).toEqual([
      {
        kind: 'group',
        rotation: 45,
        origin: { x: 30, y: 35 },
        commands: [
          {
            kind: 'clip',
            rect: { x: 10, y: 20, width: 40, height: 30 },
            commands: [
              expect.objectContaining({
                kind: 'text',
                text: 'Clipped text',
                x: 10,
                y: 32,
                maxWidth: 40,
              }),
            ],
          },
        ],
      },
    ]);
    expect(
      objectToDisplayCommands(
        {
          ...object,
          rotation: undefined,
        } as never,
        context,
      ),
    ).toEqual([
      {
        kind: 'clip',
        rect: { x: 10, y: 20, width: 40, height: 30 },
        commands: [expect.objectContaining({ kind: 'text', text: 'Clipped text' })],
      },
    ]);
  });

  it('rotates images while keeping missing resources fail-closed', () => {
    expect(
      objectToDisplayCommands(
        {
          ...common,
          kind: 'image',
          resourceId: 'image',
          rotation: 180,
        } as never,
        context,
      ),
    ).toEqual([
      {
        kind: 'group',
        rotation: 180,
        origin: { x: 30, y: 35 },
        commands: [
          {
            kind: 'image',
            resourceId: 'image',
            rect: { x: 10, y: 20, width: 40, height: 30 },
            fit: 'contain',
          },
        ],
      },
    ]);
    expect(() =>
      objectToDisplayCommands(
        { ...common, kind: 'image', resourceId: 'missing', rotation: 90 } as never,
        context,
      ),
    ).toThrowError(expect.objectContaining({ code: 'OBJECT_RESOURCE_MISSING' }));
  });
});
