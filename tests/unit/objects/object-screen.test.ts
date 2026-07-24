import { describe, expect, it } from 'vitest';
import { projectObjectsToScreen, transformObjectByKeyboard } from '../../../src/objects';

const geometry = {
  rowOffset: (row: number) => row * 20,
  columnOffset: (column: number) => column * 80,
};

const common = {
  anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 40, height: 30 } },
  locked: false,
  templateRepeat: 'shared',
  accessibility: { name: 'Object' },
} as const;

describe('screen object projection', () => {
  it('uses the shared display command geometry, culls outside the viewport, and sorts paint order', () => {
    const projections = projectObjectsToScreen(
      [
        {
          ...common,
          id: 'front',
          kind: 'text-box',
          zIndex: 2,
          text: 'Front',
          rotation: 90,
          style: { color: '#111111', fontFamily: 'Arial', fontSize: 12 },
        },
        {
          ...common,
          id: 'outside',
          kind: 'shape',
          zIndex: 0,
          anchor: { type: 'absolute', rect: { x: 500, y: 500, width: 20, height: 20 } },
          shape: 'rectangle',
          style: { fill: '#ff0000' },
        },
        {
          ...common,
          id: 'back',
          kind: 'shape',
          zIndex: 1,
          shape: 'rectangle',
          style: { fill: '#ffeecc', stroke: '#112233', strokeWidth: 2 },
        },
      ] as never,
      {
        geometry,
        resources: [],
        viewport: { x: 0, y: 0, width: 100, height: 100 },
      },
    );

    expect(projections.map(({ object }) => object.id)).toEqual(['back', 'front']);
    expect(projections[0]).toMatchObject({
      bounds: { x: 10, y: 20, width: 40, height: 30 },
      commands: [
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
    });
    expect(projections[1]?.commands).toEqual([
      {
        kind: 'group',
        rotation: 90,
        origin: { x: 30, y: 35 },
        commands: [
          {
            kind: 'clip',
            rect: { x: 10, y: 20, width: 40, height: 30 },
            commands: [expect.objectContaining({ kind: 'text', text: 'Front' })],
          },
        ],
      },
    ]);
    expect(projections[1]?.bounds).toEqual({ x: 15, y: 15, width: 30, height: 40 });
  });

  it('emits a real image command only for a resolved decoded image', () => {
    const decoded = {};
    const [projection] = projectObjectsToScreen(
      [
        {
          id: 'logo',
          kind: 'image',
          anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 40, height: 30 } },
          zIndex: 1,
          locked: false,
          templateRepeat: 'shared',
          resourceId: 'logo-resource',
          fit: 'contain',
          accessibility: { name: 'Logo' },
        },
      ] as never,
      {
        geometry,
        resources: [
          {
            id: 'logo-resource',
            kind: 'image',
            mimeType: 'image/png',
            url: 'data:image/png;base64,AAAA',
          },
        ] as never,
        resolvedResources: {
          'logo-resource': {
            type: 'image',
            mimeType: 'image/png',
            width: 1,
            height: 1,
            decoded,
          },
        },
        viewport: { x: 0, y: 0, width: 100, height: 100 },
      },
    );

    expect(projection?.commands).toEqual([
      {
        kind: 'image',
        resourceId: 'logo-resource',
        rect: { x: 10, y: 20, width: 40, height: 30 },
        fit: 'contain',
      },
    ]);
    expect(projection?.imageResources).toEqual({
      'logo-resource': { source: decoded, width: 1, height: 1 },
    });
    expect(projection?.diagnostics).toEqual([]);
  });

  it('keeps rotated objects whose rotated bounds intersect the viewport', () => {
    const [projection] = projectObjectsToScreen(
      [
        {
          ...common,
          id: 'rotated-edge',
          kind: 'shape',
          zIndex: 1,
          anchor: { type: 'absolute', rect: { x: 101, y: 40, width: 20, height: 80 } },
          rotation: 45,
          shape: 'rectangle',
          style: { fill: '#ffeecc' },
        },
      ] as never,
      {
        geometry,
        resources: [],
        viewport: { x: 0, y: 0, width: 100, height: 100 },
      },
    );

    expect(projection?.object.id).toBe('rotated-edge');
    expect(projection?.bounds.x).toBeLessThan(100);
  });

  it.each([
    ['missing', []],
    [
      'dangerous',
      [
        {
          id: 'image',
          kind: 'image',
          mimeType: 'image/png',
          url: 'javascript:alert(1)',
        },
      ],
    ],
  ])('replaces a %s image resource with a safe placeholder and diagnostic', (_case, resources) => {
    const [projection] = projectObjectsToScreen(
      [
        {
          ...common,
          id: 'picture',
          kind: 'image',
          zIndex: 1,
          resourceId: 'image',
        },
      ] as never,
      {
        geometry,
        resources: resources as never,
        viewport: { x: 0, y: 0, width: 100, height: 100 },
      },
    );

    expect(projection?.diagnostics).toEqual([
      expect.objectContaining({
        code: 'OBJECT_RESOURCE_MISSING',
        objectId: 'picture',
        resourceId: 'image',
      }),
    ]);
    expect(projection?.commands).toEqual([
      { kind: 'fill-rect', rect: common.anchor.rect, color: '#f3f4f6' },
      { kind: 'stroke-rect', rect: common.anchor.rect, color: '#9ca3af', width: 1 },
      expect.objectContaining({ kind: 'text', text: 'Image unavailable' }),
    ]);
  });

  it('keeps safe but unresolved image metadata as a diagnosed placeholder', () => {
    const [projection] = projectObjectsToScreen(
      [
        {
          ...common,
          id: 'picture',
          kind: 'image',
          zIndex: 1,
          resourceId: 'image',
        },
      ] as never,
      {
        geometry,
        resources: [
          {
            id: 'image',
            kind: 'image',
            mimeType: 'image/png',
            url: 'https://example.test/image.png',
          },
        ] as never,
        viewport: { x: 0, y: 0, width: 100, height: 100 },
      },
    );

    expect(projection?.diagnostics).toEqual([
      expect.objectContaining({
        code: 'OBJECT_RESOURCE_MISSING',
        objectId: 'picture',
        resourceId: 'image',
      }),
    ]);
    expect(projection?.commands).toEqual([
      { kind: 'fill-rect', rect: common.anchor.rect, color: '#f3f4f6' },
      { kind: 'stroke-rect', rect: common.anchor.rect, color: '#9ca3af', width: 1 },
      expect.objectContaining({ kind: 'text', text: 'Image unavailable' }),
    ]);
  });
});

describe('keyboard object transforms', () => {
  const object = {
    ...common,
    id: 'shape',
    kind: 'shape',
    zIndex: 1,
    shape: 'rectangle',
    rotation: 359,
    style: { fill: '#ffeecc' },
  } as const;

  it('moves, resizes, and rotates absolute objects without mutating the input', () => {
    expect(transformObjectByKeyboard(object as never, { type: 'move', x: 3, y: -2 })).toMatchObject(
      {
        anchor: { rect: { x: 13, y: 18, width: 40, height: 30 } },
      },
    );
    expect(
      transformObjectByKeyboard(object as never, { type: 'resize', x: -50, y: 4 }),
    ).toMatchObject({
      anchor: { rect: { x: 10, y: 20, width: 1, height: 34 } },
    });
    expect(
      transformObjectByKeyboard(object as never, { type: 'rotate', degrees: 2 }),
    ).toMatchObject({
      rotation: 1,
    });
    expect(object.anchor.rect).toEqual({ x: 10, y: 20, width: 40, height: 30 });
  });

  it('allows locked objects to be selected but refuses every transform', () => {
    expect(
      transformObjectByKeyboard({ ...object, locked: true } as never, {
        type: 'move',
        x: 1,
        y: 0,
      }),
    ).toBeUndefined();
  });
});
