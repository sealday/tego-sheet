import { describe, expect, it } from 'vitest';
import {
  parseSpreadsheetDocument,
  serializeSpreadsheetDocument,
  type SpreadsheetDocumentInput,
} from '../../../src/document';

function documentWithObjects(objects: readonly unknown[]): SpreadsheetDocumentInput {
  return {
    schemaVersion: 2,
    id: 'object-shapes',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Objects',
          cells: [],
          merges: [],
          objects: objects as never,
        },
      ],
      styles: [],
      validations: [],
      settings: { dateSystem: 'excel-1900' },
    },
    templates: [],
    resources: {
      items: [{ id: 'image-1', kind: 'image', mimeType: 'image/png' }],
    },
    extensions: {},
  };
}

const common = {
  anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 80, height: 40 } },
  zIndex: 1,
  locked: false,
  templateRepeat: 'shared',
  accessibility: { name: 'Object' },
} as const;

describe('persistent shape objects and rotation', () => {
  it('round-trips bounded shapes and normalizes finite rotations', () => {
    const parsed = parseSpreadsheetDocument(
      documentWithObjects([
        {
          ...common,
          id: 'shape-1',
          kind: 'shape',
          shape: 'ellipse',
          rotation: 450,
          style: { fill: '#ffeecc', stroke: '#112233', strokeWidth: 2 },
        },
        {
          ...common,
          id: 'image-object',
          kind: 'image',
          resourceId: 'image-1',
          rotation: -90,
        },
        {
          ...common,
          id: 'text-object',
          kind: 'text-box',
          text: 'Backward compatible',
          style: { color: '#111111', fontFamily: 'Arial', fontSize: 12 },
        },
      ]),
    );

    if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
    const objects = parsed.document.workbook.sheets[0]!.objects;
    expect(objects.find(({ id }) => id === 'shape-1')).toMatchObject({
      kind: 'shape',
      shape: 'ellipse',
      rotation: 90,
    });
    expect(objects.find(({ id }) => id === 'image-object')).toMatchObject({
      kind: 'image',
      rotation: 270,
    });
    expect(objects.find(({ id }) => id === 'text-object')).not.toHaveProperty('rotation');
    expect(
      parseSpreadsheetDocument(JSON.parse(serializeSpreadsheetDocument(parsed.document))),
    ).toMatchObject({
      ok: true,
      document: parsed.document,
    });
  });

  it.each([
    {
      object: {
        ...common,
        id: 'invalid-kind',
        kind: 'shape',
        shape: 'triangle',
        style: { fill: '#ffffff', stroke: '#000000', strokeWidth: 1 },
      },
      path: '$.workbook.sheets[0].objects[0].shape',
    },
    {
      object: {
        ...common,
        id: 'invalid-width',
        kind: 'shape',
        shape: 'rectangle',
        style: { fill: '#ffffff', stroke: '#000000', strokeWidth: -1 },
      },
      path: '$.workbook.sheets[0].objects[0].style.strokeWidth',
    },
  ])('rejects unsupported or unbounded shape data at $path', ({ object, path }) => {
    expect(parseSpreadsheetDocument(documentWithObjects([object]))).toMatchObject({
      ok: false,
      diagnostics: [
        expect.objectContaining({
          code: 'DOCUMENT_SCHEMA_INVALID',
          details: { path },
        }),
      ],
    });
  });
});
