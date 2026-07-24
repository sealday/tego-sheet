import { describe, expect, it } from 'vitest';
import { createSpreadsheetDocument, type SpreadsheetDocument } from '../../../src/document';
import { createFontMetrics } from '../../../src/presentation';
import type { PrintDisplayCommand } from '../../../src/print';
import {
  compileSpreadsheetTemplate,
  createResourceResolverRegistry,
  expandAdvancedTemplate,
  renderSpreadsheetTemplate,
  type AdvancedCompileOptions,
  type SpreadsheetTemplate,
} from '../../../src/template';

const limits: AdvancedCompileOptions['limits'] = {
  maxExpandedCells: 100,
  maxExpandedRows: 100,
  maxExpandedColumns: 100,
  maxGeneratedSheets: 10,
  maxExpandedObjects: 100,
  maxPages: 100,
  maxResources: 10,
  maxResourceBytes: 1_000_000,
  maxTotalResourceBytes: 2_000_000,
  maxResolveConcurrency: 2,
  maxLayoutTimeMs: 1_000,
  maxNestingDepth: 8,
};
const options: AdvancedCompileOptions = { subtemplates: new Map(), limits };
const objectRange = {
  sheetId: 'sheet-1' as never,
  start: { row: 0, column: 0 },
  end: { row: 0, column: 0 },
};

function png(): Uint8Array {
  const value = new Uint8Array(45);
  value.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  new DataView(value.buffer).setUint32(8, 13);
  value.set(new TextEncoder().encode('IHDR'), 12);
  new DataView(value.buffer).setUint32(16, 1);
  new DataView(value.buffer).setUint32(20, 1);
  value.set(new TextEncoder().encode('IEND'), 37);
  return value;
}

function fixture(
  policy: 'per-item' | 'shared' | 'forbidden' | undefined,
  kind: 'shape' | 'image' = 'shape',
): { readonly document: SpreadsheetDocument; readonly template: SpreadsheetTemplate } {
  const base = createSpreadsheetDocument({
    id: 'object-repeat',
    sheetId: 'sheet-1',
    sheetName: 'Template',
  });
  const shared = policy === 'shared';
  const object =
    kind === 'image'
      ? {
          id: 'badge',
          kind: 'image' as const,
          anchor: {
            type: 'one-cell' as const,
            cell: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
            offset: { x: 2, y: 3 },
            size: { width: 12, height: 8 },
          },
          zIndex: 1,
          locked: false,
          templateRepeat: 'per-item' as const,
          resourceId: 'badge-resource' as never,
          accessibility: { name: 'Badge' },
        }
      : {
          id: 'badge',
          kind: 'shape' as const,
          shape: 'rectangle' as const,
          anchor: shared
            ? {
                type: 'absolute' as const,
                rect: { x: 2, y: 3, width: 12, height: 8 },
              }
            : {
                type: 'one-cell' as const,
                cell: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
                offset: { x: 2, y: 3 },
                size: { width: 12, height: 8 },
              },
          zIndex: 1,
          locked: false,
          templateRepeat: shared ? ('shared' as const) : ('per-item' as const),
          style: { fill: '#ffeecc', stroke: '#112233', strokeWidth: 1 },
          accessibility: { name: 'Badge' },
        };
  const template = {
    id: 'template-1' as never,
    name: 'Object repeat',
    bindings: [
      {
        id: 'rows',
        type: 'repeat-rows',
        range: objectRange,
        source: 'items',
        empty: 'remove',
        pageBreak: 'auto',
        ...(policy === undefined ? {} : { objectPolicy: policy }),
        objects: [
          {
            id: 'badge',
            anchor: objectRange,
            anchorMode: shared ? 'absolute' : 'range',
            ...(kind === 'image' ? { resourceId: 'badge-resource' } : {}),
          },
        ],
      },
    ],
    printProfiles: [],
  } as unknown as SpreadsheetTemplate;
  return {
    template,
    document: {
      ...base,
      workbook: {
        ...base.workbook,
        sheets: [
          {
            ...base.workbook.sheets[0]!,
            cells: [{ row: 0, column: 0, cell: { input: { type: 'string', value: 'seed' } } }],
            objects: [object as never],
          },
        ],
      },
      resources:
        kind === 'image'
          ? { items: [{ id: 'badge-resource' as never, kind: 'image', mimeType: 'image/png' }] }
          : base.resources,
      templates: [template],
    },
  };
}

function compile(document: SpreadsheetDocument, template: SpreadsheetTemplate) {
  return compileSpreadsheetTemplate(document, template.id, options);
}

const environment = {
  locale: 'en-US',
  timeZone: 'UTC',
  dateSystem: 'excel-1900' as const,
  clock: new Date('2026-01-01T00:00:00.000Z'),
  fontMetrics: createFontMetrics({
    fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
    fallbackFont: 'Arial',
    fallback: { averageAdvance: 6, lineHeight: 12 },
  }),
};

function flattenCommands(commands: readonly PrintDisplayCommand[]): readonly PrintDisplayCommand[] {
  return commands.flatMap((command) =>
    command.kind === 'clip' || command.kind === 'group'
      ? [command, ...flattenCommands(command.commands)]
      : [command],
  );
}

describe('repeat-row floating object materialization', () => {
  it('creates deterministic per-item IDs, transformed anchors, and resource mappings', () => {
    const { document, template } = fixture('per-item', 'image');
    const compiled = compile(document, template).template!;
    const first = expandAdvancedTemplate(compiled, { items: ['north', 'south'] }, limits);
    const second = expandAdvancedTemplate(compiled, { items: ['north', 'south'] }, limits);

    expect(second).toEqual(first);
    expect(first.document?.workbook.sheets[0]?.objects).toMatchObject([
      {
        id: 'badge~rows~1',
        resourceId: 'badge-resource',
        anchor: { type: 'one-cell', cell: { row: 0, column: 0 }, offset: { x: 2, y: 3 } },
      },
      {
        id: 'badge~rows~2',
        resourceId: 'badge-resource',
        anchor: { type: 'one-cell', cell: { row: 1, column: 0 }, offset: { x: 2, y: 3 } },
      },
    ]);
    expect(first.objectMappings).toMatchObject([
      {
        objectId: 'badge~rows~1',
        resourceId: 'badge-resource',
        policy: 'per-item',
        itemIndex: 0,
        generated: objectRange,
      },
      {
        objectId: 'badge~rows~2',
        resourceId: 'badge-resource',
        policy: 'per-item',
        itemIndex: 1,
        generated: {
          sheetId: 'sheet-1',
          start: { row: 1, column: 0 },
          end: { row: 1, column: 0 },
        },
      },
    ]);
  });

  it('detects intersecting persistent row objects without duplicated binding metadata', () => {
    const seeded = fixture('per-item');
    const binding = seeded.template.bindings[0] as Extract<
      SpreadsheetTemplate['bindings'][number],
      { readonly type: 'repeat-rows' }
    >;
    const template = {
      ...seeded.template,
      bindings: [
        {
          id: binding.id,
          type: binding.type,
          range: binding.range,
          source: binding.source,
          empty: binding.empty,
          pageBreak: binding.pageBreak,
          objectPolicy: binding.objectPolicy,
        },
      ],
    } as SpreadsheetTemplate;
    const document = { ...seeded.document, templates: [template] };
    const compiled = compile(document, template);
    const expanded = expandAdvancedTemplate(compiled.template!, { items: [{}, {}] }, limits);

    expect(compiled.hasErrors).toBe(false);
    expect(expanded.document?.workbook.sheets[0]?.objects.map(({ id }) => id)).toEqual([
      'badge~rows~1',
      'badge~rows~2',
    ]);
  });

  it('preserves mapping-only references alongside persistent repeated objects', () => {
    const seeded = fixture('per-item');
    const binding = seeded.template.bindings[0] as Extract<
      SpreadsheetTemplate['bindings'][number],
      { readonly type: 'repeat-rows' }
    >;
    const template = {
      ...seeded.template,
      bindings: [
        {
          ...binding,
          objects: [
            ...(binding.objects ?? []),
            {
              id: 'legacy-overlay',
              anchor: objectRange,
              anchorMode: 'range' as const,
            },
          ],
        },
      ],
    } as SpreadsheetTemplate;
    const document = { ...seeded.document, templates: [template] };
    const result = expandAdvancedTemplate(
      compile(document, template).template!,
      { items: [{}, {}] },
      limits,
    );

    expect(result.objectMappings.map(({ objectId }) => objectId)).toEqual([
      'badge~rows~1',
      'legacy-overlay',
      'badge~rows~2',
      'legacy-overlay',
    ]);
  });

  it('keeps shared objects once and removes per-item objects with an empty repeated collection', () => {
    const shared = fixture('shared');
    const sharedResult = expandAdvancedTemplate(
      compile(shared.document, shared.template).template!,
      { items: [{}, {}, {}] },
      limits,
    );
    expect(sharedResult.document?.workbook.sheets[0]?.objects).toHaveLength(1);
    expect(sharedResult.document?.workbook.sheets[0]?.objects[0]).toMatchObject({
      id: 'badge',
      anchor: { type: 'absolute' },
    });
    expect(sharedResult.objectMappings).toEqual([
      expect.objectContaining({ objectId: 'badge', policy: 'shared', itemIndex: 0 }),
    ]);

    const perItem = fixture('per-item');
    const empty = expandAdvancedTemplate(
      compile(perItem.document, perItem.template).template!,
      { items: [] },
      limits,
    );
    expect(empty.document?.workbook.sheets[0]?.objects).toEqual([]);
    expect(empty.objectMappings).toEqual([]);
  });

  it('fails compilation for missing and forbidden repeat policies with the required-policy code', () => {
    for (const policy of [undefined, 'forbidden'] as const) {
      const { document, template } = fixture(policy);
      const result = compile(document, template);
      expect(result.hasErrors).toBe(true);
      expect(result.diagnostics).toContainEqual(
        expect.objectContaining({
          code: 'OBJECT_REPEAT_POLICY_REQUIRED',
          location: { bindingId: 'rows' },
        }),
      );
    }

    const detected = fixture(undefined);
    const binding = detected.template.bindings[0] as Extract<
      SpreadsheetTemplate['bindings'][number],
      { readonly type: 'repeat-rows' }
    >;
    const template = {
      ...detected.template,
      bindings: [
        {
          id: binding.id,
          type: binding.type,
          range: binding.range,
          source: binding.source,
          empty: binding.empty,
          pageBreak: binding.pageBreak,
        },
      ],
    } as SpreadsheetTemplate;
    const result = compile({ ...detected.document, templates: [template] }, template);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'OBJECT_REPEAT_POLICY_REQUIRED' }),
    );
  });

  it('rejects a binding policy that contradicts the persistent object policy', () => {
    const seeded = fixture('per-item');
    const object = seeded.document.workbook.sheets[0]!.objects[0]!;
    const document = {
      ...seeded.document,
      workbook: {
        ...seeded.document.workbook,
        sheets: [
          {
            ...seeded.document.workbook.sheets[0]!,
            objects: [{ ...object, templateRepeat: 'shared' as const }],
          },
        ],
      },
    };

    expect(compile(document, seeded.template).diagnostics).toContainEqual(
      expect.objectContaining({ code: 'OBJECT_REPEAT_POLICY_REQUIRED' }),
    );
  });

  it('enforces the object budget atomically before exposing a partial expanded document', () => {
    const { document, template } = fixture('per-item');
    const result = expandAdvancedTemplate(
      compile(document, template).template!,
      { items: [{}, {}] },
      { ...limits, maxExpandedObjects: 1 },
    );

    expect(result.document).toBeUndefined();
    expect(result.objectMappings).toEqual([]);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
    expect(document.workbook.sheets[0]?.objects).toHaveLength(1);
  });

  it('keeps object anchors aligned when a nested conditional removes repeated rows', () => {
    const seeded = fixture('per-item');
    const template = {
      ...seeded.template,
      bindings: [
        {
          ...(seeded.template.bindings[0] as object),
          range: {
            ...objectRange,
            end: { row: 1, column: 0 },
          },
          objects: [
            {
              id: 'badge',
              anchor: objectRange,
              anchorMode: 'range',
            },
          ],
        },
        {
          id: 'show-detail',
          type: 'conditional-range',
          range: {
            ...objectRange,
            start: { row: 1, column: 0 },
            end: { row: 1, column: 0 },
          },
          when: 'item.show',
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const document = {
      ...seeded.document,
      workbook: {
        ...seeded.document.workbook,
        sheets: [
          {
            ...seeded.document.workbook.sheets[0]!,
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'string' as const, value: 'item' } } },
              { row: 1, column: 0, cell: { input: { type: 'string' as const, value: 'detail' } } },
            ],
          },
        ],
      },
      templates: [template],
    };
    const result = expandAdvancedTemplate(
      compile(document, template).template!,
      { items: [{ show: false }, { show: true }] },
      limits,
    );

    expect(
      result.document?.workbook.sheets[0]?.objects.map((object) => ({
        id: object.id,
        row: object.anchor.type === 'one-cell' ? object.anchor.cell.row : -1,
      })),
    ).toEqual([
      { id: 'badge~rows~1', row: 0 },
      { id: 'badge~rows~2', row: 1 },
    ]);
  });

  it('renders one copied object per forced item page', async () => {
    const seeded = fixture('per-item');
    const template = {
      ...seeded.template,
      bindings: [
        {
          ...(seeded.template.bindings[0] as object),
          pageBreak: 'before-each-item',
        },
      ],
      printProfiles: [
        {
          id: 'profile',
          name: 'Pages',
          targets: [
            {
              type: 'range',
              range: {
                ...objectRange,
                end: { row: 2, column: 0 },
              },
            },
          ],
          page: {
            paper: { type: 'custom', width: 120, height: 100 },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: false,
          showHeadings: false,
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const document = { ...seeded.document, templates: [template] };
    const compiled = compile(document, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [{}, {}, {}] },
        profileId: 'profile',
        missingValue: 'error',
      },
      environment,
    );

    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.print.pages).toHaveLength(3);
    expect(
      result.document?.print.displayList.pages.map(
        (page) =>
          flattenCommands(page.commands).filter(
            (command) => command.kind === 'fill-rect' && command.color === '#ffeecc',
          ).length,
      ),
    ).toEqual([1, 1, 1]);
    expect(result.document?.objects.map(({ objectId }) => objectId)).toEqual([
      'badge~rows~1',
      'badge~rows~2',
      'badge~rows~3',
    ]);
    await result.document?.resources.dispose();
  });

  it('reports each materialized two-cell image once in generated object mappings', async () => {
    const seeded = fixture('per-item', 'image');
    const object = seeded.document.workbook.sheets[0]!.objects[0]!;
    const template = {
      ...seeded.template,
      printProfiles: [
        {
          id: 'profile',
          name: 'Images',
          targets: [{ type: 'range', range: { ...objectRange, end: { row: 1, column: 0 } } }],
          page: {
            paper: { type: 'custom', width: 120, height: 100 },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: false,
          showHeadings: false,
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const document = {
      ...seeded.document,
      workbook: {
        ...seeded.document.workbook,
        sheets: [
          {
            ...seeded.document.workbook.sheets[0]!,
            objects: [
              {
                ...object,
                anchor: {
                  type: 'two-cell' as const,
                  from: {
                    sheetId: 'sheet-1' as never,
                    row: 0,
                    column: 0,
                    offset: { x: 0, y: 0 },
                  },
                  to: {
                    sheetId: 'sheet-1' as never,
                    row: 0,
                    column: 0,
                    offset: { x: 12, y: 8 },
                  },
                },
              },
            ],
          },
        ],
      },
      templates: [template],
    };
    const compiled = compile(document, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [{}, {}] },
        profileId: 'profile',
        missingValue: 'error',
        resourceRefs: [
          {
            id: 'badge-resource',
            type: 'image',
            resolverId: 'test',
            key: 'badge',
            expectedMime: 'image/png',
          },
        ],
      },
      {
        ...environment,
        decodeImage: async () => ({ width: 1, height: 1, representation: {} }),
        resourceRegistry: createResourceResolverRegistry([
          {
            id: 'test',
            supports: () => true,
            resolve: async () => ({
              bytes: png(),
              mimeType: 'image/png',
              width: 1,
              height: 1,
            }),
          },
        ]),
      },
    );

    expect(result.diagnostics).toEqual([]);
    expect(result.document?.objects.map(({ objectId }) => objectId)).toEqual([
      'badge~rows~1',
      'badge~rows~2',
    ]);
    await result.document?.resources.dispose();
  });

  it('combines repeated objects with filtered and descending projected print rows', async () => {
    const seeded = fixture('per-item');
    const template = {
      ...seeded.template,
      bindings: [
        {
          ...(seeded.template.bindings[0] as object),
          range: {
            ...objectRange,
            start: { row: 1, column: 0 },
            end: { row: 1, column: 0 },
          },
          objects: [
            {
              id: 'badge',
              anchor: {
                ...objectRange,
                start: { row: 1, column: 0 },
                end: { row: 1, column: 0 },
              },
              anchorMode: 'range',
            },
          ],
        },
        {
          id: 'value',
          type: 'value',
          target: { sheetId: 'sheet-1', row: 1, column: 0 },
          expression: 'item',
        },
      ],
      printProfiles: [
        {
          id: 'profile',
          name: 'Projected',
          targets: [
            {
              type: 'range',
              range: {
                ...objectRange,
                end: { row: 3, column: 0 },
              },
            },
          ],
          page: {
            paper: { type: 'custom', width: 120, height: 160 },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: false,
          showHeadings: false,
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const sourceObject = seeded.document.workbook.sheets[0]!.objects[0]!;
    const document = {
      ...seeded.document,
      workbook: {
        ...seeded.document.workbook,
        sheets: [
          {
            ...seeded.document.workbook.sheets[0]!,
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'string' as const, value: 'header' } } },
              { row: 1, column: 0, cell: { input: { type: 'number' as const, value: 0 } } },
            ],
            objects: [
              {
                ...sourceObject,
                anchor: {
                  type: 'one-cell' as const,
                  cell: { sheetId: 'sheet-1' as never, row: 1, column: 0 },
                  offset: { x: 2, y: 3 },
                  size: { width: 12, height: 8 },
                },
              },
            ],
            filterViews: [
              {
                id: 'descending',
                name: 'Descending',
                range: {
                  ...objectRange,
                  end: { row: 3, column: 0 },
                },
                sorts: [{ column: 0, direction: 'descending' as const }],
                filters: [{ column: 0, operator: 'greaterThanOrEqual' as const, value: 2 }],
                visibility: 'document' as const,
              },
            ],
          },
        ],
      },
      templates: [template],
    };
    const compiled = compile(document, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [1, 2, 3] },
        profileId: 'profile',
        missingValue: 'error',
        activeFilterViews: [{ sheetId: 'sheet-1' as never, viewId: 'descending' }],
      },
      environment,
    );
    const objectRects =
      result.document?.print.displayList.pages.flatMap((page) =>
        flattenCommands(page.commands).flatMap((command) =>
          command.kind === 'fill-rect' && command.color === '#ffeecc' ? [command.rect] : [],
        ),
      ) ?? [];

    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(objectRects).toHaveLength(2);
    expect(objectRects.map(({ y }) => y).sort((left, right) => left - right)).toEqual([33, 53]);
    await result.document?.resources.dispose();
  });
});
