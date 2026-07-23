import { describe, expect, it } from 'vitest';
import { createSpreadsheetDocument } from '../../../src/document';
import type { SpreadsheetDocument } from '../../../src/document';
import {
  compileSpreadsheetTemplate,
  expandAdvancedTemplate,
  type AdvancedCompileOptions,
  type SpreadsheetTemplate,
} from '../../../src/template';

const range = (startRow: number, endRow: number, startColumn = 0, endColumn = 1) => ({
  sheetId: 'sheet-1' as never,
  start: { row: startRow, column: startColumn },
  end: { row: endRow, column: endColumn },
});

function source(bindings: SpreadsheetTemplate['bindings']) {
  const base = createSpreadsheetDocument({
    id: 'document-1',
    sheetId: 'sheet-1',
    sheetName: 'Template',
  });
  const template = {
    id: 'template-1' as never,
    name: 'Advanced',
    bindings,
    printProfiles: [],
  } as SpreadsheetTemplate;
  return {
    template,
    document: {
      ...base,
      workbook: {
        ...base.workbook,
        sheets: [
          {
            ...base.workbook.sheets[0]!,
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'string', value: 'seed' } } },
              { row: 1, column: 0, cell: { input: { type: 'formula', source: '=A1' } } },
            ],
          },
        ],
      },
      templates: [template],
    } as SpreadsheetDocument,
  };
}

const options: AdvancedCompileOptions = {
  subtemplates: new Map(),
  limits: {
    maxExpandedCells: 1_000,
    maxExpandedRows: 1_000,
    maxExpandedColumns: 1_000,
    maxGeneratedSheets: 20,
    maxPages: 100,
    maxResources: 20,
    maxResourceBytes: 1_000_000,
    maxTotalResourceBytes: 2_000_000,
    maxResolveConcurrency: 2,
    maxLayoutTimeMs: 1_000,
    maxNestingDepth: 8,
  },
};

describe('TP2 advanced template structures', () => {
  it('builds a deterministic three-level containment tree and preserves parent scope', () => {
    const { document, template } = source([
      {
        id: 'outer',
        type: 'repeat-rows',
        range: range(0, 8),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'middle',
        type: 'repeat-rows',
        range: range(1, 6),
        source: 'item.children',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'inner',
        type: 'repeat-rows',
        range: range(2, 3),
        source: 'item.values',
        empty: 'remove',
        pageBreak: 'auto',
      },
    ] as never);
    const result = compileSpreadsheetTemplate(document, template.id, options);
    expect(result.hasErrors).toBe(false);
    expect(result.template?.ir.regionTree).toMatchObject([
      {
        bindingId: 'outer',
        children: [{ bindingId: 'middle', children: [{ bindingId: 'inner' }] }],
      },
    ]);
  });

  it('materializes nested repeat values in outer-to-inner input order', () => {
    const { document, template } = source([
      {
        id: 'outer',
        type: 'repeat-rows',
        range: range(0, 5),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'inner',
        type: 'repeat-rows',
        range: range(2, 2),
        source: 'item.values',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 2, column: 0 },
        expression: 'parent.name + ":" + item',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      {
        groups: [
          { name: 'A', values: ['a1', 'a2'] },
          { name: 'B', values: ['b1'] },
        ],
      },
      options.limits,
    );
    const values = expanded.document?.workbook.sheets[0]?.cells
      .filter(({ column, cell }) => column === 0 && cell.input.type === 'string')
      .map(({ cell }) => (cell.input.type === 'string' ? cell.input.value : ''));
    expect(values).toEqual(expect.arrayContaining(['A:a1', 'A:a2', 'B:b1']));
  });

  it('rejects partial overlap and reports a complete subtemplate cycle', () => {
    const { document, template } = source([
      {
        id: 'a',
        type: 'repeat-range',
        range: range(0, 2),
        source: 'rows',
        axis: 'vertical',
        empty: 'remove',
      },
      {
        id: 'b',
        type: 'repeat-range',
        range: range(2, 4),
        source: 'rows',
        axis: 'vertical',
        empty: 'remove',
      },
    ] as never);
    expect(compileSpreadsheetTemplate(document, template.id, options).diagnostics).toContainEqual(
      expect.objectContaining({ code: 'PARTIALLY_OVERLAPPING_REGION' }),
    );

    const a = {
      ...template,
      id: 'a',
      bindings: [
        { id: 'ab', type: 'subtemplate', range: range(0, 0), templateId: 'b', source: 'root' },
      ],
    } as never;
    const b = {
      ...template,
      id: 'b',
      bindings: [
        { id: 'ba', type: 'subtemplate', range: range(0, 0), templateId: 'a', source: 'root' },
      ],
    } as never;
    const cycle = compileSpreadsheetTemplate({ ...document, templates: [a] }, 'a' as never, {
      ...options,
      subtemplates: new Map([
        ['a' as never, a],
        ['b' as never, b],
      ]),
    });
    expect(cycle.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'SUBTEMPLATE_CYCLE',
        message: expect.stringContaining('a -> b -> a'),
      }),
    );
  });

  it('reuses an explicitly registered subtemplate without runtime discovery', () => {
    const { document, template } = source([
      {
        id: 'child-slot',
        type: 'subtemplate',
        range: range(1, 1, 0, 0),
        templateId: 'child',
        source: 'customer',
      },
    ] as never);
    const child = {
      id: 'child' as never,
      name: 'Child',
      bindings: [
        {
          id: 'child-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
          expression: 'item.name',
        },
      ],
      printProfiles: [],
    } as SpreadsheetTemplate;
    const advancedOptions = {
      ...options,
      subtemplates: new Map([['child' as never, child]]),
    };
    const compiled = compileSpreadsheetTemplate(document, template.id, advancedOptions).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { customer: { name: 'Ada' } },
      options.limits,
    );
    expect(expanded.document?.workbook.sheets[0]?.cells).toContainEqual(
      expect.objectContaining({
        row: 1,
        column: 0,
        cell: { input: { type: 'string', value: 'Ada' } },
      }),
    );
  });

  it('expands horizontal, two-dimensional, per-page and per-sheet repeats atomically', () => {
    const { document, template } = source([
      {
        id: 'columns',
        type: 'repeat-columns',
        range: range(0, 0, 0, 0),
        source: 'columns',
        empty: 'remove',
      },
      {
        id: 'matrix',
        type: 'repeat-range',
        range: range(1, 1, 0, 0),
        source: 'matrix',
        axis: 'both',
        empty: 'remove',
      },
      { id: 'pages', type: 'repeat-page', range: range(0, 1), source: 'pages', empty: 'remove' },
      {
        id: 'sheets',
        type: 'repeat-sheet',
        range: range(0, 1),
        source: 'sheets',
        name: 'item.name',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      {
        columns: [1, 2, 3],
        matrix: [
          [1, 2],
          [3, 4],
        ],
        pages: ['p1', 'p2'],
        sheets: [{ name: 'North' }, { name: 'South' }],
      },
      options.limits,
    );
    expect(expanded.document?.workbook.sheets.map(({ name }) => name)).toEqual([
      'Template',
      'North',
      'South',
    ]);
    expect(expanded.forcedPageBreaks.get('sheet-1')).toEqual(expect.arrayContaining([2]));
    expect(expanded.structuralMappings.length).toBeGreaterThan(0);
  });

  it('requires an explicit floating-object copy policy and stops before partial output', () => {
    const { document, template } = source([
      {
        id: 'objects',
        type: 'repeat-range',
        range: range(0, 0),
        source: 'rows',
        axis: 'vertical',
        empty: 'remove',
        objects: [{ id: 'logo' }],
      },
    ] as never);
    expect(compileSpreadsheetTemplate(document, template.id, options).diagnostics).toContainEqual(
      expect.objectContaining({ code: 'OBJECT_REPEAT_POLICY_REQUIRED' }),
    );
    const forbidden = {
      ...template,
      bindings: [{ ...(template.bindings[0] as object), objectPolicy: 'forbidden' }],
    } as unknown as SpreadsheetTemplate;
    expect(
      compileSpreadsheetTemplate({ ...document, templates: [forbidden] }, forbidden.id, options)
        .diagnostics,
    ).toContainEqual(expect.objectContaining({ code: 'OBJECT_REPEAT_FORBIDDEN' }));

    const valid = {
      ...template,
      bindings: [{ ...(template.bindings[0] as object), objectPolicy: 'per-item' }],
    } as unknown as SpreadsheetTemplate;
    const compiled = compileSpreadsheetTemplate(
      { ...document, templates: [valid] },
      valid.id,
      options,
    ).template!;
    expect(
      compileSpreadsheetTemplate(
        {
          ...document,
          templates: [
            {
              ...valid,
              bindings: [{ ...(valid.bindings[0] as object), objectPolicy: 'shared' }],
            } as unknown as SpreadsheetTemplate,
          ],
        },
        valid.id,
        options,
      ).hasErrors,
    ).toBe(false);
    const expanded = expandAdvancedTemplate(
      compiled,
      { rows: Array.from({ length: 100 }, () => 1) },
      {
        ...options.limits,
        maxExpandedCells: 2,
      },
    );
    expect(expanded.document).toBeUndefined();
    expect(expanded.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
  });
});
