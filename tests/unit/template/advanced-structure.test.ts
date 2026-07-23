import { describe, expect, it, vi } from 'vitest';
import { createSpreadsheetDocument } from '../../../src/document';
import type { SpreadsheetDocument } from '../../../src/document';
import { createFontMetrics } from '../../../src/presentation';
import {
  compileSpreadsheetTemplate,
  expandAdvancedTemplate,
  renderSpreadsheetTemplate,
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
  it('keeps structural containment sheet-local and rejects coincident regions', () => {
    const first = createSpreadsheetDocument({
      id: 'sheet-local-regions',
      sheetId: 'sheet-1',
      sheetName: 'First',
    });
    const secondSheet = {
      ...first.workbook.sheets[0]!,
      id: 'sheet-2' as never,
      name: 'Second',
    };
    const sheetLocal = {
      id: 'sheet-local' as never,
      name: 'Sheet local',
      bindings: [
        {
          id: 'first',
          type: 'repeat-rows',
          range: range(0, 0, 0, 0),
          source: 'first',
          empty: 'remove',
          pageBreak: 'auto',
        },
        {
          id: 'second',
          type: 'repeat-rows',
          range: {
            ...range(0, 0, 0, 0),
            sheetId: 'sheet-2' as never,
          },
          source: 'second',
          empty: 'remove',
          pageBreak: 'auto',
        },
      ],
      printProfiles: [],
    } as unknown as SpreadsheetTemplate;
    const document = {
      ...first,
      workbook: { ...first.workbook, sheets: [first.workbook.sheets[0]!, secondSheet] },
      templates: [sheetLocal],
    };
    const compiled = compileSpreadsheetTemplate(document, sheetLocal.id, options);
    expect(compiled.hasErrors).toBe(false);
    expect(compiled.template?.ir.regionTree?.map(({ bindingId }) => bindingId)).toEqual([
      'first',
      'second',
    ]);

    const coincident = {
      ...sheetLocal,
      id: 'coincident' as never,
      bindings: [
        sheetLocal.bindings[0]!,
        {
          ...sheetLocal.bindings[0]!,
          id: 'same-place',
          type: 'repeat-columns',
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const rejected = compileSpreadsheetTemplate(
      { ...document, templates: [coincident] },
      coincident.id,
      options,
    );
    expect(rejected.hasErrors).toBe(true);
    expect(rejected.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'INVALID_NESTING' }),
    );
  });

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

  it('materializes a vertical repeat-range inside a repeat-rows scope', () => {
    const { document, template } = source([
      {
        id: 'outer',
        type: 'repeat-rows',
        range: range(0, 2, 0, 0),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'inner-range',
        type: 'repeat-range',
        range: range(1, 1, 0, 0),
        source: 'item.values',
        axis: 'vertical',
        empty: 'remove',
      },
      {
        id: 'inner-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 1, column: 0 },
        expression: 'item',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { groups: [{ values: ['A', 'B'] }, { values: ['C'] }] },
      options.limits,
    );

    expect(expanded.diagnostics).toEqual([]);
    expect(
      expanded.document?.workbook.sheets[0]?.cells
        .filter(({ cell }) => cell.input.type === 'string')
        .map(({ row, cell }) => [row, cell.input.type === 'string' ? cell.input.value : '']),
    ).toEqual(
      expect.arrayContaining([
        [1, 'A'],
        [2, 'B'],
        [5, 'C'],
      ]),
    );
    expect(
      expanded.structuralMappings.filter(({ bindingId }) => bindingId === 'inner-range'),
    ).toHaveLength(3);
  });

  it('binds nested horizontal repeats to their parent item scope', () => {
    const { document, template } = source([
      {
        id: 'outer',
        type: 'repeat-columns',
        range: range(0, 0, 0, 3),
        source: 'groups',
        empty: 'remove',
      },
      {
        id: 'inner',
        type: 'repeat-columns',
        range: range(0, 0, 1, 1),
        source: 'item.values',
        empty: 'remove',
      },
      {
        id: 'value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 1 },
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
      .filter(({ row, cell }) => row === 0 && cell.input.type === 'string')
      .map(({ cell }) => (cell.input.type === 'string' ? cell.input.value : ''));
    expect(values).toEqual(expect.arrayContaining(['A:a1', 'A:a2', 'B:b1']));
  });

  it('preserves immediate parent scopes through a three-level mixed-axis tree', () => {
    const { document, template } = source([
      {
        id: 'outer',
        type: 'repeat-rows',
        range: range(0, 4, 0, 5),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'middle',
        type: 'repeat-columns',
        range: range(1, 3, 1, 4),
        source: 'item.columns',
        empty: 'remove',
      },
      {
        id: 'inner',
        type: 'repeat-rows',
        range: range(2, 2, 2, 2),
        source: 'item.values',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'outer-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 0 },
        expression: 'item.name',
      },
      {
        id: 'middle-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 1, column: 1 },
        expression: 'parent.name + ":" + item.name',
      },
      {
        id: 'inner-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 2, column: 2 },
        expression: 'parent.name + ":" + item',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      {
        groups: [
          {
            name: 'A',
            columns: [
              { name: 'X', values: [1, 2] },
              { name: 'Y', values: [3] },
            ],
          },
          { name: 'B', columns: [{ name: 'Z', values: [4] }] },
        ],
      },
      options.limits,
    );
    const values = new Map(
      expanded.document?.workbook.sheets[0]?.cells
        .filter(({ cell }) => cell.input.type === 'string')
        .map(({ row, column, cell }) => [
          `${row}:${column}`,
          cell.input.type === 'string' ? cell.input.value : '',
        ]),
    );
    expect(values).toMatchObject(
      new Map([
        ['0:0', 'A'],
        ['1:1', 'A:X'],
        ['2:2', 'X:1'],
        ['3:2', 'X:2'],
        ['1:5', 'A:Y'],
        ['2:6', 'Y:3'],
        ['6:0', 'B'],
        ['7:1', 'B:Z'],
        ['8:2', 'Z:4'],
      ]),
    );
  });

  it('executes conditional children and column-to-row mixed nesting in parent scope', () => {
    const conditionalFixture = source([
      {
        id: 'outer',
        type: 'repeat-rows',
        range: range(0, 1, 0, 0),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'visible',
        type: 'conditional-range',
        range: range(1, 1, 0, 0),
        when: 'item.visible',
      },
      {
        id: 'name',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 0 },
        expression: 'item.name',
      },
    ] as never);
    const conditional = expandAdvancedTemplate(
      compileSpreadsheetTemplate(
        conditionalFixture.document,
        conditionalFixture.template.id,
        options,
      ).template!,
      {
        groups: [
          { name: 'A', visible: true },
          { name: 'B', visible: false },
        ],
      },
      options.limits,
    );
    expect(
      conditional.document?.workbook.sheets[0]?.cells.map(({ row, cell }) => [
        row,
        cell.input.type === 'string' ? cell.input.value : cell.input.type,
      ]),
    ).toEqual([
      [0, 'A'],
      [1, 'formula'],
      [2, 'B'],
    ]);

    const mixedFixture = source([
      {
        id: 'columns',
        type: 'repeat-columns',
        range: range(0, 2, 0, 1),
        source: 'groups',
        empty: 'remove',
      },
      {
        id: 'rows',
        type: 'repeat-rows',
        range: range(1, 1, 0, 0),
        source: 'item.rows',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'mixed-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 1, column: 0 },
        expression: 'parent.name + ":" + item',
      },
    ] as never);
    const mixed = expandAdvancedTemplate(
      compileSpreadsheetTemplate(mixedFixture.document, mixedFixture.template.id, options)
        .template!,
      {
        groups: [
          { name: 'A', rows: ['a1', 'a2'] },
          { name: 'B', rows: ['b1'] },
        ],
      },
      options.limits,
    );
    const values = mixed.document?.workbook.sheets[0]?.cells
      .filter(({ cell }) => cell.input.type === 'string')
      .map(({ row, column, cell }) => [
        row,
        column,
        cell.input.type === 'string' ? cell.input.value : '',
      ]);
    expect(values).toEqual(
      expect.arrayContaining([
        [1, 0, 'A:a1'],
        [2, 0, 'A:a2'],
        [1, 2, 'B:b1'],
      ]),
    );
    expect(mixed.structuralMappings.filter(({ bindingId }) => bindingId === 'rows')).toHaveLength(
      3,
    );
  });

  it('applies same-level regions bottom-right to top-left with final mappings', () => {
    const bindings = [
      {
        id: 'top',
        type: 'repeat-range',
        range: range(0, 0, 0, 0),
        source: 'top',
        axis: 'vertical',
        empty: 'remove',
      },
      {
        id: 'bottom',
        type: 'repeat-range',
        range: range(2, 2, 0, 0),
        source: 'bottom',
        axis: 'vertical',
        empty: 'remove',
      },
      {
        id: 'top-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 0 },
        expression: 'item',
      },
      {
        id: 'bottom-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 2, column: 0 },
        expression: 'item',
      },
    ] as never;
    const run = (ordered: SpreadsheetTemplate['bindings']) => {
      const fixture = source(ordered);
      const expanded = expandAdvancedTemplate(
        compileSpreadsheetTemplate(fixture.document, fixture.template.id, options).template!,
        { top: ['T1', 'T2'], bottom: ['B1', 'B2'] },
        options.limits,
      );
      return {
        values: expanded.document?.workbook.sheets[0]?.cells
          .filter(({ column, cell }) => column === 0 && cell.input.type === 'string')
          .map(({ row, cell }) => [row, cell.input.type === 'string' ? cell.input.value : '']),
        bottomMappings: expanded.structuralMappings
          .filter(({ bindingId }) => bindingId === 'bottom')
          .map(({ generated }) => generated.start.row),
      };
    };
    const expected = {
      values: [
        [0, 'T1'],
        [1, 'T2'],
        [3, 'B1'],
        [4, 'B2'],
      ],
      bottomMappings: [3, 4],
    };
    expect(run(bindings)).toEqual(expected);
    expect(run([...bindings].reverse())).toEqual(expected);
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
          expression: 'name',
        },
      ],
      printProfiles: [],
    } as unknown as SpreadsheetTemplate;
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

  it('pastes expanded subtemplate cells, formulas, merges, styles and validations', () => {
    const fixture = source([
      {
        id: 'child-slot',
        type: 'subtemplate',
        range: range(2, 2, 0, 1),
        templateId: 'child',
        source: 'customer',
      },
    ] as never);
    const child = {
      id: 'child' as never,
      name: 'Child',
      bindings: [
        {
          id: 'child-rows',
          type: 'repeat-rows',
          range: range(0, 0, 0, 1),
          source: 'items',
          empty: 'remove',
          pageBreak: 'auto',
        },
        {
          id: 'child-value',
          type: 'value',
          target: { sheetId: 'sheet-1', row: 0, column: 0 },
          expression: 'item',
        },
      ],
      printProfiles: [],
    } as unknown as SpreadsheetTemplate;
    const document = {
      ...fixture.document,
      workbook: {
        ...fixture.document.workbook,
        sheets: [
          {
            ...fixture.document.workbook.sheets[0]!,
            cells: [
              {
                row: 0,
                column: 0,
                cell: {
                  input: { type: 'string', value: 'seed' },
                  styleId: 'style-1',
                  validationId: 'validation-1',
                  metadata: { source: 'child' },
                },
              },
              { row: 0, column: 1, cell: { input: { type: 'formula', source: '=A1' } } },
            ],
            merges: [{ start: { row: 0, column: 0 }, end: { row: 0, column: 1 } }],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const advancedOptions = {
      ...options,
      subtemplates: new Map([['child' as never, child]]),
    };
    const compiled = compileSpreadsheetTemplate(
      document,
      fixture.template.id,
      advancedOptions,
    ).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { customer: { items: ['Ada', 'Lin'] } },
      options.limits,
    );
    const sheet = expanded.document?.workbook.sheets[0];
    expect(sheet?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          row: 2,
          column: 0,
          cell: expect.objectContaining({
            input: { type: 'string', value: 'Ada' },
            styleId: 'style-1',
            validationId: 'validation-1',
            metadata: { source: 'child' },
          }),
        }),
        expect.objectContaining({
          row: 2,
          column: 1,
          cell: { input: { type: 'formula', source: '=A3' } },
        }),
        expect.objectContaining({
          row: 3,
          column: 0,
          cell: expect.objectContaining({ input: { type: 'string', value: 'Lin' } }),
        }),
        expect.objectContaining({
          row: 3,
          column: 1,
          cell: { input: { type: 'formula', source: '=A4' } },
        }),
      ]),
    );
    expect(sheet?.merges).toEqual(
      expect.arrayContaining([
        { start: { row: 2, column: 0 }, end: { row: 2, column: 1 } },
        { start: { row: 3, column: 0 }, end: { row: 3, column: 1 } },
      ]),
    );
    expect(
      expanded.structuralMappings.find(({ bindingId }) => bindingId === 'child-slot')?.generated,
    ).toEqual(range(2, 3, 0, 1));
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
      { id: 'pages', type: 'repeat-page', range: range(2, 2), source: 'pages', empty: 'remove' },
      {
        id: 'sheets',
        type: 'repeat-sheet',
        range: range(3, 3),
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
    expect(expanded.forcedPageBreaks.get('sheet-1')).toEqual(expect.arrayContaining([3]));
    expect(expanded.structuralMappings.length).toBeGreaterThan(0);
  });

  it('routes a sheet print target to each generated repeat-sheet boundary', async () => {
    const fixture = source([
      {
        id: 'sheets',
        type: 'repeat-sheet',
        range: range(0, 0, 0, 0),
        source: 'regions',
        name: 'item.name',
      },
      {
        id: 'region-name',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 0 },
        expression: 'item.name',
      },
    ] as never);
    const printable = {
      ...fixture.template,
      printProfiles: [
        {
          id: 'profile',
          name: 'Generated sheets',
          targets: [{ type: 'sheet', sheetId: 'sheet-1' }],
          page: {
            paper: { type: 'custom', width: 240, height: 160 },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          manualBreaks: [],
          showGridlines: true,
          showHeadings: false,
        },
        {
          id: 'range-profile',
          name: 'Generated sheet ranges',
          targets: [{ type: 'range', range: range(0, 1, 0, 0) }],
          page: {
            paper: { type: 'custom', width: 240, height: 160 },
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: { type: 'fixed', value: 1 },
          },
          repeatRows: range(0, 0, 0, 0),
          repeatColumns: range(0, 1, 0, 0),
          manualBreaks: [{ sheetId: 'sheet-1', beforeRow: 1 }],
          showGridlines: true,
          showHeadings: false,
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const document = {
      ...fixture.document,
      workbook: {
        ...fixture.document.workbook,
        sheets: fixture.document.workbook.sheets.map((sheet) => ({
          ...sheet,
          cells: sheet.cells.map((entry) =>
            entry.row === 1
              ? {
                  ...entry,
                  cell: { input: { type: 'string' as const, value: 'BODY' } },
                }
              : entry,
          ),
        })),
      },
      templates: [printable],
    };
    const compiled = compileSpreadsheetTemplate(document, printable.id, options).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { regions: [{ name: 'North' }, { name: 'South' }] },
        profileId: 'profile',
        missingValue: 'error',
      },
      {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        clock: new Date('2026-01-01T00:00:00.000Z'),
        fontMetrics: createFontMetrics({
          fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
          fallbackFont: 'Arial',
          fallback: { averageAdvance: 6, lineHeight: 12 },
        }),
      },
    );
    expect(result.diagnostics).toEqual([]);
    expect(
      result.document?.workbook.sheets.slice(1).map(({ id, name, cells }) => ({
        id,
        name,
        value: cells[0]?.cell.input,
      })),
    ).toEqual([
      {
        id: 'sheet-1~sheets~1',
        name: 'North',
        value: { type: 'string', value: 'North' },
      },
      {
        id: 'sheet-1~sheets~2',
        name: 'South',
        value: { type: 'string', value: 'South' },
      },
    ]);
    expect(result.document?.print.pages.map(({ targetId }) => targetId)).toEqual([
      expect.stringContaining('sheet-1~sheets~1'),
      expect.stringContaining('sheet-1~sheets~2'),
    ]);
    await result.document?.resources.dispose();

    const rangeResult = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { regions: [{ name: 'North' }, { name: 'South' }] },
        profileId: 'range-profile',
        missingValue: 'error',
      },
      {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        clock: new Date('2026-01-01T00:00:00.000Z'),
        fontMetrics: createFontMetrics({
          fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
          fallbackFont: 'Arial',
          fallback: { averageAdvance: 6, lineHeight: 12 },
        }),
      },
    );
    expect(rangeResult.diagnostics).toEqual([]);
    expect(rangeResult.document?.print.pages.map(({ targetId }) => targetId)).toEqual([
      expect.stringContaining('sheet-1~sheets~1'),
      expect.stringContaining('sheet-1~sheets~1'),
      expect.stringContaining('sheet-1~sheets~2'),
      expect.stringContaining('sheet-1~sheets~2'),
    ]);
    expect(
      rangeResult.document?.print.displayList.pages.map(({ commands }) =>
        commands
          .filter((command) => command.kind === 'text')
          .map((command) => command.text)
          .filter((text) => text === 'North' || text === 'South'),
      ),
    ).toEqual([['North'], ['North'], ['South'], ['South']]);
    await rangeResult.document?.resources.dispose();
  });

  it.each([
    {
      label: 'horizontal',
      binding: {
        id: 'repeat',
        type: 'repeat-columns',
        range: range(0, 0, 0, 0),
        source: 'items',
        empty: 'remove',
      },
      data: { items: ['A', 'B', 'C'] },
      expected: [
        [0, 0, 'A'],
        [0, 1, 'B'],
        [0, 2, 'C'],
      ],
    },
    {
      label: 'two-dimensional',
      binding: {
        id: 'repeat',
        type: 'repeat-range',
        range: range(0, 0, 0, 0),
        source: 'items',
        axis: 'both',
        empty: 'remove',
      },
      data: {
        items: [
          ['A', 'B'],
          ['C', 'D'],
        ],
      },
      expected: [
        [0, 0, 'A'],
        [0, 1, 'B'],
        [1, 0, 'C'],
        [1, 1, 'D'],
      ],
    },
    {
      label: 'per-page',
      binding: {
        id: 'repeat',
        type: 'repeat-page',
        range: range(0, 0, 0, 0),
        source: 'items',
        empty: 'remove',
      },
      data: { items: ['A', 'B', 'C'] },
      expected: [
        [0, 0, 'A'],
        [1, 0, 'B'],
        [2, 0, 'C'],
      ],
    },
  ])('binds each $label copy in its own item scope', ({ binding, data, expected }) => {
    const { document, template } = source([
      binding,
      {
        id: 'value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 0 },
        expression: 'item',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const result = expandAdvancedTemplate(compiled, data, options.limits);
    const cells = result.document?.workbook.sheets[0]?.cells.map(({ row, column, cell }) => [
      row,
      column,
      cell.input.type === 'string' ? cell.input.value : '',
    ]);
    expect(cells).toEqual(expect.arrayContaining(expected));
    if (binding.type === 'repeat-page') {
      expect(result.forcedPageBreaks.get('sheet-1')).toEqual([1, 2]);
    }
  });

  it('removes empty advanced regions and preserves repeated merges', () => {
    const { document, template } = source([
      {
        id: 'repeat',
        type: 'repeat-columns',
        range: range(0, 0, 0, 0),
        source: 'items',
        empty: 'remove',
      },
    ] as never);
    const empty = expandAdvancedTemplate(
      compileSpreadsheetTemplate(document, template.id, options).template!,
      { items: [] },
      options.limits,
    );
    expect(empty.document?.workbook.sheets[0]?.cells).not.toContainEqual(
      expect.objectContaining({ row: 0, column: 0 }),
    );

    const nested = source([
      {
        id: 'outer',
        type: 'repeat-rows',
        range: range(0, 1),
        source: 'items',
        empty: 'remove',
        pageBreak: 'auto',
      },
    ] as never);
    const mergedDocument = {
      ...nested.document,
      workbook: {
        ...nested.document.workbook,
        sheets: [
          {
            ...nested.document.workbook.sheets[0]!,
            merges: [{ start: { row: 0, column: 0 }, end: { row: 0, column: 1 } }],
          },
        ],
      },
    };
    const merged = expandAdvancedTemplate(
      compileSpreadsheetTemplate(mergedDocument, nested.template.id, options).template!,
      { items: [{}, {}] },
      options.limits,
    );
    expect(merged.document?.workbook.sheets[0]?.merges).toEqual([
      { start: { row: 0, column: 0 }, end: { row: 0, column: 1 } },
      { start: { row: 2, column: 0 }, end: { row: 2, column: 1 } },
    ]);
  });

  it('requires an explicit object policy and rejects estimated allocation before cloning', () => {
    const { document, template } = source([
      {
        id: 'objects',
        type: 'repeat-range',
        range: range(0, 0),
        source: 'rows',
        axis: 'vertical',
        empty: 'remove',
        objects: [
          {
            id: 'logo',
            anchor: range(0, 0, 0, 0),
            anchorMode: 'range',
            resourceId: 'logo-resource',
          },
        ],
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
    const objectExpansion = expandAdvancedTemplate(
      compiled,
      { rows: ['north', 'south'] },
      options.limits,
    );
    expect(objectExpansion.objectMappings).toEqual([
      {
        objectId: 'logo',
        resourceId: 'logo-resource',
        policy: 'per-item',
        itemIndex: 0,
        source: range(0, 0, 0, 0),
        generated: range(0, 0, 0, 0),
      },
      {
        objectId: 'logo',
        resourceId: 'logo-resource',
        policy: 'per-item',
        itemIndex: 1,
        source: range(0, 0, 0, 0),
        generated: range(1, 1, 0, 0),
      },
    ]);
    expect(
      compileSpreadsheetTemplate(
        {
          ...document,
          templates: [
            {
              ...valid,
              bindings: [
                {
                  ...(valid.bindings[0] as object),
                  objectPolicy: 'shared',
                  objects: [
                    {
                      id: 'logo',
                      anchor: range(0, 0, 0, 0),
                      anchorMode: 'absolute',
                      resourceId: 'logo-resource',
                    },
                  ],
                },
              ],
            } as unknown as SpreadsheetTemplate,
          ],
        },
        valid.id,
        options,
      ),
    ).toMatchObject({ hasErrors: false });
    const expanded = expandAdvancedTemplate(
      compiled,
      { rows: Array.from({ length: 100 }, () => 1) },
      {
        ...options.limits,
        maxExpandedCells: 2,
      },
    );
    expect(expanded.document).toBeUndefined();
    expect(expanded.structuralMappings).toEqual([]);
    expect(expanded.forcedPageBreaks.size).toBe(0);
    expect(document.workbook.sheets[0]?.cells).toHaveLength(2);
    expect(expanded.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
  });

  it('preflights pure nested columns before evaluating cloned value formatters', () => {
    const tick = vi.fn((value: unknown) => value);
    const { document, template } = source([
      {
        id: 'outer-columns',
        type: 'repeat-columns',
        range: range(0, 1, 0, 2),
        source: 'groups',
        empty: 'remove',
      },
      {
        id: 'inner-columns',
        type: 'repeat-columns',
        range: range(0, 1, 1, 1),
        source: 'item.values',
        empty: 'remove',
      },
      {
        id: 'formatted',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 1 },
        expression: 'tick(item)',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { groups: [{ values: ['A', 'B', 'C'] }] },
      { ...options.limits, maxExpandedCells: 2 },
      { tick },
    );

    expect(expanded.document).toBeUndefined();
    expect(expanded.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
    expect(tick).not.toHaveBeenCalled();
  });

  it('expands a pure horizontal tree made from repeat-range bindings', () => {
    const { document, template } = source([
      {
        id: 'outer-range',
        type: 'repeat-range',
        range: range(0, 1, 0, 2),
        source: 'groups',
        axis: 'horizontal',
        empty: 'remove',
      },
      {
        id: 'inner-range',
        type: 'repeat-range',
        range: range(0, 1, 1, 1),
        source: 'item.values',
        axis: 'horizontal',
        empty: 'remove',
      },
      {
        id: 'inner-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 0, column: 1 },
        expression: 'item',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { groups: [{ values: ['A', 'B'] }, { values: ['C'] }] },
      options.limits,
    );

    expect(expanded.diagnostics).toEqual([]);
    expect(
      expanded.document?.workbook.sheets[0]?.cells
        .filter(({ cell }) => cell.input.type === 'string')
        .map(({ column, cell }) => [column, cell.input.type === 'string' ? cell.input.value : '']),
    ).toEqual(
      expect.arrayContaining([
        [1, 'A'],
        [2, 'B'],
        [5, 'C'],
      ]),
    );
    expect(
      expanded.structuralMappings.filter(({ bindingId }) => bindingId === 'inner-range'),
    ).toHaveLength(3);
  });

  it('preflights the column limit of a mixed horizontal root', () => {
    const { document, template } = source([
      {
        id: 'outer-columns',
        type: 'repeat-columns',
        range: range(0, 1, 0, 1),
        source: 'groups',
        empty: 'remove',
      },
      {
        id: 'inner-rows',
        type: 'repeat-rows',
        range: range(1, 1, 0, 1),
        source: 'item.values',
        empty: 'remove',
        pageBreak: 'auto',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { groups: Array.from({ length: 10 }, () => ({ values: [] })) },
      { ...options.limits, maxExpandedColumns: 2 },
    );

    expect(expanded.document).toBeUndefined();
    expect(expanded.structuralMappings).toEqual([]);
    expect(expanded.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
  });

  it('materializes a nested two-dimensional range in deterministic row-major order', () => {
    const { document, template } = source([
      {
        id: 'outer-rows',
        type: 'repeat-rows',
        range: range(0, 2, 0, 2),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'matrix',
        type: 'repeat-range',
        range: range(1, 1, 1, 1),
        source: 'item.matrix',
        axis: 'both',
        empty: 'remove',
      },
      {
        id: 'matrix-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 1, column: 1 },
        expression: 'item',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      {
        groups: [
          {
            matrix: [
              ['A', 'B'],
              ['C', 'D'],
            ],
          },
        ],
      },
      options.limits,
    );

    expect(expanded.diagnostics).toEqual([]);
    expect(
      expanded.document?.workbook.sheets[0]?.cells
        .filter(({ cell }) => cell.input.type === 'string' && cell.input.value !== 'seed')
        .map(({ row, column, cell }) => [
          row,
          column,
          cell.input.type === 'string' ? cell.input.value : '',
        ]),
    ).toEqual([
      [1, 1, 'A'],
      [1, 2, 'B'],
      [2, 1, 'C'],
      [2, 2, 'D'],
    ]);
    expect(
      expanded.structuralMappings
        .filter(({ bindingId }) => bindingId === 'matrix')
        .map(({ itemIndex, generated }) => [
          itemIndex,
          generated.start.row,
          generated.start.column,
        ]),
    ).toEqual([
      [0, 1, 1],
      [1, 1, 2],
      [2, 2, 1],
      [3, 2, 2],
    ]);
  });

  it('preflights nested two-dimensional cells before evaluating value formatters', () => {
    const tick = vi.fn((value: unknown) => value);
    const { document, template } = source([
      {
        id: 'outer-rows',
        type: 'repeat-rows',
        range: range(0, 2, 0, 2),
        source: 'groups',
        empty: 'remove',
        pageBreak: 'auto',
      },
      {
        id: 'matrix',
        type: 'repeat-range',
        range: range(1, 1, 1, 1),
        source: 'item.matrix',
        axis: 'both',
        empty: 'remove',
      },
      {
        id: 'matrix-value',
        type: 'value',
        target: { sheetId: 'sheet-1', row: 1, column: 1 },
        expression: 'tick(item)',
      },
    ] as never);
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      {
        groups: [
          {
            matrix: [
              ['A', 'B'],
              ['C', 'D'],
            ],
          },
        ],
      },
      { ...options.limits, maxExpandedCells: 2 },
      { tick },
    );

    expect(expanded.document).toBeUndefined();
    expect(expanded.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
    expect(tick).not.toHaveBeenCalled();
  });

  it('counts value-created cells before cloning an empty advanced range', () => {
    const base = createSpreadsheetDocument({
      id: 'blank-allocation',
      sheetId: 'sheet-1',
      sheetName: 'Blank',
    });
    const template = {
      id: 'blank-allocation-template' as never,
      name: 'Blank allocation',
      bindings: [
        {
          id: 'copies',
          type: 'repeat-range',
          range: range(0, 0, 0, 0),
          source: 'items',
          axis: 'vertical',
          empty: 'remove',
        },
        {
          id: 'created-value',
          type: 'value',
          target: { sheetId: 'sheet-1', row: 0, column: 0 },
          expression: 'item',
        },
      ],
      printProfiles: [],
    } as unknown as SpreadsheetTemplate;
    const document = { ...base, templates: [template] };
    const compiled = compileSpreadsheetTemplate(document, template.id, options).template!;
    const expanded = expandAdvancedTemplate(
      compiled,
      { items: [1, 2, 3, 4, 5] },
      { ...options.limits, maxExpandedCells: 1 },
    );
    expect(expanded.document).toBeUndefined();
    expect(expanded.diagnostics).toContainEqual(
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    );
  });
});
