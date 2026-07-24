import { describe, expect, it } from 'vitest';
import type { SpreadsheetDocument } from '../../../src/document';
import { createFontMetrics } from '../../../src/presentation';
import type { PrintDisplayCommand } from '../../../src/print';
import { compileSpreadsheetTemplate } from '../../../src/template/compiler';
import { renderSpreadsheetTemplate } from '../../../src/template/render';
import type { SpreadsheetTemplate } from '../../../src/template/model';

const source = {
  schemaVersion: 2,
  id: 'document-1',
  workbook: {
    sheets: [
      {
        id: 'sheet-1',
        name: 'Invoice',
        visibility: 'hidden',
        conditionalFormatting: [
          {
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 1, column: 0 },
              end: { row: 1, column: 1 },
            },
            operator: 'greaterThan',
            formula: 'A2',
            style: { bold: true },
          },
        ],
        cells: [
          { row: 0, column: 0, cell: { input: { type: 'string', value: 'Name' } } },
          { row: 1, column: 0, cell: { input: { type: 'string', value: 'item' } } },
          { row: 1, column: 1, cell: { input: { type: 'formula', source: '=A2' } } },
          { row: 3, column: 0, cell: { input: { type: 'string', value: 'conditional' } } },
        ],
        merges: [],
        rows: [{ index: 1, height: 20 }],
        columns: [
          { index: 0, width: 90 },
          { index: 1, width: 90 },
        ],
      },
    ],
    styles: [],
    validations: [],
    settings: { dateSystem: 'excel-1900' },
  },
  templates: [],
  resources: { items: [] },
  extensions: {},
} as unknown as SpreadsheetDocument;

const template: SpreadsheetTemplate = {
  id: 'template-1' as never,
  name: 'Invoice',
  bindings: [
    {
      id: 'name' as never,
      type: 'value',
      target: { sheetId: 'sheet-1' as never, row: 0, column: 1 },
      expression: 'customer.name',
    },
    {
      id: 'lines' as never,
      type: 'repeat-rows',
      range: {
        sheetId: 'sheet-1' as never,
        start: { row: 1, column: 0 },
        end: { row: 1, column: 1 },
      },
      source: 'items',
      empty: 'remove',
      pageBreak: 'auto',
    },
    {
      id: 'show-note' as never,
      type: 'conditional-range',
      range: {
        sheetId: 'sheet-1' as never,
        start: { row: 3, column: 0 },
        end: { row: 3, column: 1 },
      },
      when: 'showNote',
    },
  ],
  printProfiles: [
    {
      id: 'profile-1',
      name: 'A4',
      targets: [
        {
          type: 'range',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 0, column: 0 },
            end: { row: 5, column: 1 },
          },
        },
      ],
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
  ],
};

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

describe('template render pipeline', () => {
  it('consumes an explicit session view snapshot in print without persisting selection', async () => {
    const viewSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'string', value: 'status' } } },
              { row: 1, column: 0, cell: { input: { type: 'string', value: 'keep' } } },
              { row: 2, column: 0, cell: { input: { type: 'string', value: 'drop' } } },
            ],
            filterViews: [
              {
                id: 'kept',
                name: 'Kept',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 0, column: 0 },
                  end: { row: 2, column: 0 },
                },
                sorts: [],
                filters: [{ column: 0, operator: 'equal', value: 'keep' }],
                visibility: 'document',
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const viewTemplate = { ...template, bindings: [] };
    const compiled = compileSpreadsheetTemplate(viewSource, viewTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
        activeFilterViews: [{ sheetId: 'sheet-1' as never, viewId: 'kept' }],
      },
      environment,
    );
    const texts =
      result.document?.print.displayList.pages.flatMap((page) =>
        page.commands.flatMap((command) => (command.kind === 'text' ? [command.text] : [])),
      ) ?? [];
    expect(texts).toContain('keep');
    expect(texts).not.toContain('drop');
    expect(result.document?.workbook.sheets[0]).not.toHaveProperty('activeViewId');
  });

  it('paginates a filtered and sorted view in global projected row order', async () => {
    const viewSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: ['header', '1', '2', '3', '4'].map((value, row) => ({
              row,
              column: 0,
              cell: { input: { type: 'string' as const, value } },
            })),
            rows: [],
            filterViews: [
              {
                id: 'descending',
                name: 'Descending',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 0, column: 0 },
                  end: { row: 4, column: 0 },
                },
                sorts: [{ column: 0, direction: 'descending' }],
                filters: [{ column: 0, operator: 'greaterThanOrEqual', value: '2' }],
                visibility: 'document',
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const viewTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          targets: [
            {
              type: 'range',
              range: {
                sheetId: 'sheet-1' as never,
                start: { row: 0, column: 0 },
                end: { row: 4, column: 0 },
              },
            },
          ],
          page: {
            ...template.printProfiles[0]!.page,
            paper: { type: 'custom', width: 240, height: 60 },
          },
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(viewSource, viewTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
        activeFilterViews: [{ sheetId: 'sheet-1' as never, viewId: 'descending' }],
      },
      environment,
    );
    const texts =
      result.document?.print.displayList.pages.flatMap((page) =>
        page.commands.flatMap((command) => (command.kind === 'text' ? [command.text] : [])),
      ) ?? [];

    expect(texts).toEqual(['header', '4', '3', '2']);
    expect(result.document?.print.pages).toHaveLength(2);
  });

  it('prints persistent structured table filters and sorting in the screen projection order', async () => {
    const tableSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: ['header', '1', '2', '3', '4'].map((value, row) => ({
              row,
              column: 0,
              cell: { input: { type: 'string' as const, value } },
            })),
            rows: [],
            tables: [
              {
                id: 'table-values',
                name: 'Values',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 0, column: 0 },
                  end: { row: 4, column: 0 },
                },
                columns: [{ id: 'table-value', name: 'Value' }],
                filter: {
                  filters: [{ column: 0, operator: 'in', values: ['2', '3', '4'] }],
                  sort: { column: 0, direction: 'desc' },
                },
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const tableTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          targets: [
            {
              type: 'range',
              range: {
                sheetId: 'sheet-1' as never,
                start: { row: 0, column: 0 },
                end: { row: 4, column: 0 },
              },
            },
          ],
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(tableSource, tableTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    const texts =
      result.document?.print.displayList.pages.flatMap((page) =>
        page.commands.flatMap((command) => (command.kind === 'text' ? [command.text] : [])),
      ) ?? [];

    expect(texts).toEqual(['header', '4', '3', '2']);
  });

  it('renders persistent floating text objects through the shared display list', async () => {
    const objectSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            objects: [
              {
                id: 'notice',
                kind: 'text-box',
                anchor: {
                  type: 'absolute',
                  rect: { x: 20, y: 30, width: 80, height: 24 },
                },
                zIndex: 1,
                locked: false,
                templateRepeat: 'shared',
                text: 'Persistent notice',
                style: { color: '#123456', fontFamily: 'Arial', fontSize: 11 },
                accessibility: { name: 'Persistent notice' },
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const compiled = compileSpreadsheetTemplate(objectSource, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { customer: { name: 'A' }, items: [{}], showNote: false },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );

    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(
      flattenCommands(result.document?.print.displayList.pages[0]?.commands ?? []),
    ).toContainEqual(
      expect.objectContaining({
        kind: 'text',
        text: 'Persistent notice',
        color: '#123456',
      }),
    );
  });

  it('keeps rotated object bounds that cross into the printable page', async () => {
    const objectSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            objects: [
              {
                id: 'rotated-edge',
                kind: 'shape',
                shape: 'line',
                anchor: {
                  type: 'absolute',
                  rect: { x: 181, y: 0, width: 2, height: 100 },
                },
                rotation: 90,
                zIndex: 1,
                locked: false,
                templateRepeat: 'shared',
                style: { stroke: '#123456', strokeWidth: 4 },
                accessibility: { name: 'Rotated edge' },
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const compiled = compileSpreadsheetTemplate(objectSource, {
      ...template,
      bindings: [],
    }).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );

    expect(
      result.document?.print.displayList.pages.flatMap((page) =>
        page.commands.filter((command) => command.kind === 'group'),
      ),
    ).toEqual([expect.objectContaining({ kind: 'group', rotation: 90 })]);
  });

  it('projects persistent object anchors through sorted and filtered print rows', async () => {
    const objectSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: ['header', '1', '2', '3', '4'].map((value, row) => ({
              row,
              column: 0,
              cell: { input: { type: 'string' as const, value } },
            })),
            rows: [],
            filterViews: [
              {
                id: 'descending',
                name: 'Descending',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 0, column: 0 },
                  end: { row: 4, column: 0 },
                },
                sorts: [{ column: 0, direction: 'descending' }],
                filters: [{ column: 0, operator: 'greaterThanOrEqual', value: '2' }],
                visibility: 'document',
              },
            ],
            objects: [
              {
                id: 'sorted',
                kind: 'text-box',
                anchor: {
                  type: 'one-cell',
                  cell: { sheetId: 'sheet-1', row: 4, column: 0 },
                  offset: { x: 0, y: 0 },
                  size: { width: 40, height: 10 },
                },
                zIndex: 1,
                locked: false,
                templateRepeat: 'shared',
                text: 'Sorted object',
                style: { color: '#111111', fontFamily: 'Arial', fontSize: 10 },
                accessibility: { name: 'Sorted object' },
              },
              {
                id: 'filtered',
                kind: 'text-box',
                anchor: {
                  type: 'one-cell',
                  cell: { sheetId: 'sheet-1', row: 1, column: 0 },
                  offset: { x: 0, y: 0 },
                  size: { width: 40, height: 10 },
                },
                zIndex: 2,
                locked: false,
                templateRepeat: 'shared',
                text: 'Filtered object',
                style: { color: '#111111', fontFamily: 'Arial', fontSize: 10 },
                accessibility: { name: 'Filtered object' },
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const viewTemplate = { ...template, bindings: [] };
    const compiled = compileSpreadsheetTemplate(objectSource, viewTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
        activeFilterViews: [{ sheetId: 'sheet-1' as never, viewId: 'descending' }],
      },
      environment,
    );
    const objectCommands =
      result.document?.print.displayList.pages.flatMap((page) =>
        flattenCommands(page.commands).filter(
          (command) =>
            command.kind === 'text' &&
            (command.text === 'Sorted object' || command.text === 'Filtered object'),
        ),
      ) ?? [];

    expect(objectCommands).toEqual([
      expect.objectContaining({ kind: 'text', text: 'Sorted object', y: 40 }),
    ]);
  });

  it('maps two-cell object markers across projected page boundaries', async () => {
    const objectSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: ['header', '1', '2', '3', '4'].map((value, row) => ({
              row,
              column: 0,
              cell: { input: { type: 'string' as const, value } },
            })),
            rows: [],
            filterViews: [
              {
                id: 'descending',
                name: 'Descending',
                range: {
                  sheetId: 'sheet-1',
                  start: { row: 0, column: 0 },
                  end: { row: 4, column: 0 },
                },
                sorts: [{ column: 0, direction: 'descending' }],
                filters: [{ column: 0, operator: 'greaterThanOrEqual', value: '2' }],
                visibility: 'document',
              },
            ],
            objects: [
              {
                id: 'spanning',
                kind: 'text-box',
                anchor: {
                  type: 'two-cell',
                  from: {
                    sheetId: 'sheet-1',
                    row: 2,
                    column: 0,
                    offset: { x: 0, y: 0 },
                  },
                  to: {
                    sheetId: 'sheet-1',
                    row: 4,
                    column: 0,
                    offset: { x: 40, y: 0 },
                  },
                },
                zIndex: 1,
                locked: false,
                templateRepeat: 'shared',
                text: 'Spanning object',
                style: { color: '#111111', fontFamily: 'Arial', fontSize: 10 },
                accessibility: { name: 'Spanning object' },
              },
            ],
          },
        ],
      },
    } as unknown as SpreadsheetDocument;
    const viewTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          targets: [
            {
              type: 'range',
              range: {
                sheetId: 'sheet-1' as never,
                start: { row: 0, column: 0 },
                end: { row: 4, column: 0 },
              },
            },
          ],
          page: {
            ...template.printProfiles[0]!.page,
            paper: { type: 'custom', width: 240, height: 60 },
          },
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(objectSource, viewTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
        activeFilterViews: [{ sheetId: 'sheet-1' as never, viewId: 'descending' }],
      },
      environment,
    );
    const spanningCommands =
      result.document?.print.displayList.pages.flatMap((page, pageIndex) =>
        flattenCommands(page.commands).flatMap((command) =>
          command.kind === 'text' && command.text === 'Spanning object'
            ? [{ pageIndex, y: command.y, maxWidth: command.maxWidth }]
            : [],
        ),
      ) ?? [];

    expect(spanningCommands).toEqual([
      { pageIndex: 0, y: 40, maxWidth: 40 },
      { pageIndex: 1, y: 0, maxWidth: 40 },
    ]);
  });

  it('rejects stale compiled sources without partial output', async () => {
    const compiled = compileSpreadsheetTemplate(source, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: 'changed',
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.document).toBeUndefined();
    expect(result.diagnostics).toEqual([
      expect.objectContaining({ code: 'TEMPLATE_SOURCE_STALE' }),
    ]);
  });

  it('expands rows, translates relative formulas, removes false conditions, and paginates', async () => {
    const compiled = compileSpreadsheetTemplate(source, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {
          customer: { name: 'Ada' },
          items: [{ name: 'A' }, { name: 'B' }, { name: 'C' }],
          showNote: false,
        },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.workbook.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          row: 0,
          column: 1,
          cell: { input: { type: 'string', value: 'Ada' } },
        }),
        expect.objectContaining({
          row: 2,
          column: 1,
          cell: { input: { type: 'formula', source: '=A3' } },
        }),
        expect.objectContaining({
          row: 3,
          column: 1,
          cell: { input: { type: 'formula', source: '=A4' } },
        }),
      ]),
    );
    expect(
      result.document?.workbook.sheets[0]?.cells.some(
        ({ cell }) => cell.input.type === 'string' && cell.input.value === 'conditional',
      ),
    ).toBe(false);
    expect(result.document?.print.pages.length).toBeGreaterThan(0);
    expect(result.document?.print.displayList.pages).toHaveLength(
      result.document?.print.pages.length ?? 0,
    );
    expect(result.document?.print.profile).toEqual(template.printProfiles[0]);
    expect(result.document?.worksheets).toEqual([
      {
        sheetId: 'sheet-1',
        visibility: 'hidden',
        conditionalFormatting: [
          expect.objectContaining({
            type: 'cell-is',
            range: {
              sheetId: 'sheet-1',
              start: { row: 1, column: 0 },
              end: { row: 1, column: 1 },
            },
            formula: 'A2',
          }),
        ],
      },
    ]);
    expect(result.document?.calculatedCells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          address: { sheetId: 'sheet-1', row: 2, column: 1 },
          value: { type: 'string', value: 'item' },
        }),
      ]),
    );
  });

  it('supports abort and expansion limits atomically', async () => {
    const compiled = compileSpreadsheetTemplate(source, template).template!;
    const controller = new AbortController();
    controller.abort();
    const aborted = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [] },
        profileId: 'profile-1',
        missingValue: 'error',
        signal: controller.signal,
      },
      environment,
    );
    expect(aborted.document).toBeUndefined();
    expect(aborted.diagnostics[0]).toMatchObject({ code: 'RENDER_ABORTED' });

    const limited = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {
          customer: { name: 'A' },
          items: Array.from({ length: 10 }, () => ({})),
          showNote: true,
        },
        profileId: 'profile-1',
        missingValue: 'error',
        limits: { maxExpandedRows: 2 },
      },
      environment,
    );
    expect(limited.document).toBeUndefined();
    expect(limited.diagnostics).toEqual([
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    ]);
  });

  it('materializes repeat bindings on blank cells and applies the declared formatter', async () => {
    const repeatTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'lines' as never,
          type: 'repeat-rows',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 1, column: 0 },
            end: { row: 1, column: 2 },
          },
          source: 'items',
          empty: 'remove',
          pageBreak: 'auto',
        },
        {
          id: 'item-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 1, column: 2 },
          expression: 'item.name',
          formatter: 'upper',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, repeatTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [{ name: 'alpha' }, { name: 'beta' }] },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      {
        ...environment,
        formatters: { upper: (value) => String(value).toUpperCase() },
      },
    );
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.workbook.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          row: 1,
          column: 2,
          cell: { input: { type: 'string', value: 'ALPHA' } },
        }),
        expect.objectContaining({
          row: 2,
          column: 2,
          cell: { input: { type: 'string', value: 'BETA' } },
        }),
      ]),
    );
  });

  it('reports missing repeat values and enforces the cell budget before expansion', async () => {
    const repeatTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'lines' as never,
          type: 'repeat-rows',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 1, column: 0 },
            end: { row: 1, column: 2 },
          },
          source: 'items',
          empty: 'remove',
          pageBreak: 'auto',
        },
        {
          id: 'item-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 1, column: 2 },
          expression: 'item.name',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, repeatTemplate).template!;
    const missing = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [{}] },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(missing.document).toBeUndefined();
    expect(missing.diagnostics).toEqual([
      expect.objectContaining({ code: 'MISSING_DATA', location: { bindingId: 'item-name' } }),
    ]);

    const limited = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: Array.from({ length: 100 }, () => ({ name: 'x' })) },
        profileId: 'profile-1',
        missingValue: 'error',
        limits: { maxExpandedCells: 5 },
      },
      environment,
    );
    expect(limited.document).toBeUndefined();
    expect(limited.diagnostics).toEqual([
      expect.objectContaining({ code: 'EXPANSION_LIMIT_EXCEEDED' }),
    ]);
  });

  it('starts every repeated item on a new page when requested', async () => {
    const repeat = template.bindings.find(
      (
        binding,
      ): binding is Extract<SpreadsheetTemplate['bindings'][number], { type: 'repeat-rows' }> =>
        binding.type === 'repeat-rows',
    )!;
    const pageBreakTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          ...repeat,
          pageBreak: 'before-each-item',
        },
      ],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          page: {
            ...template.printProfiles[0]!.page,
            paper: { type: 'custom', width: 400, height: 400 },
          },
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, pageBreakTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [{}, {}, {}] },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.print.pages).toHaveLength(3);
  });

  it('keeps the untouched template row for an empty keep-template-row repeat', async () => {
    const keepTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'lines' as never,
          type: 'repeat-rows',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 1, column: 0 },
            end: { row: 1, column: 1 },
          },
          source: 'items',
          empty: 'keep-template-row',
          pageBreak: 'auto',
        },
        {
          id: 'item-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 1, column: 0 },
          expression: 'item.name',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, keepTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [] },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.workbook.sheets[0]?.cells).toContainEqual(
      expect.objectContaining({
        row: 1,
        column: 0,
        cell: { input: { type: 'string', value: 'item' } },
      }),
    );
  });

  it('maps scalar targets after conditional row removal', async () => {
    const mappedTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'hide-first' as never,
          type: 'conditional-range',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 0, column: 0 },
            end: { row: 0, column: 1 },
          },
          when: 'visible',
        },
        {
          id: 'mapped-value' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 1, column: 0 },
          expression: 'replacement',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, mappedTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { visible: false, replacement: 'mapped' },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.document?.workbook.sheets[0]?.cells).toContainEqual(
      expect.objectContaining({
        row: 0,
        column: 0,
        cell: { input: { type: 'string', value: 'mapped' } },
      }),
    );
  });

  it('does not execute formatter accessors and returns an atomic diagnostic', async () => {
    const formatterTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'formatted-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 0, column: 1 },
          expression: 'customer.name',
          formatter: 'danger',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, formatterTemplate).template!;
    const formatters = Object.defineProperty({}, 'danger', {
      enumerable: true,
      get() {
        throw new Error('must not execute');
      },
    });
    await expect(
      renderSpreadsheetTemplate(
        {
          template: compiled,
          currentDocumentHash: compiled.sourceDocumentHash,
          data: { customer: { name: 'Ada' } },
          profileId: 'profile-1',
          missingValue: 'error',
        },
        { ...environment, formatters },
      ),
    ).resolves.toEqual({
      diagnostics: [expect.objectContaining({ code: 'UNKNOWN_FORMATTER' })],
    });
  });

  it('renders merged geometry, titles, headings, bands, and gridline policy into the display list', async () => {
    const displaySource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            merges: [{ start: { row: 0, column: 0 }, end: { row: 0, column: 1 } }],
          },
        ],
      },
    } as SpreadsheetDocument;
    const displayTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          targets: [
            {
              type: 'range',
              range: {
                sheetId: 'sheet-1' as never,
                start: { row: 0, column: 0 },
                end: { row: 3, column: 1 },
              },
            },
          ],
          page: {
            ...template.printProfiles[0]!.page,
            paper: { type: 'custom', width: 260, height: 100 },
          },
          repeatRows: {
            sheetId: 'sheet-1' as never,
            start: { row: 0, column: 0 },
            end: { row: 0, column: 1 },
          },
          header: { left: 'Invoice {{customer.name}}' },
          footer: { center: 'Page {{page}}/{{pages}} {{date}}' },
          showGridlines: false,
          showHeadings: true,
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(displaySource, displayTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { customer: { name: 'Ada' } },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.print.pages.length).toBeGreaterThan(1);
    const pageCommands =
      result.document?.print.displayList.pages.map(({ commands }) => commands) ?? [];
    const texts = pageCommands.map((commands) =>
      commands.filter((command) => command.kind === 'text').map((command) => command.text),
    );
    expect(texts[0]).toEqual(
      expect.arrayContaining(['Invoice Ada', 'Page 1/2 2026-01-01', 'A', 'B', '1', 'Name']),
    );
    expect(texts[1]).toContain('Name');
    expect(pageCommands.flat().some((command) => command.kind === 'stroke-rect')).toBe(false);
    expect(
      pageCommands[0]?.some(
        (command) => command.kind === 'fill-rect' && command.rect.width === 178,
      ),
    ).toBe(true);
  });

  it('maps expanded sheet targets and manual breaks exactly once', async () => {
    const mappingTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [template.bindings.find(({ id }) => id === ('lines' as never))!],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          targets: [{ type: 'sheet', sheetId: 'sheet-1' as never }],
          manualBreaks: [{ sheetId: 'sheet-1' as never, beforeRow: 3 }],
          page: {
            ...template.printProfiles[0]!.page,
            paper: { type: 'custom', width: 400, height: 400 },
          },
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, mappingTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: { items: [{}, {}, {}] },
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.print.pages.map(({ rowStart, rowEnd }) => [rowStart, rowEnd])).toEqual([
      [0, 4],
      [5, 5],
    ]);
  });

  it('renders a stable continuation of a merge split across pages', async () => {
    const mergeSource = {
      ...source,
      workbook: {
        ...source.workbook,
        sheets: [
          {
            ...source.workbook.sheets[0]!,
            cells: [
              { row: 0, column: 0, cell: { input: { type: 'string', value: 'Merged title' } } },
            ],
            merges: [{ start: { row: 0, column: 0 }, end: { row: 1, column: 0 } }],
            rows: [
              { index: 0, height: 60 },
              { index: 1, height: 60 },
            ],
            columns: [{ index: 0, width: 100 }],
          },
        ],
      },
    } as SpreadsheetDocument;
    const mergeTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [],
      printProfiles: [
        {
          ...template.printProfiles[0]!,
          targets: [
            {
              type: 'range',
              range: {
                sheetId: 'sheet-1' as never,
                start: { row: 0, column: 0 },
                end: { row: 1, column: 0 },
              },
            },
          ],
          page: {
            ...template.printProfiles[0]!.page,
            paper: { type: 'custom', width: 140, height: 100 },
          },
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(mergeSource, mergeTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
      },
      environment,
    );
    expect(result.document?.print.pages).toHaveLength(2);
    const secondPageTexts =
      result.document?.print.displayList.pages[1]?.commands
        .filter((command) => command.kind === 'text')
        .map(({ text }) => text) ?? [];
    expect(secondPageTexts).toContain('Merged title');
  });

  it('returns an atomic diagnostic when a formatter is not registered', async () => {
    const formatterTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'formatted-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 0, column: 1 },
          expression: 'missingFormatter(customer.name)',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, formatterTemplate).template!;

    await expect(
      renderSpreadsheetTemplate(
        {
          template: compiled,
          currentDocumentHash: compiled.sourceDocumentHash,
          data: { customer: { name: 'Ada' } },
          profileId: 'profile-1',
          missingValue: 'error',
        },
        environment,
      ),
    ).resolves.toEqual({
      diagnostics: [
        expect.objectContaining({
          code: 'UNKNOWN_FORMATTER',
          severity: 'error',
          stage: 'render',
        }),
      ],
    });
  });

  it('returns an atomic diagnostic when a registered formatter throws', async () => {
    const formatterTemplate: SpreadsheetTemplate = {
      ...template,
      bindings: [
        {
          id: 'formatted-name' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 0, column: 1 },
          expression: 'explode(customer.name)',
        },
      ],
    };
    const compiled = compileSpreadsheetTemplate(source, formatterTemplate).template!;

    await expect(
      renderSpreadsheetTemplate(
        {
          template: compiled,
          currentDocumentHash: compiled.sourceDocumentHash,
          data: { customer: { name: 'Ada' } },
          profileId: 'profile-1',
          missingValue: 'error',
        },
        {
          ...environment,
          formatters: {
            explode() {
              throw new Error('formatter implementation detail');
            },
          },
        },
      ),
    ).resolves.toEqual({
      diagnostics: [
        expect.objectContaining({
          code: 'FORMATTER_FAILED',
          severity: 'error',
          stage: 'render',
          message: 'A template formatter failed during rendering',
        }),
      ],
    });
  });
});
