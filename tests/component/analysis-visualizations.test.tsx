import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { parseSpreadsheetDocument, TegoSheet, type TegoSheetHandle } from '../../src';
import type {
  DocumentSheetId,
  SpreadsheetDocument,
  SpreadsheetDocumentInput,
} from '../../src/document';
import { createFontMetrics } from '../../src/presentation';
import type { PrintDisplayCommand } from '../../src/print';
import { compileSpreadsheetTemplate } from '../../src/template/compiler';
import { renderSpreadsheetTemplate } from '../../src/template/render';
import type { SpreadsheetTemplate } from '../../src/template/model';
import { createCanvasHarness } from '../helpers/canvas-harness';
import { CanvasEngine, type CanvasRenderSnapshot } from '../../src/engine';

let canvasHarness: ReturnType<typeof createCanvasHarness>;

const sourceInput: SpreadsheetDocumentInput = {
  schemaVersion: 2,
  id: 'analysis-visualizations',
  workbook: {
    sheets: [
      {
        id: 'sheet-1',
        name: 'Dashboard',
        rowCount: 12,
        columnCount: 6,
        cells: [
          { row: 0, column: 0, cell: { input: { type: 'number', value: 4 } } },
          { row: 0, column: 1, cell: { input: { type: 'number', value: 9 } } },
          { row: 0, column: 2, cell: { input: { type: 'formula', source: '=A1+B1' } } },
        ],
        rows: [],
        columns: [
          { index: 0, width: 80 },
          { index: 1, width: 80 },
          { index: 2, width: 80 },
          { index: 3, width: 80 },
        ],
        merges: [],
        charts: [
          {
            id: 'revenue',
            type: 'line',
            title: 'Revenue',
            series: [
              {
                id: 'actual',
                values: {
                  sheetId: 'sheet-1' as DocumentSheetId,
                  start: { row: 0, column: 0 },
                  end: { row: 0, column: 2 },
                },
              },
            ],
            anchor: {
              type: 'one-cell',
              cell: { sheetId: 'sheet-1' as DocumentSheetId, row: 2, column: 0 },
              offset: { x: 4, y: 4 },
              size: { width: 220, height: 110 },
            },
          },
        ],
        sparklines: [
          {
            id: 'revenue-trend',
            type: 'column',
            source: {
              sheetId: 'sheet-1' as DocumentSheetId,
              start: { row: 0, column: 0 },
              end: { row: 0, column: 2 },
            },
            target: { sheetId: 'sheet-1' as DocumentSheetId, row: 1, column: 3 },
          },
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
};

const template: SpreadsheetTemplate = {
  id: 'template-1' as never,
  name: 'Dashboard',
  bindings: [],
  printProfiles: [
    {
      id: 'profile-1',
      name: 'Dashboard',
      targets: [
        {
          type: 'range',
          range: {
            sheetId: 'sheet-1' as never,
            start: { row: 0, column: 0 },
            end: { row: 8, column: 4 },
          },
        },
      ],
      page: {
        paper: { type: 'custom', width: 500, height: 400 },
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

function sourceDocument(): SpreadsheetDocument {
  const result = parseSpreadsheetDocument(structuredClone(sourceInput));
  if (!result.ok) throw new Error(JSON.stringify(result.diagnostics));
  return result.document;
}

function filteredVisualizationDocument(): SpreadsheetDocument {
  const input = structuredClone(sourceInput);
  const sheet = input.workbook.sheets[0]!;
  sheet.cells = ['header', 1, 2, 3, 4].map((value, row) => ({
    row,
    column: 0,
    cell:
      typeof value === 'number'
        ? { input: { type: 'number' as const, value } }
        : { input: { type: 'string' as const, value } },
  }));
  sheet.filterViews = [
    {
      id: 'descending',
      name: 'Descending',
      range: {
        sheetId: 'sheet-1' as DocumentSheetId,
        start: { row: 0, column: 0 },
        end: { row: 4, column: 0 },
      },
      sorts: [{ column: 0, direction: 'descending' }],
      filters: [{ column: 0, operator: 'greaterThanOrEqual', value: 2 }],
      visibility: 'document',
    },
  ];
  const values = {
    sheetId: 'sheet-1' as DocumentSheetId,
    start: { row: 1, column: 0 },
    end: { row: 4, column: 0 },
  };
  sheet.charts = [
    {
      id: 'sorted-chart',
      type: 'line',
      title: 'Sorted chart',
      series: [{ id: 'values', values }],
      anchor: {
        type: 'one-cell',
        cell: { sheetId: 'sheet-1' as DocumentSheetId, row: 4, column: 0 },
        offset: { x: 0, y: 0 },
        size: { width: 160, height: 60 },
      },
    },
    {
      id: 'filtered-chart',
      type: 'line',
      title: 'Filtered chart',
      series: [{ id: 'values', values }],
      anchor: {
        type: 'one-cell',
        cell: { sheetId: 'sheet-1' as DocumentSheetId, row: 1, column: 0 },
        offset: { x: 0, y: 0 },
        size: { width: 160, height: 60 },
      },
    },
    {
      id: 'spanning-chart',
      type: 'line',
      title: 'Spanning chart',
      series: [{ id: 'values', values }],
      anchor: {
        type: 'two-cell',
        from: {
          sheetId: 'sheet-1' as DocumentSheetId,
          row: 2,
          column: 0,
          offset: { x: 0, y: 0 },
        },
        to: {
          sheetId: 'sheet-1' as DocumentSheetId,
          row: 4,
          column: 2,
          offset: { x: 0, y: 0 },
        },
      },
    },
  ];
  sheet.sparklines = [
    {
      id: 'sorted-sparkline',
      type: 'column',
      source: values,
      target: { sheetId: 'sheet-1' as DocumentSheetId, row: 3, column: 3 },
    },
  ];
  const parsed = parseSpreadsheetDocument(input);
  if (!parsed.ok) throw new Error(JSON.stringify(parsed.diagnostics));
  return parsed.document;
}

function flatten(commands: readonly PrintDisplayCommand[]): readonly PrintDisplayCommand[] {
  return commands.flatMap((command) =>
    command.kind === 'clip' || command.kind === 'group'
      ? [command, ...flatten(command.commands)]
      : [command],
  );
}

beforeEach(() => {
  canvasHarness = createCanvasHarness();
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    () => canvasHarness.canvas.getContext('2d') as CanvasRenderingContext2D,
  );
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    queueMicrotask(() => callback(1));
    return 1;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
});

afterEach(() => {
  cleanup();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe('persisted analysis visualization production', () => {
  it('mounts persisted charts and sparklines as deterministic, accessible TegoSheet overlays', async () => {
    const document = sourceDocument();
    const rendered = render(<TegoSheet defaultDocument={document} />);
    const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
    Object.defineProperties(root, {
      clientWidth: { configurable: true, value: 500 },
      clientHeight: { configurable: true, value: 300 },
    });

    act(() => fireEvent(window, new Event('resize')));

    expect(
      await rendered.findByRole('img', {
        name: /Revenue.*4.*9.*13.*revenue-trend.*4.*9.*13/,
      }),
    ).not.toBeNull();
    await waitFor(() => {
      expect(canvasHarness.operations).toContainEqual({
        name: 'fillText',
        args: ['Revenue', expect.any(Number), expect.any(Number), expect.any(Number)],
      });
    });
  });

  it('feeds persisted chart and sparkline commands into the template print display list', async () => {
    const document = sourceDocument();
    const compiled = compileSpreadsheetTemplate(document, template).template;
    if (compiled === undefined) throw new Error('template must compile');

    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
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

    expect(result.diagnostics).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ severity: 'error' })]),
    );
    const commands = flatten(result.document?.print.displayList.pages[0]?.commands ?? []);
    expect(commands).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ kind: 'text', text: 'Revenue' }),
        expect.objectContaining({ kind: 'path' }),
      ]),
    );
    expect(
      commands.filter(
        (command) =>
          command.kind === 'fill-rect' &&
          command.rect.x >= 250 &&
          command.rect.y >= 20 &&
          command.rect.height > 0,
      ).length,
    ).toBeGreaterThan(0);
  });

  it('maps screen chart and sparkline rows through an active sorted view', async () => {
    const ref = createRef<TegoSheetHandle>();
    const renderSpy = vi.spyOn(CanvasEngine.prototype, 'render');
    const rendered = render(
      <TegoSheet ref={ref} defaultDocument={filteredVisualizationDocument()} />,
    );
    const root = rendered.container.querySelector<HTMLElement>('[data-tego-sheet]')!;
    Object.defineProperties(root, {
      clientWidth: { configurable: true, value: 500 },
      clientHeight: { configurable: true, value: 300 },
    });
    act(() => fireEvent(window, new Event('resize')));
    await waitFor(() => expect(ref.current).not.toBeNull());

    act(() =>
      ref.current!.activateFilterView(
        ref.current!.getDocument().workbook.sheets[0]!.id as never,
        'descending',
      ),
    );
    let snapshot: CanvasRenderSnapshot | undefined;
    await waitFor(() => {
      snapshot = renderSpy.mock.calls.at(-1)?.[0];
      expect(snapshot?.visualizationCommands?.length).toBeGreaterThan(0);
    });
    const chartClips =
      snapshot?.visualizationCommands?.filter((command) => command.kind === 'clip') ?? [];
    const sparkBars =
      snapshot?.visualizationCommands?.filter(
        (command) => command.kind === 'fill-rect' && command.rect.width > 3,
      ) ?? [];

    expect(chartClips[1]).toMatchObject({ kind: 'clip', rect: { y: 25 } });
    expect(sparkBars).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: 'fill-rect',
          rect: expect.objectContaining({ y: expect.closeTo(52, 5) }),
        }),
      ]),
    );
  });

  it('uses object-equivalent filtered and two-cell chart projection in print', async () => {
    const document = filteredVisualizationDocument();
    const compiled = compileSpreadsheetTemplate(document, template).template;
    if (compiled === undefined) throw new Error('template must compile');
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile-1',
        missingValue: 'error',
        activeFilterViews: [{ sheetId: 'sheet-1' as never, viewId: 'descending' }],
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
    const commands = flatten(result.document?.print.displayList.pages[0]?.commands ?? []);
    const titles = commands.flatMap((command) =>
      command.kind === 'text' && command.text.endsWith('chart')
        ? [{ text: command.text, y: command.y }]
        : [],
    );
    const spanningClip = (
      result.document?.print.displayList.pages[0]?.commands.filter(
        (command) => command.kind === 'clip' && command.rect.width === 160,
      ) ?? []
    ).at(-1);

    expect(titles).toEqual([
      { text: 'Sorted chart', y: 48 },
      { text: 'Spanning chart', y: 48 },
    ]);
    expect(spanningClip).toMatchObject({
      kind: 'clip',
      rect: { y: 30, height: 40 },
    });
  });
});
