import { describe, expect, it } from 'vitest';
import {
  createFontMetrics,
  createPresentationCache,
  createPresentationResolver,
  type CellPresentation,
} from '../../../src/presentation';
import { createPrintDisplayList } from '../../../src/print/display-list';
import { createFormulaEngine } from '../../../src/formula';
import { parseSpreadsheetDocument } from '../../../src/document';
import { CanvasEngine, createSheetGridModel, createViewportMetrics } from '../../../src/engine';
import { createCanvasHarness } from '../../helpers/canvas-harness';

function documentFixture() {
  const parsed = parseSpreadsheetDocument({
    schemaVersion: 2,
    id: 'presentation-document',
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'number', value: 1234.5 }, styleId: 'currency' },
            },
            {
              row: 0,
              column: 1,
              cell: { input: { type: 'formula', source: '=A1*2' } },
            },
            {
              row: 1,
              column: 0,
              cell: {
                input: { type: 'string', value: 'hidden row' },
                validationId: 'required',
                editable: false,
                printable: false,
              },
            },
          ],
          merges: [],
          rows: [{ index: 1, hidden: true }],
          columns: [],
        },
      ],
      styles: [
        {
          id: 'currency',
          value: {
            numberFormat: '$#,##0.00',
            color: '#123456',
            backgroundColor: '#ffeecc',
            fontFamily: 'Inter',
            fontSize: 12,
            bold: true,
            horizontalAlign: 'right',
          },
        },
      ],
      validations: [{ id: 'required', value: { message: 'Required value' } }],
      settings: { dateSystem: 'excel-1900', localeHint: 'en-US' },
    },
    templates: [],
    resources: { items: [] },
    extensions: {},
  });
  if (!parsed.ok) throw new Error('fixture must parse');
  return parsed.document;
}

function resolverFixture(maximumEntries = 100) {
  const document = documentFixture();
  const formulaEngine = createFormulaEngine();
  const program = formulaEngine.compile(document);
  formulaEngine.recalculate(program, [], {
    locale: 'en-US',
    timeZone: 'UTC',
    dateSystem: 'excel-1900',
    clock: { now: () => 0 },
    tick: 0,
    functionRegistryVersion: 'builtin-1',
  });
  const cache = createPresentationCache({ maximumEntries, maximumBytes: 16_384 });
  return {
    document,
    cache,
    resolver: createPresentationResolver({
      document,
      formulaProgram: program,
      cache,
      revisions: {
        document: 1,
        calculation: 1,
        condition: 0,
        style: 1,
        environment: 1,
      },
      environment: {
        locale: 'en-US',
        timeZone: 'UTC',
        dateSystem: 'excel-1900',
        target: 'screen',
      },
      validation: ({ row }) =>
        row === 1 ? { status: 'error', message: 'Required value' } : { status: 'valid' },
      annotations: ({ row, column }) =>
        row === 0 && column === 0 ? [{ kind: 'note', text: 'Quarterly revenue' }] : [],
    }),
  };
}

describe('shared cell presentation', () => {
  it('resolves typed value, formatted text, style, validation, annotations and visibility once', () => {
    const { resolver } = resolverFixture();

    expect(resolver.resolve({ sheetId: 'sheet-1' as never, row: 0, column: 0 })).toEqual(
      expect.objectContaining({
        value: { type: 'number', value: 1234.5 },
        formattedText: '$1,234.50',
        style: expect.objectContaining({
          color: '#123456',
          backgroundColor: '#ffeecc',
          fontFamily: 'Inter',
          fontSize: 12,
          bold: true,
          horizontalAlign: 'right',
        }),
        validation: { status: 'valid' },
        annotations: [{ kind: 'note', text: 'Quarterly revenue' }],
        visibility: { hidden: false, printable: true },
      }),
    );

    expect(resolver.resolve({ sheetId: 'sheet-1' as never, row: 0, column: 1 })).toEqual(
      expect.objectContaining({
        value: { type: 'number', value: 2469 },
        formattedText: '2469',
      }),
    );

    expect(resolver.resolve({ sheetId: 'sheet-1' as never, row: 1, column: 0 })).toEqual(
      expect.objectContaining({
        validation: { status: 'error', message: 'Required value' },
        visibility: { hidden: true, printable: false },
        accessibility: expect.objectContaining({
          readOnly: true,
          invalid: true,
          description: 'Required value',
        }),
      }),
    );
  });

  it('returns the same frozen presentation for the same revision tuple', () => {
    const { resolver } = resolverFixture();
    const address = { sheetId: 'sheet-1' as never, row: 0, column: 0 };

    const first = resolver.resolve(address);
    const second = resolver.resolve(address);

    expect(second).toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.style)).toBe(true);
  });

  it('enforces entry and byte budgets with deterministic least-recently-used eviction', () => {
    const { resolver, cache } = resolverFixture(2);
    const first = resolver.resolve({ sheetId: 'sheet-1' as never, row: 0, column: 0 });
    resolver.resolve({ sheetId: 'sheet-1' as never, row: 0, column: 1 });
    resolver.resolve({ sheetId: 'sheet-1' as never, row: 1, column: 0 });

    expect(cache.stats()).toMatchObject({ entries: 2, evictions: 1 });
    expect(resolver.resolve({ sheetId: 'sheet-1' as never, row: 0, column: 0 })).not.toBe(first);
    expect(cache.stats().bytes).toBeLessThanOrEqual(16_384);
  });
});

describe('device-independent print display list', () => {
  it('keeps Canvas and print text in parity by consuming the same presentation', () => {
    const { resolver } = resolverFixture();
    const harness = createCanvasHarness();
    const sheet = {
      rows: { len: 1, 0: { cells: { 0: { text: 'legacy text must not win' } } } },
      cols: { len: 1 },
    };
    const engine = new CanvasEngine(harness.canvas, {
      animationFrame: harness.animationFrame,
      measurement: harness.measurement,
    });
    const presentation = resolver.resolve({
      sheetId: 'sheet-1' as never,
      row: 0,
      column: 0,
    });
    engine.render({
      sheet,
      viewport: createViewportMetrics(createSheetGridModel(sheet), {
        width: 300,
        height: 200,
      }),
      presentations: {
        resolve: ({ row, column }) =>
          resolver.resolve({ sheetId: 'sheet-1' as never, row, column }),
      },
    });
    harness.animationFrame.flush();
    const canvasText = harness.operations.find(({ name }) => name === 'fillText')?.args[0];
    const print = createPrintDisplayList({
      pages: [
        {
          width: 300,
          height: 200,
          cells: [
            {
              rect: { x: 0, y: 0, width: 100, height: 20 },
              presentation,
            },
          ],
        },
      ],
      fontMetrics: createFontMetrics({
        fonts: { Inter: { averageAdvance: 6, lineHeight: 14 } },
        fallbackFont: 'Fallback',
        fallback: { averageAdvance: 7, lineHeight: 15 },
      }),
    });
    const printText = print.pages[0]!.commands.find(({ kind }) => kind === 'text');

    expect(canvasText).toBe(presentation.formattedText);
    expect(printText).toEqual(
      expect.objectContaining({ kind: 'text', text: presentation.formattedText }),
    );
    engine.dispose();
  });

  it('uses the exact shared formatted text and produces deterministic frozen commands', () => {
    const { resolver } = resolverFixture();
    const presentation = resolver.resolve({
      sheetId: 'sheet-1' as never,
      row: 0,
      column: 0,
    });
    const input = {
      pages: [
        {
          width: 600,
          height: 800,
          cells: [
            {
              rect: { x: 10, y: 20, width: 100, height: 24 },
              presentation,
            },
          ],
        },
      ],
      fontMetrics: createFontMetrics({
        fonts: { Inter: { averageAdvance: 6, lineHeight: 14 } },
        fallbackFont: 'Fallback',
        fallback: { averageAdvance: 7, lineHeight: 15 },
      }),
    } as const;

    const first = createPrintDisplayList(input);
    const second = createPrintDisplayList(input);

    expect(second).toEqual(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.pages[0]!.commands)).toBe(true);
    expect(first.pages[0]!.commands).toContainEqual(
      expect.objectContaining({ kind: 'text', text: presentation.formattedText }),
    );
    expect(JSON.stringify(first)).not.toMatch(/devicePixelRatio|scroll|selection|zoom/u);
  });

  it('reports a stable missing-font diagnostic and uses injected fallback metrics', () => {
    const presentation: CellPresentation = {
      address: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
      value: { type: 'string', value: 'hello' },
      formattedText: 'hello',
      style: {
        color: '#000000',
        backgroundColor: '#ffffff',
        fontFamily: 'Missing',
        fontSize: 10,
        bold: false,
        italic: false,
        horizontalAlign: 'left',
        verticalAlign: 'middle',
        wrap: false,
      },
      validation: { status: 'valid' },
      annotations: [],
      visibility: { hidden: false, printable: true },
      accessibility: { label: 'hello', readOnly: false, invalid: false },
    };
    const list = createPrintDisplayList({
      pages: [
        {
          width: 200,
          height: 200,
          cells: [{ rect: { x: 0, y: 0, width: 100, height: 20 }, presentation }],
        },
      ],
      fontMetrics: createFontMetrics({
        fonts: {},
        fallbackFont: 'Fallback',
        fallback: { averageAdvance: 7, lineHeight: 15 },
      }),
    });

    expect(list.diagnostics).toEqual([
      expect.objectContaining({
        code: 'PRESENTATION_FONT_MISSING',
        severity: 'warning',
        details: { requestedFont: 'Missing', fallbackFont: 'Fallback' },
      }),
    ]);
    expect(list.pages[0]!.commands).toContainEqual(
      expect.objectContaining({ kind: 'text', fontFamily: 'Fallback' }),
    );
  });
});
