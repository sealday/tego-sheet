import { afterEach, expect, it, vi } from 'vitest';
import { SpreadsheetDocumentController } from '../../src/core/controller/spreadsheet-document-controller';
import { testDocument } from '../helpers/workbook-builders';
import { createSelectionState } from '../../src/engine';
import { createEngineAdapter } from '../../src/react/adapters/engine-adapter';
import { createEventDispatcher } from '../../src/react/adapters/event-dispatcher';
import { createCanvasHarness } from '../helpers/canvas-harness';
import { createFontMetrics } from '../../src/presentation';
import type { SpreadsheetDocument } from '../../src/document';

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

it('stages the target selection before callbacks and paints only that target snapshot', () => {
  const frames: FrameRequestCallback[] = [];
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    frames.push(callback);
    return frames.length;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  const root = document.createElement('div');
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 500 },
    clientHeight: { configurable: true, value: 300 },
  });
  const canvas = createCanvasHarness();
  const controller = new SpreadsheetDocumentController(
    testDocument({ rows: { len: 2 }, cols: { len: 3 } }),
  );
  const sheet = controller.getSheetIds()[0]!;
  const engine = createEngineAdapter({
    root,
    canvas: canvas.canvas as unknown as HTMLCanvasElement,
  });
  engine.render(controller.getSnapshot(), sheet);
  frames.shift()!(0);
  canvas.operations.length = 0;
  const targetState = createSelectionState({ row: 0, column: 1 });
  const target = { sheet, range: targetState.range, active: targetState.active };
  const order: string[] = [];
  const dispatcher = createEventDispatcher({
    controller,
    getCallbacks: () => ({
      onDocumentChange: () => order.push('change'),
      onCellEdit: () => order.push('cell-edit'),
      onSelectionChange: () => {
        expect(engine.publicSelection()?.active.column).toBe(1);
        order.push('selection');
      },
    }),
    schedulePaint: () => {
      expect(engine.publicSelection()?.active.column).toBe(1);
      order.push('paint');
      engine.render(controller.getSnapshot(), sheet);
    },
  });

  dispatcher.dispatchUi(
    { type: 'set-cell-text', address: { sheet, row: 0, column: 0 }, text: 'committed' },
    'keyboard',
    {
      selectionAfterCommit: target,
      beforeSelectionNotify: () => {
        engine.stageSelection(targetState);
      },
    },
  );
  frames.shift()!(1);

  expect(order).toEqual(['change', 'cell-edit', 'selection', 'paint']);
  const selectionStrokes = canvas.operations.filter((operation) => operation.name === 'strokeRect');
  expect(selectionStrokes.at(-1)?.args.slice(0, 4)).toEqual([160, 25, 100, 25]);
  engine.dispose();
});

it('@parity:view.scroll-sync stages offscreen selection and scroll together without scheduling an early paint', () => {
  const frames: FrameRequestCallback[] = [];
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    frames.push(callback);
    return frames.length;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  const root = document.createElement('div');
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 180 },
    clientHeight: { configurable: true, value: 100 },
  });
  const canvas = createCanvasHarness();
  const controller = new SpreadsheetDocumentController(
    testDocument({ rows: { len: 2 }, cols: { len: 3 } }),
  );
  const sheet = controller.getSheetIds()[0]!;
  const engine = createEngineAdapter({
    root,
    canvas: canvas.canvas as unknown as HTMLCanvasElement,
  });
  engine.render(controller.getSnapshot(), sheet);
  frames.shift()!(0);
  canvas.operations.length = 0;

  const target = createSelectionState({ row: 0, column: 2 });
  engine.stageSelection(target);

  expect(engine.publicSelection()?.active).toEqual({ row: 0, column: 2 });
  expect(engine.interactionSnapshot()!.viewport.scroll.x).toBeGreaterThan(0);
  expect(frames).toEqual([]);
  expect(canvas.operations).toEqual([]);
  engine.dispose();
});

it('suppresses repeated paints for a failed controller snapshot and retries on explicit layout recovery', () => {
  const frames: FrameRequestCallback[] = [];
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    frames.push(callback);
    return frames.length;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  const root = document.createElement('div');
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 180 },
    clientHeight: { configurable: true, value: 100 },
  });
  const canvas = createCanvasHarness();
  const context = canvas.canvas.getContext('2d')!;
  const failure = new Error('paint failed');
  Object.defineProperty(context, 'clearRect', {
    configurable: true,
    value: () => {
      throw failure;
    },
  });
  const controller = new SpreadsheetDocumentController(
    testDocument({ rows: { len: 2 }, cols: { len: 3 } }),
  );
  const sheet = controller.getSheetIds()[0]!;
  const onRenderError = vi.fn();
  const engine = createEngineAdapter({
    root,
    canvas: canvas.canvas as unknown as HTMLCanvasElement,
    onRenderError,
  });
  const snapshot = controller.getSnapshot();

  engine.render(snapshot, sheet);
  expect(() => frames.shift()!(0)).not.toThrow();
  expect(onRenderError).toHaveBeenCalledOnce();
  expect(onRenderError).toHaveBeenCalledWith(failure);

  engine.setSelection(createSelectionState({ row: 0, column: 1 }));
  engine.render(snapshot, sheet);
  expect(frames).toEqual([]);

  Object.defineProperty(context, 'clearRect', {
    configurable: true,
    value: vi.fn(),
  });
  engine.recalculateLayout();
  expect(frames).toHaveLength(1);
  expect(() => frames.shift()!(1)).not.toThrow();
  expect(onRenderError).toHaveBeenCalledOnce();
  engine.dispose();
});

it('paints template value, repeat, print, and invalid decorations with canvas geometry', () => {
  const frames: FrameRequestCallback[] = [];
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    frames.push(callback);
    return frames.length;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  const root = document.createElement('div');
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 500 },
    clientHeight: { configurable: true, value: 300 },
  });
  const canvas = createCanvasHarness();
  const controller = new SpreadsheetDocumentController(
    testDocument({ rows: { len: 3 }, cols: { len: 3 } }),
  );
  const sheet = controller.getSheetIds()[0]!;
  const engine = createEngineAdapter({
    root,
    canvas: canvas.canvas as unknown as HTMLCanvasElement,
  });
  engine.render(controller.getSnapshot(), sheet);
  frames.shift()!(0);
  canvas.operations.length = 0;
  engine.updateTemplateDecorations([
    {
      range: { start: { row: 0, column: 0 }, end: { row: 0, column: 0 } },
      kind: 'value',
      label: 'customer',
    },
    {
      range: { start: { row: 1, column: 0 }, end: { row: 2, column: 1 } },
      kind: 'repeat',
      label: 'lines',
      invalid: true,
    },
    {
      range: { start: { row: 0, column: 0 }, end: { row: 2, column: 2 } },
      kind: 'print',
      label: 'A4',
    },
  ]);
  frames.shift()!(1);

  expect(
    canvas.operations.filter(({ name }) => name === 'fillText').map(({ args }) => args[0]),
  ).toEqual(expect.arrayContaining(['customer', 'lines', 'A4']));
  expect(
    canvas.operations.filter(({ name }) => name === 'set:strokeStyle').map(({ args }) => args[0]),
  ).toEqual(expect.arrayContaining(['#2563eb', '#dc2626', '#059669']));
  engine.dispose();
});

it('resolves a data URL image through the resource pipeline and repaints with drawImage', async () => {
  const frames: FrameRequestCallback[] = [];
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    frames.push(callback);
    return frames.length;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  const root = document.createElement('div');
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 500 },
    clientHeight: { configurable: true, value: 300 },
  });
  const canvas = createCanvasHarness();
  const base = testDocument({ rows: { len: 2 }, cols: { len: 3 } });
  const png = new Uint8Array(45);
  png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  new DataView(png.buffer).setUint32(8, 13);
  png.set(new TextEncoder().encode('IHDR'), 12);
  new DataView(png.buffer).setUint32(16, 1);
  new DataView(png.buffer).setUint32(20, 1);
  png.set(new TextEncoder().encode('IEND'), 37);
  const documentWithImage = {
    ...base,
    workbook: {
      ...base.workbook,
      sheets: [
        {
          ...base.workbook.sheets[0]!,
          objects: [
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
          ],
        },
      ],
    },
    resources: {
      items: [
        {
          id: 'logo-resource',
          kind: 'image',
          mimeType: 'image/png',
          url: `data:image/png;base64,${btoa(String.fromCharCode(...png))}`,
        },
      ],
    },
  } as unknown as SpreadsheetDocument;
  const controller = new SpreadsheetDocumentController(documentWithImage);
  const sheet = controller.getSheetIds()[0]!;
  const source = {};
  const onObjectDiagnostics = vi.fn();
  const engine = createEngineAdapter({
    root,
    canvas: canvas.canvas as unknown as HTMLCanvasElement,
    onObjectDiagnostics,
    renderEnvironment: {
      locale: 'en-US',
      timeZone: 'UTC',
      dateSystem: 'excel-1900',
      clock: new Date('2026-07-24T00:00:00.000Z'),
      fontMetrics: createFontMetrics({
        fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
        fallbackFont: 'Arial',
        fallback: { averageAdvance: 6, lineHeight: 12 },
      }),
      decodeImage: async () => ({ width: 1, height: 1, representation: source }),
    },
  });

  engine.render(controller.getSnapshot(), sheet);
  frames.shift()!(0);
  expect(onObjectDiagnostics).toHaveBeenCalledWith([
    expect.objectContaining({ code: 'OBJECT_RESOURCE_MISSING' }),
  ]);
  await vi.waitFor(() => expect(frames.length).toBeGreaterThan(0));
  frames.shift()!(1);

  expect(canvas.operations).toContainEqual({
    name: 'drawImage',
    args: [source, 75, 45, 30, 30],
  });
  expect(onObjectDiagnostics).toHaveBeenLastCalledWith([]);
  engine.dispose();
});

it('reports identical object diagnostics again when the active sheet changes', () => {
  const frames: FrameRequestCallback[] = [];
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    frames.push(callback);
    return frames.length;
  });
  vi.stubGlobal('cancelAnimationFrame', vi.fn());
  const root = document.createElement('div');
  Object.defineProperties(root, {
    clientWidth: { configurable: true, value: 500 },
    clientHeight: { configurable: true, value: 300 },
  });
  const canvas = createCanvasHarness();
  const base = testDocument({ rows: { len: 2 }, cols: { len: 3 } });
  const object = {
    id: 'shared-object-id',
    kind: 'image',
    anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 40, height: 30 } },
    zIndex: 1,
    locked: false,
    templateRepeat: 'shared',
    resourceId: 'shared-resource-id',
    fit: 'contain',
    accessibility: { name: 'Unresolved' },
  } as const;
  const first = { ...base.workbook.sheets[0]!, objects: [object] };
  const second = { ...first, id: 'second-sheet', name: 'Second' };
  const controller = new SpreadsheetDocumentController({
    ...base,
    workbook: { ...base.workbook, sheets: [first, second] },
    resources: {
      items: [
        {
          id: 'shared-resource-id',
          kind: 'image',
          mimeType: 'image/png',
          url: 'https://example.test/not-fetched.png',
        },
      ],
    },
  } as unknown as SpreadsheetDocument);
  const [firstSheet, secondSheet] = controller.getSheetIds();
  const onObjectDiagnostics = vi.fn();
  const engine = createEngineAdapter({
    root,
    canvas: canvas.canvas as unknown as HTMLCanvasElement,
    onObjectDiagnostics,
  });

  engine.render(controller.getSnapshot(), firstSheet!);
  frames.shift()!(0);
  engine.render(controller.getSnapshot(), secondSheet!);
  frames.shift()!(1);

  const reportedSheets = onObjectDiagnostics.mock.calls
    .filter(([diagnostics]) => diagnostics.length > 0)
    .map(([diagnostics]) => diagnostics[0].location.sheetId);
  expect(reportedSheets).toEqual([firstSheet, secondSheet]);
  engine.dispose();
});
