import { describe, expect, it, vi } from 'vitest';
import type { SpreadsheetDocument } from '../../../src/document';
import { createFontMetrics } from '../../../src/presentation';
import {
  compileSpreadsheetTemplate,
  createResourceResolverRegistry,
  renderSpreadsheetTemplate,
  type SpreadsheetTemplate,
} from '../../../src/template';

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

const source = {
  schemaVersion: 2,
  id: 'document-resource-render',
  workbook: {
    sheets: [{ id: 'sheet-1', name: 'Output', cells: [], merges: [], rows: [], columns: [] }],
    styles: [],
    validations: [],
    settings: { dateSystem: 'excel-1900' },
  },
  templates: [],
  resources: { items: [] },
  extensions: {},
} as unknown as SpreadsheetDocument;

const profile = {
  id: 'profile',
  name: 'Resource output',
  targets: [{ type: 'sheet', sheetId: 'sheet-1' }],
  page: {
    paper: { type: 'A4' },
    orientation: 'portrait',
    margins: { top: 10, right: 10, bottom: 10, left: 10 },
    scale: { type: 'fixed', value: 1 },
  },
  manualBreaks: [],
  showGridlines: false,
  showHeadings: false,
} as const;

const template = {
  id: 'template-resource' as never,
  name: 'Resource',
  bindings: [],
  printProfiles: [profile],
} as unknown as SpreadsheetTemplate;

const environment = {
  locale: 'zh-CN',
  timeZone: 'Asia/Shanghai',
  dateSystem: 'excel-1900' as const,
  clock: new Date('2026-07-23T00:00:00.000Z'),
  fontMetrics: createFontMetrics({
    fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
    fallbackFont: 'Arial',
    fallback: { averageAdvance: 6, lineHeight: 12 },
  }),
};

describe('resource/render atomic boundary', () => {
  it('publishes only a ready resource store and transfers disposal ownership', async () => {
    const dispose = vi.fn();
    const compiled = compileSpreadsheetTemplate(source, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile',
        missingValue: 'error',
        resourceRefs: [
          {
            id: 'logo',
            type: 'image',
            resolverId: 'app',
            key: 'logo',
            expectedMime: 'image/png',
          },
        ],
      },
      {
        ...environment,
        resourcePurpose: 'print',
        resourceRegistry: createResourceResolverRegistry([
          {
            id: 'app',
            supports: () => true,
            resolve: async () => ({
              bytes: png(),
              mimeType: 'image/png',
              width: 1,
              height: 1,
              dispose,
            }),
          },
        ]),
      },
    );
    expect(result.document?.resources.byReference.logo?.contentHash).toMatch(/^sha256:/u);
    expect(dispose).not.toHaveBeenCalled();
    await result.document?.resources.dispose();
    expect(dispose).toHaveBeenCalledOnce();
  });

  it('maps a template resource binding through the ready store into the print display list', async () => {
    const resourceTemplate = {
      ...template,
      resourceBindings: [
        {
          id: 'logo-binding',
          target: {
            sheetId: 'sheet-1',
            start: { row: 0, column: 0 },
            end: { row: 0, column: 0 },
          },
          resourceId: 'logo',
          fit: 'contain',
        },
      ],
    } as unknown as SpreadsheetTemplate;
    const compiled = compileSpreadsheetTemplate(source, resourceTemplate).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile',
        missingValue: 'error',
        resourceRefs: [
          {
            id: 'logo',
            type: 'image',
            resolverId: 'app',
            key: 'logo',
            expectedMime: 'image/png',
          },
        ],
      },
      {
        ...environment,
        resourceRegistry: createResourceResolverRegistry([
          {
            id: 'app',
            supports: () => true,
            resolve: async () => ({ bytes: png(), mimeType: 'image/png' }),
          },
        ]),
      },
    );

    expect(result.diagnostics).toEqual([]);
    expect(result.document?.print.displayList.pages[0]?.commands).toContainEqual(
      expect.objectContaining({
        kind: 'image',
        resourceId: 'logo',
        fit: 'contain',
        rect: { x: 10, y: 10, width: 100, height: 20 },
      }),
    );
    await result.document?.resources.dispose();
  });

  it('releases ready resources when a later render stage fails', async () => {
    const dispose = vi.fn();
    const failing = {
      ...template,
      bindings: [
        {
          id: 'missing' as never,
          type: 'value',
          target: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
          expression: 'missing',
        },
      ],
    } as SpreadsheetTemplate;
    const compiled = compileSpreadsheetTemplate(source, failing).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile',
        missingValue: 'error',
        resourceRefs: [{ id: 'binary', type: 'binary', resolverId: 'app', key: '42' }],
      },
      {
        ...environment,
        resourceRegistry: createResourceResolverRegistry([
          {
            id: 'app',
            supports: () => true,
            resolve: async () => ({
              bytes: new Uint8Array(),
              mimeType: 'application/octet-stream',
              dispose,
            }),
          },
        ]),
      },
    );
    expect(result.document).toBeUndefined();
    expect(result.diagnostics).toContainEqual(expect.objectContaining({ code: 'MISSING_DATA' }));
    expect(dispose).toHaveBeenCalledOnce();
  });

  it('waits for CJK font readiness before any pagination measurement', async () => {
    let ready = false;
    const metrics = {
      resolve: environment.fontMetrics.resolve,
      measure(text: string, family: string, size: number) {
        expect(ready).toBe(true);
        return environment.fontMetrics.measure(text, family, size);
      },
    };
    const compiled = compileSpreadsheetTemplate(source, template).template!;
    const result = await renderSpreadsheetTemplate(
      {
        template: compiled,
        currentDocumentHash: compiled.sourceDocumentHash,
        data: {},
        profileId: 'profile',
        missingValue: 'error',
        resourceRefs: [
          {
            id: 'cjk',
            type: 'font',
            resolverId: 'font',
            key: 'cjk',
            expectedMime: 'font/ttf',
          },
        ],
      },
      {
        ...environment,
        fontMetrics: metrics,
        resourceRegistry: createResourceResolverRegistry([
          {
            id: 'font',
            supports: () => true,
            resolve: async () => ({
              bytes: new Uint8Array([0, 1, 0, 0]),
              mimeType: 'font/ttf',
              font: {
                family: 'Noto Sans CJK',
                waitUntilReady: async () => {
                  ready = true;
                },
              },
            }),
          },
        ]),
      },
    );
    expect(ready).toBe(true);
    expect(result.document).toBeDefined();
  });
});
