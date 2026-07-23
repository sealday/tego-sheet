import type { GeneratedDocument } from '../../../src/template';

export function outputGeneratedDocument(): GeneratedDocument {
  return {
    workbook: {
      sheets: [
        {
          id: 'sheet-1',
          name: 'Invoice',
          cells: [
            {
              row: 0,
              column: 0,
              cell: { input: { type: 'string', value: '=literal' }, styleId: 'style-heading' },
            },
            {
              row: 1,
              column: 0,
              cell: { input: { type: 'number', value: 2 } },
            },
            {
              row: 1,
              column: 1,
              cell: { input: { type: 'formula', source: '=A2+1' }, validationId: 'positive' },
            },
            {
              row: 2,
              column: 0,
              cell: { input: { type: 'boolean', value: true } },
            },
          ],
          merges: [{ start: { row: 0, column: 0 }, end: { row: 0, column: 1 } }],
          rows: [{ index: 0, height: 24 }],
          columns: [
            { index: 0, width: 120 },
            { index: 1, width: 80, hidden: true },
          ],
          rowCount: 3,
          columnCount: 2,
        },
      ],
      styles: [
        {
          id: 'style-heading',
          value: {
            color: '#112233',
            backgroundColor: '#ddeeff',
            fontFamily: 'Noto Sans',
            fontSize: 14,
            bold: true,
            horizontalAlign: 'center',
            verticalAlign: 'middle',
            wrap: false,
            numberFormat: '0.00',
          },
        },
      ],
      validations: [
        {
          id: 'positive',
          value: {
            type: 'decimal',
            operator: 'greaterThan',
            formula1: 0,
            allowBlank: false,
          },
        },
      ],
      settings: { dateSystem: 'excel-1900', localeHint: 'zh-CN' },
    },
    print: {
      pages: [
        {
          id: 'invoice-1',
          index: 0,
          targetId: 'target-invoice',
          width: 210,
          height: 297,
          rowStart: 0,
          rowEnd: 1,
          columnStart: 0,
          columnEnd: 1,
        },
        {
          id: 'invoice-2',
          index: 1,
          targetId: 'target-invoice',
          width: 210,
          height: 297,
          rowStart: 2,
          rowEnd: 2,
          columnStart: 0,
          columnEnd: 1,
        },
      ],
      displayList: {
        diagnostics: [],
        pages: [
          {
            index: 0,
            width: 210,
            height: 297,
            commands: [
              {
                kind: 'fill-rect',
                rect: { x: 10, y: 10, width: 190, height: 20 },
                color: '#ddeeff',
              },
              {
                kind: 'stroke-rect',
                rect: { x: 10, y: 10, width: 190, height: 20 },
                color: '#112233',
                width: 0.5,
              },
              {
                kind: 'text',
                text: '中文 Invoice',
                x: 12,
                y: 24,
                maxWidth: 186,
                fontFamily: 'Noto Sans',
                fontSize: 12,
                color: '#112233',
                horizontalAlign: 'left',
              },
              {
                kind: 'link',
                rect: { x: 12, y: 12, width: 80, height: 14 },
                href: 'https://example.com/invoice',
                label: 'Invoice link',
              },
            ],
          },
          {
            index: 1,
            width: 210,
            height: 297,
            commands: [
              {
                kind: 'line',
                x1: 10,
                y1: 20,
                x2: 200,
                y2: 20,
                color: '#112233',
                width: 0.5,
              },
            ],
          },
        ],
      },
    },
    resources: {
      byHash: {},
      byReference: {},
      totalBytes: 0,
      dispose: async () => undefined,
    },
    objects: [],
    diagnostics: [],
    metadata: {
      templateId: 'template-output',
      profileId: 'invoice-print',
      sourceDocumentHash: 'sha256:output-fixture',
      locale: 'zh-CN',
      timeZone: 'Asia/Shanghai',
      generatedAt: '2026-07-23T00:00:00.000Z',
    },
  };
}
