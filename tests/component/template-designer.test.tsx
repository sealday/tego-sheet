import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import type { Diagnostic } from '../../src/document';
import { createSpreadsheetDocument } from '../../src/document';
import { TegoSheet } from '../../src/react/tego-sheet';
import { createFontMetrics } from '../../src/presentation';
import { TemplateDesigner } from '../../src/react/template-designer';
import { TemplatePreview } from '../../src/react/preview';
import type { GeneratedDocument, SpreadsheetTemplate } from '../../src/template';

const template: SpreadsheetTemplate = {
  id: 'template-1' as never,
  name: 'Invoice',
  bindings: [
    {
      id: 'binding-1' as never,
      type: 'value',
      target: { sheetId: 'sheet-1' as never, row: 0, column: 0 },
      expression: 'customer.name',
    },
  ],
  printProfiles: [
    {
      id: 'profile-1',
      name: 'A4',
      targets: [{ type: 'sheet', sheetId: 'sheet-1' as never }],
      page: {
        paper: { type: 'A4' },
        orientation: 'portrait',
        margins: { top: 20, right: 20, bottom: 20, left: 20 },
        scale: { type: 'fixed', value: 1 },
      },
      manualBreaks: [],
      showGridlines: true,
      showHeadings: false,
    },
  ],
};

describe('TemplateDesigner', () => {
  it('@parity:output.template-designer edits the SDK template model and locates binding diagnostics', () => {
    const onChange = vi.fn();
    const onLocate = vi.fn();
    const diagnostics: readonly Diagnostic[] = [
      {
        code: 'MISSING_DATA',
        severity: 'error',
        domain: 'template',
        stage: 'resolve',
        message: 'Customer is missing',
        location: { bindingId: 'binding-1' as never },
      },
    ];
    render(
      <TemplateDesigner
        template={template}
        diagnostics={diagnostics}
        onChange={onChange}
        onLocateBinding={onLocate}
        selection={{
          sheetId: 'sheet-1' as never,
          start: { row: 2, column: 1 },
          end: { row: 4, column: 2 },
        }}
      />,
    );

    fireEvent.change(screen.getByLabelText('Expression for binding-1'), {
      target: { value: 'customer.legalName' },
    });
    expect(onChange).toHaveBeenCalledWith(
      expect.objectContaining({
        bindings: [expect.objectContaining({ expression: 'customer.legalName' })],
      }),
    );
    fireEvent.click(screen.getByRole('button', { name: /Customer is missing/u }));
    expect(onLocate).toHaveBeenCalledWith('binding-1');

    fireEvent.change(screen.getByLabelText('Formatter for binding-1'), {
      target: { value: 'currency' },
    });
    expect(onChange).toHaveBeenCalledWith(
      expect.objectContaining({
        bindings: [expect.objectContaining({ formatter: 'currency' })],
      }),
    );

    fireEvent.click(screen.getByRole('button', { name: 'Add repeat rows' }));
    expect(onChange).toHaveBeenCalledWith(
      expect.objectContaining({
        bindings: expect.arrayContaining([
          expect.objectContaining({
            type: 'repeat-rows',
            range: expect.objectContaining({
              start: { row: 2, column: 1 },
              end: { row: 4, column: 2 },
            }),
          }),
        ]),
      }),
    );

    fireEvent.change(screen.getByLabelText('Paper'), { target: { value: 'Letter' } });
    expect(onChange).toHaveBeenCalledWith(
      expect.objectContaining({
        printProfiles: [
          expect.objectContaining({ page: expect.objectContaining({ paper: { type: 'Letter' } }) }),
        ],
      }),
    );
    fireEvent.click(screen.getByRole('button', { name: 'Use selection as print range' }));
    expect(onChange).toHaveBeenCalledWith(
      expect.objectContaining({
        printProfiles: [
          expect.objectContaining({
            targets: [
              expect.objectContaining({
                type: 'range',
                range: expect.objectContaining({
                  start: { row: 2, column: 1 },
                  end: { row: 4, column: 2 },
                }),
              }),
            ],
          }),
        ],
      }),
    );
    fireEvent.click(screen.getByRole('button', { name: 'Delete binding-1' }));
    expect(onChange).toHaveBeenCalledWith(expect.objectContaining({ bindings: [] }));
  });
});

describe('TemplatePreview', () => {
  it('renders the exact immutable generated pages as accessible SVG', () => {
    const document = {
      print: {
        pages: [
          {
            id: 'page-a',
            index: 0,
            width: 200,
            height: 300,
            targetId: 't',
            rowStart: 0,
            rowEnd: 0,
          },
        ],
        displayList: {
          diagnostics: [],
          pages: [
            {
              index: 0,
              width: 200,
              height: 300,
              commands: [
                {
                  kind: 'text',
                  text: 'Invoice',
                  x: 10,
                  y: 20,
                  maxWidth: 100,
                  fontFamily: 'Arial',
                  fontSize: 10,
                  color: '#000',
                  horizontalAlign: 'left',
                },
              ],
            },
          ],
        },
      },
    } as unknown as GeneratedDocument;
    const before = JSON.stringify(document);
    render(<TemplatePreview document={document} />);

    const page = screen.getByRole('img');
    expect(page.getAttribute('data-page-id')).toBe('page-a');
    expect(page.getAttribute('viewBox')).toBe('0 0 200 300');
    expect(page.textContent).toContain('Invoice');
    expect(JSON.stringify(document)).toBe(before);
  });

  it('is the explicit TegoSheet preview mode and reports deterministic pages', async () => {
    const document = createSpreadsheetDocument({
      id: 'preview-document',
      sheetId: 'sheet-1',
      sheetName: 'Invoice',
    });
    const onDiagnostics = vi.fn();
    render(
      <TegoSheet
        document={document}
        mode="preview"
        template={template}
        sampleData={{ customer: { name: 'Ada' } }}
        onDiagnostics={onDiagnostics}
        renderEnvironment={{
          locale: 'en-US',
          timeZone: 'UTC',
          dateSystem: 'excel-1900',
          clock: new Date('2026-01-01T00:00:00.000Z'),
          fontMetrics: createFontMetrics({
            fonts: { Arial: { averageAdvance: 6, lineHeight: 12 } },
            fallbackFont: 'Arial',
            fallback: { averageAdvance: 6, lineHeight: 12 },
          }),
        }}
      />,
    );

    await waitFor(() => expect(screen.getByRole('img')).toBeTruthy());
    expect(onDiagnostics).toHaveBeenCalled();
  });
});
