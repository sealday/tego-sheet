import { describe, expect, it } from 'vitest';
import { createInvoiceOutputFixture } from '../../../website/src/components/playground/output-studio-fixtures';
import { renderOutputRevision } from '../../../website/src/components/playground/output-studio-pipeline';

describe('Output Studio rendering', () => {
  it('renders the prepared invoice into two deterministic pages', async () => {
    const fixture = createInvoiceOutputFixture();
    const result = await renderOutputRevision({
      revision: 1,
      document: fixture.document,
      template: fixture.template,
      activePrintProfileId: fixture.template.printProfiles[0]!.id,
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });

    expect(result.revision).toBe(1);
    expect(result.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(result.document?.print.pages).toHaveLength(2);
    expect(result.document?.print.displayList.pages).toHaveLength(2);
    expect(result.document?.metadata.generatedAt).toBe('2026-07-25T00:00:00.000Z');
  });

  it('returns blocking diagnostics instead of an artifact for invalid bindings', async () => {
    const fixture = createInvoiceOutputFixture();
    const customerBinding = fixture.template.bindings[0]!;
    if (customerBinding.type !== 'value') throw new TypeError('Expected a value binding');
    const result = await renderOutputRevision({
      revision: 2,
      document: fixture.document,
      template: {
        ...fixture.template,
        bindings: [{ ...customerBinding, expression: 'missing.customer' }],
      },
      activePrintProfileId: fixture.template.printProfiles[0]!.id,
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });

    expect(result.document).toBeUndefined();
    expect(result.diagnostics.some(({ severity }) => severity === 'error')).toBe(true);
  });

  it('returns blocking diagnostics instead of throwing when no print profile exists', async () => {
    const fixture = createInvoiceOutputFixture();
    const result = await renderOutputRevision({
      revision: 3,
      document: fixture.document,
      template: { ...fixture.template, printProfiles: [] },
      activePrintProfileId: fixture.template.printProfiles[0]!.id,
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });

    expect(result).toMatchObject({
      revision: 3,
      diagnostics: [expect.objectContaining({ code: 'INVALID_PRINT_TARGET', severity: 'error' })],
    });
    expect(result.document).toBeUndefined();
  });

  it('returns an atomic aborted result for a cancelled revision', async () => {
    const fixture = createInvoiceOutputFixture();
    const controller = new AbortController();
    controller.abort();

    const result = await renderOutputRevision({
      revision: 4,
      document: fixture.document,
      template: fixture.template,
      activePrintProfileId: fixture.template.printProfiles[0]!.id,
      data: fixture.data,
      environment: fixture.environment,
      signal: controller.signal,
    });

    expect(result).toMatchObject({
      revision: 4,
      diagnostics: [expect.objectContaining({ code: 'RENDER_ABORTED', severity: 'error' })],
    });
    expect(result.document).toBeUndefined();
  });

  it('renders the explicitly selected print profile', async () => {
    const fixture = createInvoiceOutputFixture();
    const selectedProfile = {
      ...fixture.template.printProfiles[0]!,
      id: 'invoice-letter-landscape',
      name: 'Invoice · Letter landscape',
      page: {
        ...fixture.template.printProfiles[0]!.page,
        paper: { type: 'Letter' as const },
        orientation: 'landscape' as const,
      },
    };
    const result = await renderOutputRevision({
      revision: 5,
      document: fixture.document,
      template: {
        ...fixture.template,
        printProfiles: [...fixture.template.printProfiles, selectedProfile],
      },
      activePrintProfileId: selectedProfile.id,
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });

    expect(result.document?.print.pages[0]).toMatchObject({
      width: expect.any(Number),
      height: expect.any(Number),
    });
    expect(result.document!.print.pages[0]!.width).toBeGreaterThan(
      result.document!.print.pages[0]!.height,
    );
  });

  it('blocks an explicit print profile that no longer exists', async () => {
    const fixture = createInvoiceOutputFixture();
    const result = await renderOutputRevision({
      revision: 6,
      document: fixture.document,
      template: fixture.template,
      activePrintProfileId: 'removed-profile',
      data: fixture.data,
      environment: fixture.environment,
      signal: new AbortController().signal,
    });

    expect(result.document).toBeUndefined();
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'INVALID_PRINT_TARGET',
        severity: 'error',
        message: 'Selected print profile no longer exists',
      }),
    );
  });
});
