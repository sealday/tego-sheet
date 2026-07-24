import { describe, expect, it, vi } from 'vitest';
import {
  createCellEditorSession,
  createCellTypeRegistry,
  resolveCustomCell,
} from '../../../src/sdk/cells';

const statusPlugin = {
  manifest: {
    id: 'status-cell',
    apiVersion: '1.0' as const,
    execution: 'trusted-main' as const,
    environments: ['browser'] as const,
    capabilities: [] as const,
  },
  type: 'acme.status',
  schemaVersion: 2,
  validate: (value: unknown): value is { id: string; label: string } =>
    typeof value === 'object' &&
    value !== null &&
    'id' in value &&
    typeof value.id === 'string' &&
    'label' in value &&
    typeof value.label === 'string',
  deserialize: (value: unknown, storedVersion: number) =>
    storedVersion === 1 && typeof value === 'string'
      ? { ok: true as const, value: { id: value, label: value } }
      : { ok: true as const, value: value as { id: string; label: string } },
  serialize: (value: { id: string; label: string }) => value,
  format: (value: { id: string; label: string }) => value.label,
  accessibility: (value: { id: string; label: string }) => `Status: ${value.label}`,
};

describe('cell extension SDK', () => {
  it('registers versioned plugins and resolves one text semantic across channels', () => {
    const registry = createCellTypeRegistry({ apiVersion: '1.0', environment: 'browser' });
    const unregister = registry.register(statusPlugin);

    expect(
      resolveCustomCell(
        registry,
        {
          type: 'custom',
          cellType: 'acme.status',
          schemaVersion: 1,
          value: 'open',
        },
        { locale: 'en-US' },
      ),
    ).toEqual({
      status: 'resolved',
      value: { id: 'open', label: 'open' },
      formattedText: 'open',
      accessibilityLabel: 'Status: open',
      printText: 'open',
      diagnostics: [
        {
          code: 'CELL_PRINT_FALLBACK',
          message: 'Cell plugin acme.status has no print renderer; formatted text was used',
        },
      ],
    });
    unregister();
    expect(registry.get('acme.status')).toBeUndefined();
  });

  it('rejects duplicate types and oversized serialized values', () => {
    const registry = createCellTypeRegistry({ apiVersion: '1.0', environment: 'browser' });
    registry.register(statusPlugin);
    expect(() => registry.register(statusPlugin)).toThrow(/already registered/u);
    expect(() =>
      resolveCustomCell(
        registry,
        {
          type: 'custom',
          cellType: 'acme.status',
          schemaVersion: 2,
          value: { id: 'large', label: 'x'.repeat(70_000) },
        },
        { locale: 'en-US' },
      ),
    ).toThrow(/65536/u);
  });

  it('degrades unknown plugins without losing serialized data', () => {
    const registry = createCellTypeRegistry({ apiVersion: '1.0', environment: 'browser' });
    const input = {
      type: 'custom' as const,
      cellType: 'missing.person',
      schemaVersion: 3,
      value: { userId: 'u-1', name: '<Ada>' },
    };

    const result = resolveCustomCell(registry, input, { locale: 'en-US' });

    expect(result.status).toBe('fallback');
    if (result.status !== 'fallback') throw new Error('expected fallback');
    expect(result.formattedText).toBe('{"name":"<Ada>","userId":"u-1"}');
    expect(result.serializedInput).toEqual(input);
    expect(result.diagnostics[0]?.code).toBe('CELL_PLUGIN_UNAVAILABLE');
  });

  it('creates an exactly-once commit or cancel editor session', () => {
    const commit = vi.fn();
    const cancel = vi.fn();
    const controller = new AbortController();
    const session = createCellEditorSession({
      initialValue: { id: 'open', label: 'Open' },
      signal: controller.signal,
      validate: statusPlugin.validate,
      commit,
      cancel,
    });

    expect(session.commit({ id: 'done', label: 'Done' })).toBe(true);
    expect(session.commit({ id: 'again', label: 'Again' })).toBe(false);
    expect(session.cancel()).toBe(false);
    expect(commit).toHaveBeenCalledTimes(1);
    expect(cancel).not.toHaveBeenCalled();
  });

  it('cancels an editor session on abort without committing', () => {
    const commit = vi.fn();
    const cancel = vi.fn();
    const controller = new AbortController();
    const session = createCellEditorSession({
      initialValue: { id: 'open', label: 'Open' },
      signal: controller.signal,
      validate: statusPlugin.validate,
      commit,
      cancel,
    });
    controller.abort();

    expect(session.commit({ id: 'done', label: 'Done' })).toBe(false);
    expect(cancel).toHaveBeenCalledTimes(1);
    session.dispose();
  });
});
