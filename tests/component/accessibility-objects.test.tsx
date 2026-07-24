import { fireEvent, render } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { AccessibilityObjects } from '../../src/react/accessibility/accessibility-objects';
import type { ScreenObjectProjection } from '../../src/objects';

const projection = (id: string, locked = false): ScreenObjectProjection =>
  ({
    object: {
      id,
      kind: 'shape',
      anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 40, height: 30 } },
      zIndex: 1,
      locked,
      templateRepeat: 'shared',
      shape: 'rectangle',
      style: { fill: '#ffeecc' },
      accessibility: {
        name: locked ? 'Locked chart' : 'Revenue chart',
        description: locked ? 'Quarterly chart, locked' : 'Quarterly revenue',
      },
    },
    bounds: { x: 10, y: 20, width: 40, height: 30 },
    commands: [],
    diagnostics: [],
  }) as never;

describe('object accessibility layer', () => {
  it('exposes visible objects with role, name, description, selection, and lock state', () => {
    const view = render(
      <AccessibilityObjects
        objects={[projection('editable'), projection('locked', true)]}
        selectedObjectId="editable"
        onSelect={vi.fn()}
        onChange={vi.fn()}
      />,
    );

    const editable = view.getByRole('option', { name: 'Revenue chart' });
    expect(editable.getAttribute('aria-description')).toBe('Quarterly revenue');
    expect(editable.getAttribute('aria-selected')).toBe('true');
    expect(editable.getAttribute('aria-readonly')).toBe('false');
    const locked = view.getByRole('option', { name: 'Locked chart' });
    expect(locked.getAttribute('aria-selected')).toBe('false');
    expect(locked.getAttribute('aria-readonly')).toBe('true');
    view.unmount();
  });

  it('selects locked objects but only dispatches keyboard transforms for editable objects', () => {
    const onSelect = vi.fn();
    const onChange = vi.fn();
    const view = render(
      <AccessibilityObjects
        objects={[projection('editable'), projection('locked', true)]}
        selectedObjectId={null}
        onSelect={onSelect}
        onChange={onChange}
      />,
    );

    const locked = view.getByRole('option', { name: 'Locked chart' });
    fireEvent.click(locked);
    fireEvent.keyDown(locked, { key: 'ArrowRight' });
    expect(onSelect).toHaveBeenCalledWith('locked');
    expect(onChange).not.toHaveBeenCalled();

    const editable = view.getByRole('option', { name: 'Revenue chart' });
    fireEvent.keyDown(editable, { key: 'ArrowRight', shiftKey: true });
    expect(onSelect).toHaveBeenCalledWith('editable');
    expect(onChange).toHaveBeenCalledWith(
      expect.objectContaining({
        id: 'editable',
        anchor: { type: 'absolute', rect: { x: 10, y: 20, width: 41, height: 30 } },
      }),
    );
    view.unmount();
  });
});
