import { cleanup, fireEvent, render, within } from '@testing-library/react';
import { afterEach, expect, it, vi } from 'vitest';
import { FilterMenu } from '../../src/ui/menus/filter-menu';

const t = (_path: string, fallback: string): string => fallback;

afterEach(cleanup);

it('applies the selected filter values as a flat string array', () => {
  const onApply = vi.fn();
  const rendered = render(
    <FilterMenu
      column={1}
      values={['Keyboard', 'Mouse']}
      onApply={onApply}
      onClose={vi.fn()}
      t={t}
    />,
  );
  const dialog = rendered.getByRole('dialog', { name: 'Filter' });

  fireEvent.click(within(dialog).getByLabelText('Mouse'));
  fireEvent.click(within(dialog).getByRole('button', { name: 'Apply filter' }));

  expect(onApply).toHaveBeenCalledWith({ column: 1, operator: 'in', value: ['Keyboard'] });
});
