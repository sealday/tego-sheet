import { expect, test } from '@playwright/test';

test('objects are keyboard accessible, undoable, and locked objects remain selectable', async ({
  page,
}) => {
  await page.goto('/?objects=1');
  await expect(page.locator('[data-tego-sheet]')).toHaveAttribute('data-mode', 'uncontrolled');
  const editable = page.getByRole('option', { name: 'Browser chart', exact: true });
  await expect(editable).toHaveAttribute('aria-description', 'Keyboard-editable browser object');
  await editable.press('ArrowRight');
  await expect.poll(() => page.evaluate(() => window.__tegoHarness.objectRect()?.x)).toBe(21);

  await page.evaluate(() => window.__tegoHarness.undo());
  await expect.poll(() => page.evaluate(() => window.__tegoHarness.objectRect()?.x)).toBe(20);

  const locked = page.getByRole('option', { name: 'Locked browser chart', exact: true });
  await expect(locked).toHaveAttribute('aria-readonly', 'true');
  await locked.press('ArrowRight');
  await expect(locked).toHaveAttribute('aria-selected', 'true');
  await expect.poll(() => page.evaluate(() => window.__tegoHarness.objectRect()?.x)).toBe(20);
});
