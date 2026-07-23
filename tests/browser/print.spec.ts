import { expect, test } from '@playwright/test';
import { openHarness } from './support';

test('@parity:output.export-download crosses the consumer download boundary', async ({ page }) => {
  await openHarness(page);
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download workbook' }).click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe('workbook.json');
  const stream = await download.createReadStream();
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  expect(JSON.parse(Buffer.concat(chunks).toString('utf8'))).toHaveLength(1);
});
