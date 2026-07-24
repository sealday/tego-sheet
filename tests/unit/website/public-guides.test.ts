import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const root = process.cwd();
const read = (path: string): string => readFileSync(join(root, path), 'utf8');

describe('public migration and integration guides', () => {
  it('publishes the two supported migration paths', () => {
    const legacy = read('docs/migration-from-x-data-spreadsheet.md');
    const workbook2 = read('docs/migration-to-workbook-2.md');

    expect(legacy).toContain('Removed legacy APIs');
    expect(legacy).toContain('migrateLegacyWorkbook');
    expect(legacy).toContain('MigrationDiagnostic');
    expect(workbook2).toContain('serializeSpreadsheetDocument');
    expect(workbook2).toContain('LEGACY_FIELD_DEGRADED');
    expect(workbook2).toContain('controlled');
  });

  it.each([
    ['templates.md', 'compileSpreadsheetTemplate'],
    ['output.md', 'XlsxAdapter'],
    ['extensions.md', 'createAdapterRegistry'],
    ['integrations.md', 'createPermissionStore'],
  ])('publishes and preserves the hand-written API guide %s', (name, symbol) => {
    const document = read(`website/docs/api/${name}`);
    expect(document).toContain(symbol);
    expect(document).toContain('tego-sheet');
  });

  it('links the new public surfaces from the README', () => {
    const readme = read('readme.md');

    expect(readme).toContain('tego-sheet/interchange');
    expect(readme).toContain('tego-sheet/sdk');
    expect(readme).toContain('tego-sheet/integrations');
    expect(readme).toContain('docs/migration-to-workbook-2.md');
  });
});
