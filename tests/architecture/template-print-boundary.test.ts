import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

const root = process.cwd();

describe('template print architecture boundary', () => {
  it('removes the legacy viewport and Canvas print implementation', () => {
    expect(existsSync(resolve(root, 'src/engine/canvas/print-renderer.ts'))).toBe(false);
    expect(existsSync(resolve(root, 'src/ui/print-workbook.ts'))).toBe(false);
    expect(existsSync(resolve(root, 'src/ui/dialogs/print-dialog.tsx'))).toBe(false);
  });

  it('keeps browser output compiler-centered and independent from editor state', () => {
    const source = readFileSync(resolve(root, 'src/output/browser-print-adapter.ts'), 'utf8');
    expect(source).toContain('displayList');
    expect(source).not.toMatch(/CanvasRenderingContext2D|window\\.print|mountPrintPages/u);
    expect(source).not.toMatch(/selection|scroll|zoom|devicePixelRatio/u);
  });
});
