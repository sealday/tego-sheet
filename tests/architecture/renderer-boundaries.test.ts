import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import ts from 'typescript';
import { expect, it } from 'vitest';

const root = resolve(import.meta.dirname, '../..');

function trackedSources(...directories: string[]): readonly string[] {
  return execFileSync('git', ['ls-files', '-co', '--exclude-standard', '-z', ...directories], {
    cwd: root,
    encoding: 'utf8',
  })
    .split('\0')
    .filter((file) => /\.(?:ts|tsx)$/u.test(file));
}

it('[ARCH-F4] keeps shared presentation and display-list modules browser and React independent', () => {
  for (const file of trackedSources('src/presentation', 'src/print')) {
    const source = readFileSync(resolve(root, file), 'utf8');
    const imports = ts.preProcessFile(source).importedFiles.map(({ fileName }) => fileName);

    expect(imports, file).not.toEqual(
      expect.arrayContaining([
        expect.stringMatching(/^(?:react(?:\/|$)|react-dom(?:\/|$))/u),
        expect.stringMatching(/(?:^|\/)(?:ui|react|engine)(?:\/|$)/u),
      ]),
    );
    expect(source, file).not.toMatch(
      /\b(?:window|navigator|HTMLCanvasElement|CanvasRenderingContext2D|devicePixelRatio)\b|globalThis\.document/u,
    );
  }
});

it('[ARCH-F4] makes renderers consume presentation without importing mutation controllers', () => {
  for (const file of trackedSources('src/engine/canvas', 'src/react/accessibility', 'src/print')) {
    const source = readFileSync(resolve(root, file), 'utf8');
    const imports = ts.preProcessFile(source).importedFiles.map(({ fileName }) => fileName);

    expect(imports, file).not.toEqual(
      expect.arrayContaining([
        expect.stringMatching(/(?:controller|commands|transaction|document-patch)/u),
      ]),
    );
  }
});

it('[ARCH-F4] prevents Canvas and print renderers from owning formula formatting semantics', () => {
  for (const file of trackedSources('src/engine/canvas', 'src/print')) {
    const source = readFileSync(resolve(root, file), 'utf8');

    expect(source, file).not.toMatch(
      /(?:evaluateCell|createFormulaEvaluationBudget|formatValue|createNumberFormatter)/u,
    );
  }
});
