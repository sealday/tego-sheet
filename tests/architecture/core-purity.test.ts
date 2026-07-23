import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import ts from 'typescript';
import { expect, it } from 'vitest';
import { WorkbookController, parseWorkbook, serializeWorkbook } from '../../src/core';
import type { WorkbookInput } from '../../src/core';
import { deepFreeze } from '../helpers/deep-freeze';

const root = resolve(import.meta.dirname, '../..');
const forbiddenGlobals = new Set([
  'window',
  'document',
  'navigator',
  'HTMLElement',
  'HTMLCanvasElement',
  'CanvasRenderingContext2D',
  'EventTarget',
  'ResizeObserver',
  'requestAnimationFrame',
  'cancelAnimationFrame',
  'localStorage',
  'sessionStorage',
]);
const configPath = ts.findConfigFile(root, ts.sys.fileExists, 'tsconfig.json');
if (configPath === undefined) throw new Error('tsconfig.json must exist');
const config = ts.readConfigFile(configPath, ts.sys.readFile);
if (config.error !== undefined) {
  throw new Error(ts.flattenDiagnosticMessageText(config.error.messageText, '\n'));
}
const parsedConfig = ts.parseJsonConfigFileContent(config.config, ts.sys, root);
const purityProgram = ts.createProgram({
  rootNames: parsedConfig.fileNames,
  options: parsedConfig.options,
});
const purityChecker = purityProgram.getTypeChecker();

function pureModuleFiles(): readonly string[] {
  return execFileSync('git', ['ls-files', '-z', 'src/core', 'src/document'], {
    cwd: root,
    encoding: 'utf8',
  })
    .split('\0')
    .filter((file) => file.endsWith('.ts'));
}

it('[ARCH-3] keeps every core, controller, and document module independent of React and the browser', () => {
  expect(pureModuleFiles()).toEqual(expect.arrayContaining(['src/document/parse-document.ts']));
  expect(pureModuleFiles().length).toBeGreaterThan(20);
  for (const file of pureModuleFiles()) {
    const source = readFileSync(resolve(root, file), 'utf8');
    const imports = ts.preProcessFile(source).importedFiles.map((entry) => entry.fileName);
    expect(imports, file).not.toEqual(
      expect.arrayContaining([expect.stringMatching(/^(?:react(?:\/|$)|react-dom(?:\/|$))/)]),
    );
    expect(imports, file).not.toEqual(
      expect.arrayContaining([
        expect.stringMatching(/(?:^|\/)ui(?:\/|$)|(?:^|\/)react(?:\/|$)|(?:^|\/)engine(?:\/|$)/),
      ]),
    );

    const sourceFile = purityProgram.getSourceFile(resolve(root, file));
    if (sourceFile === undefined) throw new Error(`${file} must be part of the TypeScript program`);
    const found = new Set<string>();
    const visit = (node: ts.Node): void => {
      if (ts.isIdentifier(node) && forbiddenGlobals.has(node.text)) {
        const symbol = purityChecker.getSymbolAtLocation(node);
        const isWorkspaceDeclaration = symbol?.declarations?.some((declaration) =>
          declaration.getSourceFile().fileName.startsWith(resolve(root, 'src')),
        );
        if (!isWorkspaceDeclaration) found.add(node.text);
      }
      ts.forEachChild(node, visit);
    };
    visit(sourceFile);
    expect([...found], file).toEqual([]);
  }
});

it('[ARCH-2][ARCH-6] round-trips legacy JSON without mutating caller-owned input', () => {
  const input = deepFreeze([
    {
      name: 'Legacy',
      freeze: 'B2',
      customExtension: { nested: ['kept', 7, true] },
      rows: {
        len: 3,
        0: { cells: { 0: { text: '42', editable: false, printable: false } } },
      },
      cols: { len: 2, 0: { width: 88 } },
    },
  ] as WorkbookInput);
  const before = JSON.stringify(input);

  const parsed = parseWorkbook(input);
  const controller = new WorkbookController(input);
  const serialized = serializeWorkbook(controller.getValue());

  expect(JSON.stringify(input)).toBe(before);
  expect(serialized).toEqual(parsed);
  expect(serialized[0]).toMatchObject({
    name: 'Legacy',
    freeze: 'B2',
    customExtension: { nested: ['kept', 7, true] },
    rows: { 0: { cells: { 0: { text: '42', editable: false, printable: false } } } },
  });
  controller.dispose();
});
