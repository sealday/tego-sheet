import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { afterEach, describe, expect, it } from 'vitest';
import {
  generateTypeDoc,
  publicApiProjectionPluginPath,
  toTypeDocGenerationOptions,
  type StrictTypeDocApplication,
} from '../../website/plugins/strict-typedoc-generation';

const temporaryDirectories: string[] = [];
const execFileAsync = promisify(execFile);
const root = process.cwd();
const require = createRequire(import.meta.url);

interface ProjectionChild {
  inheritedFrom?: ProjectionTypeReference;
  name: string;
}

interface ProjectionReflection {
  children: ProjectionChild[];
  extendedTypes?: ProjectionTypeReference[];
  name: string;
}

interface ProjectionTypeReference {
  name: string;
  package?: string;
  qualifiedName?: string;
  type: string;
}

type ProjectionListener = (context: {
  project: {
    children: ProjectionReflection[];
    getChildByName: () => ProjectionReflection;
    packageName: string;
  };
}) => void;

const projectionPlugin = require(publicApiProjectionPluginPath) as {
  load: (app: {
    converter: { on: (event: string, listener: ProjectionListener) => void };
    logger: { error: (message: string) => void };
  }) => void;
};

const callbackNames = [
  'onActiveSheetChange',
  'onCellEdit',
  'onChange',
  'onError',
  'onPaste',
  'onSelectionChange',
] as const;

const createProjectionReflection = (): ProjectionReflection => ({
  children: callbackNames.map((name) => ({
    inheritedFrom: {
      name: `TegoSheetCallbacks.${name}`,
      qualifiedName: `TegoSheetCallbacks.${name}`,
      type: 'reference',
    },
    name,
  })),
  extendedTypes: [
    {
      name: 'TegoSheetCallbacks',
      package: 'tego-sheet',
      qualifiedName: 'TegoSheetCallbacks',
      type: 'reference',
    },
  ],
  name: 'TegoSheetProps',
});

const invokeProjection = (
  reflection: ProjectionReflection,
  projectChildren: ProjectionReflection[] = [reflection],
): string[] => {
  const errors: string[] = [];
  let listener: ProjectionListener | undefined;
  projectionPlugin.load({
    converter: {
      on(event, registeredListener) {
        expect(event).toBe('resolveBegin');
        listener = registeredListener;
      },
    },
    logger: { error: (message) => errors.push(message) },
  });

  listener?.({
    project: {
      children: projectChildren,
      getChildByName: () => reflection,
      packageName: 'tego-sheet',
    },
  });
  return errors;
};

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

interface FakeApplicationOptions {
  bootstrapWarnings?: number;
  conversionErrors?: number;
  conversionWarnings?: number;
  outputErrors?: number;
  outputFilePath?: string;
  outputThrows?: boolean;
  outputWarnings?: number;
  strictWarnings?: boolean;
  strictValidation?: boolean;
  validationErrors?: number;
  validationWarnings?: number;
}

const createFakeApplication = (options: FakeApplicationOptions = {}) => {
  const events: string[] = [];
  const calls = { convert: 0, generateOutputs: 0, validate: 0 };
  const logger = {
    errorCount: 0,
    validationWarningCount: 0,
    warningCount: options.bootstrapWarnings ?? 0,
    hasErrors: () => logger.errorCount > 0,
    hasWarnings: () => logger.warningCount > 0,
  };
  const project = { name: 'fake-project' };
  const app: StrictTypeDocApplication = {
    logger,
    options: {
      getValue: (name) => {
        if (name === 'treatValidationWarningsAsErrors') {
          return options.strictValidation ?? true;
        }
        return options.strictWarnings ?? false;
      },
    },
    async convert() {
      calls.convert += 1;
      events.push('convert');
      logger.errorCount += options.conversionErrors ?? 0;
      logger.warningCount += options.conversionWarnings ?? 0;
      return project;
    },
    async generateOutputs(receivedProject) {
      calls.generateOutputs += 1;
      events.push('generateOutputs');
      expect(receivedProject).toBe(project);
      if (options.outputFilePath) {
        await mkdir(join(options.outputFilePath, '..'), { recursive: true });
        await writeFile(options.outputFilePath, 'partial output');
      }
      logger.errorCount += options.outputErrors ?? 0;
      logger.warningCount += options.outputWarnings ?? 0;
      if (options.outputThrows) throw new Error('output writer failed');
    },
    validate(receivedProject) {
      calls.validate += 1;
      events.push('validate');
      expect(receivedProject).toBe(project);
      logger.errorCount += options.validationErrors ?? 0;
      logger.validationWarningCount += options.validationWarnings ?? 0;
      logger.warningCount += options.validationWarnings ?? 0;
    },
  };

  return { app, calls, events };
};

const fakeBootstrap = (app: StrictTypeDocApplication) => async () => app;

const createProjectOutput = async () => {
  const projectRoot = await mkdtemp(join(tmpdir(), 'tego-sheet-project-'));
  temporaryDirectories.push(projectRoot);
  const out = join(projectRoot, 'website/docs/api');
  return { out, projectRoot };
};

describe('strict TypeDoc generation', () => {
  it.each([
    {
      name: 'extra inherited callback',
      mutate(children: ProjectionChild[]) {
        children.push({
          inheritedFrom: {
            name: 'TegoSheetCallbacks.unexpected',
            type: 'reference',
          },
          name: 'unexpected',
        });
      },
    },
    {
      name: 'duplicate callback child',
      mutate(children: ProjectionChild[]) {
        children.push({ ...children[0] });
      },
    },
  ])('fails closed without mutating an $name projection shape', ({ mutate }) => {
    const reflection = createProjectionReflection();
    mutate(reflection.children);
    const errors = invokeProjection(reflection);

    expect(errors).toEqual([
      'public API projection expected exactly six unique TegoSheetCallbacks inherited properties',
    ]);
    expect(reflection.extendedTypes).toEqual([
      {
        name: 'TegoSheetCallbacks',
        package: 'tego-sheet',
        qualifiedName: 'TegoSheetCallbacks',
        type: 'reference',
      },
    ]);
    expect(reflection.children.every((child) => child.inheritedFrom !== undefined)).toBe(true);
  });

  it('rejects an external helper with the same unqualified name', () => {
    const reflection = createProjectionReflection();
    reflection.extendedTypes = [
      {
        name: 'TegoSheetCallbacks',
        package: 'external-callbacks',
        qualifiedName: 'TegoSheetCallbacks',
        type: 'reference',
      },
    ];

    expect(invokeProjection(reflection)).toEqual([
      'public API projection expected TegoSheetProps to extend the project TegoSheetCallbacks helper',
    ]);
    expect(reflection.extendedTypes).toHaveLength(1);
    expect(reflection.children.every((child) => child.inheritedFrom !== undefined)).toBe(true);
  });

  it('rejects a nested TegoSheetProps reflection instead of projecting by recursive name lookup', () => {
    const reflection = createProjectionReflection();
    const namespace: ProjectionReflection = {
      children: [reflection],
      name: 'NestedNamespace',
    };

    expect(invokeProjection(reflection, [namespace])).toEqual([
      'public API projection expected exactly one direct TegoSheetProps project child',
    ]);
    expect(reflection.extendedTypes).toHaveLength(1);
    expect(reflection.children.every((child) => child.inheritedFrom !== undefined)).toBe(true);
  });

  it('loads one TypeDoc runtime across the Docusaurus and native ESM plugin boundary', async () => {
    const pluginPath = join(root, 'website/plugins/strict-typedoc-generation.ts');
    const entryPoint = join(root, 'src/index.ts');
    const tsconfig = join(root, 'tsconfig.json');
    const script = `
      import { createRequire } from 'node:module';
      import { existsSync } from 'node:fs';
      import { mkdir, mkdtemp, readdir, rm } from 'node:fs/promises';
      import { tmpdir } from 'node:os';
      import { join } from 'node:path';
      import { loadFreshModule } from '@docusaurus/utils';

      const originalCwd = process.cwd();
      const projectRoot = await mkdtemp(join(tmpdir(), 'strict-typedoc-runtime-'));
      const require = createRequire(import.meta.url);
      let result;

      try {
        process.chdir(projectRoot);
        await mkdir(join(projectRoot, 'website/docs'), { recursive: true });
        const plugin = await loadFreshModule(${JSON.stringify(pluginPath)});
        await plugin(
          {
            siteDir: join(projectRoot, 'website'),
            siteConfig: { presets: [] },
          },
          {
            entryPoints: [${JSON.stringify(entryPoint)}],
            tsconfig: ${JSON.stringify(tsconfig)},
            out: 'website/docs/api',
            readme: 'none',
            excludePrivate: true,
            excludeProtected: true,
            excludeInternal: true,
            validation: {
              invalidLink: true,
              notDocumented: true,
              notExported: true,
            },
            treatWarningsAsErrors: true,
            treatValidationWarningsAsErrors: true,
          },
        );

        const output = join(projectRoot, 'website/docs/api');
        const paths = globalThis[Symbol.for('typedoc_paths')] ?? [];
        result = {
          bridgeLoaded: Object.keys(require.cache).some((path) =>
            path.endsWith('/website/plugins/strict-typedoc-runtime.cjs'),
          ),
          generated: existsSync(output) && (await readdir(output)).length > 0,
          loads: globalThis[Symbol.for('typedoc_loads')] ?? 0,
          pathCount: paths.length,
          uniquePathCount: new Set(paths).size,
        };
      } finally {
        process.chdir(originalCwd);
        await rm(projectRoot, { force: true, recursive: true });
      }

      process.stdout.write('\\nSTRICT_TYPEDOC_RESULT=' + JSON.stringify(result));
    `;

    const { stdout } = await execFileAsync(
      process.execPath,
      ['--input-type=module', '--eval', script],
      { cwd: root },
    );

    const resultLine = stdout.trim().split('\n').at(-1);
    expect(resultLine).toBeDefined();
    expect(JSON.parse(resultLine!.replace('STRICT_TYPEDOC_RESULT=', ''))).toEqual({
      bridgeLoaded: true,
      generated: true,
      loads: 1,
      pathCount: 1,
      uniquePathCount: 1,
    });
  }, 15_000);

  it('rejects bootstrap warnings before conversion even without TypeDoc strict flags', async () => {
    const { app, calls, events } = createFakeApplication({
      bootstrapWarnings: 1,
      strictValidation: false,
      strictWarnings: false,
    });

    await expect(
      generateTypeDoc(
        { out: 'website/docs/api' },
        { bootstrap: fakeBootstrap(app), clearOutput: async () => undefined },
      ),
    ).rejects.toThrow('TypeDoc bootstrap failed with 0 errors and 1 warnings');

    expect(calls).toEqual({ convert: 0, generateOutputs: 0, validate: 0 });
    expect(events).toEqual([]);
  });

  it('converts once and validates before generating successful output', async () => {
    const { app, calls, events } = createFakeApplication();

    await generateTypeDoc(
      { out: 'website/docs/api' },
      { bootstrap: fakeBootstrap(app), clearOutput: async () => undefined },
    );

    expect(calls).toEqual({ convert: 1, generateOutputs: 1, validate: 1 });
    expect(events).toEqual(['convert', 'validate', 'generateOutputs']);
  });

  it('removes rejected output and never generates when validation fails', async () => {
    const { out, projectRoot } = await createProjectOutput();
    await mkdir(out, { recursive: true });
    await writeFile(join(out, 'rejected.md'), 'stale output');
    const { app, calls, events } = createFakeApplication({ validationWarnings: 1 });

    await expect(
      generateTypeDoc({ out }, { bootstrap: fakeBootstrap(app), projectRoot }),
    ).rejects.toThrow('TypeDoc validation failed with 0 errors and 1 warnings');

    expect(calls).toEqual({ convert: 1, generateOutputs: 0, validate: 1 });
    expect(events).toEqual(['convert', 'validate']);
    expect(existsSync(out)).toBe(false);
  });

  it('removes partial output when output hooks log an error', async () => {
    const { out, projectRoot } = await createProjectOutput();
    const { app } = createFakeApplication({
      outputErrors: 1,
      outputFilePath: join(out, 'partial.md'),
    });

    await expect(
      generateTypeDoc({ out }, { bootstrap: fakeBootstrap(app), projectRoot }),
    ).rejects.toThrow('TypeDoc output failed with 1 errors and 0 warnings');

    expect(existsSync(out)).toBe(false);
  });

  it('removes partial output when strict output warnings are logged', async () => {
    const { out, projectRoot } = await createProjectOutput();
    const { app } = createFakeApplication({
      outputFilePath: join(out, 'partial.md'),
      outputWarnings: 1,
      strictWarnings: true,
    });

    await expect(
      generateTypeDoc({ out }, { bootstrap: fakeBootstrap(app), projectRoot }),
    ).rejects.toThrow('TypeDoc output failed with 0 errors and 1 warnings');

    expect(existsSync(out)).toBe(false);
  });

  it('removes partial output when an output writer throws', async () => {
    const { out, projectRoot } = await createProjectOutput();
    const { app } = createFakeApplication({
      outputFilePath: join(out, 'partial.md'),
      outputThrows: true,
    });

    await expect(
      generateTypeDoc({ out }, { bootstrap: fakeBootstrap(app), projectRoot }),
    ).rejects.toThrow('output writer failed');

    expect(existsSync(out)).toBe(false);
  });

  it('does not generate after conversion errors', async () => {
    const { app, calls, events } = createFakeApplication({ conversionErrors: 1 });

    await expect(
      generateTypeDoc(
        { out: 'website/docs/api' },
        { bootstrap: fakeBootstrap(app), clearOutput: async () => undefined },
      ),
    ).rejects.toThrow('TypeDoc conversion failed with 1 errors and 0 warnings');

    expect(calls).toEqual({ convert: 1, generateOutputs: 0, validate: 0 });
    expect(events).toEqual(['convert']);
  });

  it('rejects unsafe output paths before cleanup or bootstrap', async () => {
    const projectRoot = await mkdtemp(join(tmpdir(), 'tego-sheet-project-'));
    temporaryDirectories.push(projectRoot);
    const outsidePath = await mkdtemp(join(tmpdir(), 'tego-sheet-outside-'));
    temporaryDirectories.push(outsidePath);
    const { app, calls } = createFakeApplication();
    let bootstrapCalls = 0;
    let cleanupCalls = 0;
    const unsafePaths = [
      '.',
      projectRoot,
      'website',
      'website/docs',
      outsidePath,
      '../outside',
      'website/docs/api/../../..',
      'website/docs/api/../api',
      'website/docs/other',
    ];

    for (const out of unsafePaths) {
      await expect(
        generateTypeDoc(
          { out },
          {
            async bootstrap() {
              bootstrapCalls += 1;
              return app;
            },
            async clearOutput() {
              cleanupCalls += 1;
            },
            projectRoot,
          },
        ),
      ).rejects.toThrow('Unsafe TypeDoc output path');
    }

    expect(cleanupCalls).toBe(0);
    expect(bootstrapCalls).toBe(0);
    expect(calls).toEqual({ convert: 0, generateOutputs: 0, validate: 0 });
  });

  it('applies Docusaurus output defaults without dropping conversion settings', () => {
    const options = toTypeDocGenerationOptions(
      {
        siteDir: '/repo/website',
        siteConfig: {
          presets: [
            [
              '@docusaurus/preset-classic',
              { docs: { numberPrefixParser: false, path: './content' } },
            ],
          ],
        },
      },
      {
        entryPoints: ['src/index.ts'],
        excludePrivate: true,
        id: 'default',
        intentionallyNotDocumented: undefined,
        out: 'website/docs/api',
        plugin: ['custom-typedoc-plugin', 'typedoc-plugin-markdown'],
        sidebar: { autoConfiguration: true },
        tsconfig: 'tsconfig.json',
      },
    );

    expect(options).toMatchObject({
      docsPath: '/repo/website/content',
      entryPoints: ['src/index.ts'],
      excludePrivate: true,
      numberPrefixParser: false,
      out: 'website/docs/api',
      sidebar: { autoConfiguration: true },
      tsconfig: 'tsconfig.json',
    });
    expect(options.plugin).toEqual([
      'typedoc-plugin-markdown',
      'typedoc-docusaurus-theme',
      publicApiProjectionPluginPath,
      'custom-typedoc-plugin',
    ]);
    expect(options).not.toHaveProperty('id');
    expect(options).not.toHaveProperty('intentionallyNotDocumented');
    expect(Object.values(options)).not.toContain(undefined);
  });
});
