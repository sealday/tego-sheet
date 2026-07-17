import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const repositoryRoot = new URL('../../', import.meta.url);
const repositoryPath = fileURLToPath(repositoryRoot);
const normalizeNewlines = (value) => value.replace(/\r\n?/g, '\n');
const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const sorted = (values) => [...values].sort();
const actionPins = Object.freeze({
  'actions/checkout': {
    sha: 'df4cb1c069e1874edd31b4311f1884172cec0e10',
    version: 'v6',
  },
  'actions/setup-node': {
    sha: '249970729cb0ef3589644e2896645e5dc5ba9c38',
    version: 'v6',
  },
  'actions/upload-artifact': {
    sha: '330a01c490aca151604b8cf639adc76d48f6c5d4',
    version: 'v5',
  },
  'actions/configure-pages': {
    sha: '983d7736d9b0ae728b81ab479565c72886d7745b',
    version: 'v5',
  },
  'actions/upload-pages-artifact': {
    sha: '7b1f4a764d45c48632c6b24a0339c27f5614fb0b',
    version: 'v4',
  },
  'actions/deploy-pages': {
    sha: 'd6db90164ac5ed86f2b6aed7e0febac5b3c0c03e',
    version: 'v4',
  },
});
const documentationDevDependencyPins = Object.freeze({
  '@docusaurus/core': '3.10.2',
  '@docusaurus/preset-classic': '3.10.2',
  '@docusaurus/tsconfig': '3.10.2',
  '@docusaurus/types': '3.10.2',
  '@mdx-js/react': '3.1.1',
  'docusaurus-plugin-typedoc': '1.4.2',
  'prism-react-renderer': '2.4.1',
  typedoc: '0.28.20',
  'typedoc-docusaurus-theme': '1.4.2',
  'typedoc-plugin-markdown': '4.12.0',
});

function readRepositoryFile(path) {
  const url = new URL(path, repositoryRoot);
  assert.equal(existsSync(url), true, `${path} must exist`);
  return readFileSync(url, 'utf8');
}
const readJson = (path) => JSON.parse(readRepositoryFile(path));

function meaningfulCommands(path) {
  return readRepositoryFile(path)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith('#'));
}

function yamlBlock(source, key, indent) {
  const lines = source.split('\n');
  const header = new RegExp(`^ {${indent}}${escapeRegExp(key)}:[ \\t]*(.*)$`);
  const start = lines.findIndex((line) => header.test(line));
  assert.notEqual(start, -1, `${key} block must exist at indentation ${indent}`);

  let end = start + 1;
  while (end < lines.length) {
    const line = lines[end];
    if (line.trim() && line.length - line.trimStart().length <= indent) break;
    end += 1;
  }
  return lines.slice(start, end).join('\n');
}

function yamlScalar(source, key, indent) {
  const match = source.match(
    new RegExp(`^ {${indent}}${escapeRegExp(key)}:[ \\t]*(\\S(?:.*\\S)?)?[ \\t]*$`, 'm'),
  );
  assert.ok(match, `${key} must exist at indentation ${indent}`);
  return match[1] ?? '';
}

function unquote(value) {
  return value.replace(/^(['"])(.*)\1$/, '$2');
}

function jobBlock(jobs, jobId) {
  return yamlBlock(jobs, jobId, 2);
}

function workflowSteps(job) {
  const steps = yamlBlock(job, 'steps', 4).split('\n');
  const starts = steps.flatMap((line, index) => (/^ {6}- /.test(line) ? [index] : []));
  assert.ok(starts.length > 0, 'job must define steps');
  return starts.map((start, index) =>
    steps.slice(start, starts[index + 1] ?? steps.length).join('\n'),
  );
}

function stepField(step, key) {
  const first = step.match(new RegExp(`^ {6}- ${escapeRegExp(key)}:[ \\t]*(.*)$`, 'm'));
  if (first) return { value: first[1].trim(), block: step };

  const lines = step.split('\n');
  const field = new RegExp(`^ {8}${escapeRegExp(key)}:[ \\t]*(.*)$`);
  const start = lines.findIndex((line) => field.test(line));
  if (start === -1) return undefined;
  const value = lines[start].match(field)[1].trim();
  let end = start + 1;
  while (end < lines.length && (!lines[end].trim() || /^ {10}/.test(lines[end]))) end += 1;
  return { value, block: lines.slice(start, end).join('\n') };
}

function actionStep(job, action) {
  const pin = actionPins[action];
  assert.ok(pin, `${action} must have a repository-owned immutable pin`);
  assert.match(pin.sha, /^[0-9a-f]{40}$/);
  const expected = `${action}@${pin.sha} # ${pin.version}`;
  const step = workflowSteps(job).find(
    (candidate) => stepField(candidate, 'uses')?.value === expected,
  );
  assert.ok(step, `job must use immutable ${expected} through an anchored uses field`);
  return step;
}

function actionInputs(step) {
  return yamlBlock(step, 'with', 8);
}

function assertActionInput(step, key, expected) {
  const actual = unquote(yamlScalar(actionInputs(step), key, 10));
  assert.equal(actual, expected, `${key} must be ${expected}`);
}

function assertRunCommand(job, command) {
  const found = workflowSteps(job).some((step) => stepField(step, 'run')?.value === command);
  assert.equal(found, true, `job must run ${command}`);
}

function assertRunPattern(job, pattern, description) {
  const found = workflowSteps(job).some((step) =>
    pattern.test(stepField(step, 'run')?.value ?? ''),
  );
  assert.equal(found, true, description);
}

function stepRunScript(step) {
  const run = stepField(step, 'run');
  assert.ok(run, 'step must define run');
  if (!['|', '|-', '>', '>-'].includes(run.value)) return unquote(run.value);
  return run.block
    .split('\n')
    .slice(1)
    .map((line) => line.replace(/^ {10}/, ''))
    .join('\n')
    .trim();
}

function assertJobSetup(job, nodeVersion, runner = 'ubuntu-latest') {
  assert.equal(unquote(yamlScalar(job, 'runs-on', 4)), runner);
  assert.match(yamlScalar(job, 'timeout-minutes', 4), /^[1-9]\d*$/);
  actionStep(job, 'actions/checkout');
  const setupNode = actionStep(job, 'actions/setup-node');
  assertActionInput(setupNode, 'node-version', nodeVersion);
  assertActionInput(setupNode, 'cache', 'npm');
  assertRunCommand(job, 'npm ci');
}

function assertArtifactUpload(job, name, paths, retentionDays, condition = '${{ !cancelled() }}') {
  const uploadArtifactPin = actionPins['actions/upload-artifact'];
  const upload = workflowSteps(job).find((step) => {
    if (
      stepField(step, 'uses')?.value !==
      `actions/upload-artifact@${uploadArtifactPin.sha} # ${uploadArtifactPin.version}`
    )
      return false;
    const withBlock = yamlBlock(step, 'with', 8);
    return unquote(yamlScalar(withBlock, 'name', 10)) === name;
  });
  assert.ok(upload, `job must upload the ${name} artifact`);
  assert.equal(stepField(upload, 'if')?.value, condition);
  const inputs = actionInputs(upload);
  assert.equal(unquote(yamlScalar(inputs, 'name', 10)), name);
  assertActionInput(upload, 'if-no-files-found', 'ignore');
  assertActionInput(upload, 'retention-days', String(retentionDays));
  const pathInput = yamlBlock(inputs, 'path', 10);
  const inlinePath = unquote(yamlScalar(inputs, 'path', 10));
  const pathValues = ['|', '|-', '>', '>-'].includes(inlinePath)
    ? pathInput
        .split('\n')
        .slice(1)
        .map((line) => line.trim())
        .filter(Boolean)
    : [inlinePath];
  for (const path of paths) {
    assert.ok(pathValues.includes(path), `artifact path input must include ${path}`);
  }
  return upload;
}

function yamlList(source, key, indent) {
  const block = yamlBlock(source, key, indent);
  return block
    .split('\n')
    .slice(1)
    .map((line) => line.match(new RegExp(`^ {${indent + 2}}-\\s*(\\S.*?)\\s*$`))?.[1])
    .filter(Boolean)
    .map(unquote);
}

function configBlock(config, name, closingToken) {
  const openingToken = closingToken.startsWith('}') ? '{' : '[';
  const match = config.match(
    new RegExp(
      `^  ${escapeRegExp(name)}:\\s*${escapeRegExp(openingToken)}([\\s\\S]*?)^  ${escapeRegExp(closingToken)},?\\s*$`,
      'm',
    ),
  );
  assert.ok(match, `${name} configuration must exist`);
  return match[1];
}

test('ESLint configuration is removed', () => {
  assert.equal(existsSync(new URL('eslint.config.js', repositoryRoot)), false);
});

test('package metadata pins the repository toolchain and supported runtimes', () => {
  const packageJson = readJson('package.json');

  assert.deepEqual(
    Object.fromEntries(
      [
        'oxlint',
        'oxfmt',
        'husky',
        '@commitlint/cli',
        '@commitlint/config-conventional',
        ...Object.keys(documentationDevDependencyPins),
      ].map((name) => [name, packageJson.devDependencies[name]]),
    ),
    {
      oxlint: '1.74.0',
      oxfmt: '0.59.0',
      husky: '9.1.7',
      '@commitlint/cli': '20.5.3',
      '@commitlint/config-conventional': '20.5.3',
      ...documentationDevDependencyPins,
    },
  );
  assert.equal(packageJson.packageManager, 'npm@11.13.0');
  assert.equal(packageJson.engines?.node, '>=20.19.0');
});

test('package lock root metadata pins the documentation toolchain', () => {
  const packageLock = readJson('package-lock.json');
  const rootDevDependencies = packageLock.packages?.['']?.devDependencies;

  assert.ok(rootDevDependencies, 'package-lock.json must define root devDependencies');
  assert.deepEqual(
    Object.fromEntries(
      Object.keys(documentationDevDependencyPins).map((name) => [name, rootDevDependencies[name]]),
    ),
    documentationDevDependencyPins,
  );
});

test('package lock resolved entries pin the documentation toolchain', () => {
  const packageLock = readJson('package-lock.json');

  assert.deepEqual(
    Object.fromEntries(
      Object.keys(documentationDevDependencyPins).map((name) => [
        name,
        packageLock.packages?.[`node_modules/${name}`]?.version,
      ]),
    ),
    documentationDevDependencyPins,
  );
});

test('documentation site package metadata is only a module-scope sentinel', () => {
  const websitePackageJson = readJson('website/package.json');

  assert.deepEqual(websitePackageJson, { private: true });
  assert.equal(existsSync(new URL('website/package-lock.json', repositoryRoot)), false);
  for (const key of [
    'scripts',
    'dependencies',
    'devDependencies',
    'workspaces',
    'packageManager',
  ]) {
    assert.equal(
      Object.hasOwn(websitePackageJson, key),
      false,
      `website/package.json forbids ${key}`,
    );
  }
});

test('package lock contains only the root project and installed packages', () => {
  const packageLock = readJson('package-lock.json');
  const extraPackageRoots = Object.keys(packageLock.packages ?? {}).filter(
    (path) => path !== '' && !path.startsWith('node_modules/'),
  );

  assert.deepEqual(extraPackageRoots, [], 'package-lock.json must not define nested package roots');
});

test('legacy direct ESLint dependencies are removed', () => {
  const devDependencies = readJson('package.json').devDependencies;

  for (const dependency of [
    '@eslint/js',
    'eslint',
    'eslint-plugin-react-hooks',
    'eslint-plugin-react-refresh',
    'typescript-eslint',
  ]) {
    assert.equal(
      Object.hasOwn(devDependencies, dependency),
      false,
      `${dependency} must be removed from devDependencies`,
    );
  }
});

test('package scripts expose the Oxlint, Oxfmt, and Husky commands', () => {
  const scripts = readJson('package.json').scripts;

  assert.deepEqual(
    Object.fromEntries(
      ['lint', 'lint:fix', 'format', 'format:check', 'prepare'].map((name) => [
        name,
        scripts[name],
      ]),
    ),
    {
      lint: 'oxlint --deny-warnings .',
      'lint:fix': 'oxlint --fix .',
      format: 'oxfmt --write .',
      'format:check': 'oxfmt --check .',
      prepare: 'husky',
    },
  );

  assert.equal(scripts['docs:start'], 'npm run build && docusaurus start website');
  assert.equal(scripts['docs:build'], 'npm run build && docusaurus build website');
  assert.equal(scripts['docs:serve'], 'docusaurus serve website');
  assert.equal(
    scripts['typecheck:docs'],
    'npm run build && tsc --noEmit --project website/tsconfig.json',
  );
  assert.equal(scripts['test:docs'], 'playwright test --config playwright.docs.config.ts');
  assert.equal(
    scripts['test:docs-visual'],
    'playwright test --config playwright.docs-visual.config.ts',
  );
});

test('Oxlint enables the required plugins and React correctness rules', () => {
  const config = readJson('.oxlintrc.json');

  assert.deepEqual(config.plugins, ['eslint', 'typescript', 'unicorn', 'oxc', 'react']);
  assert.equal(config.categories.correctness, 'error');

  const jsOverride = config.overrides?.find((override) =>
    override.files?.includes('**/*.{js,mjs,cjs}'),
  );
  assert.ok(jsOverride, 'JavaScript override must exist');
  assert.deepEqual(sorted(jsOverride.files), sorted(['**/*.{js,mjs,cjs}']));
  assert.equal(jsOverride.env.node, true);
  for (const rule of [
    'no-case-declarations',
    'no-empty',
    'no-fallthrough',
    'no-prototype-builtins',
    'no-redeclare',
    'no-regex-spaces',
    'no-undef',
    'no-unexpected-multiline',
    'no-useless-assignment',
    'preserve-caught-error',
  ])
    assert.equal(jsOverride.rules[rule], 'error', `JavaScript coverage must include ${rule}`);

  const tsOverride = config.overrides?.find((override) =>
    override.files?.includes('**/*.{ts,tsx}'),
  );
  assert.ok(tsOverride, 'TypeScript override must exist');
  assert.deepEqual(sorted(tsOverride.files), sorted(['**/*.{ts,tsx}']));
  for (const rule of [
    'no-var',
    'prefer-const',
    'prefer-rest-params',
    'prefer-spread',
    'no-array-constructor',
    'typescript/ban-ts-comment',
    'typescript/no-empty-object-type',
    'typescript/no-explicit-any',
    'typescript/no-namespace',
    'typescript/no-require-imports',
    'typescript/no-unnecessary-type-constraint',
    'typescript/no-unsafe-function-type',
  ])
    assert.equal(tsOverride.rules[rule], 'error', `TypeScript coverage must include ${rule}`);

  const reactOverride = config.overrides?.find((override) =>
    override.files?.includes('src/**/*.{ts,tsx}'),
  );
  assert.ok(reactOverride, 'React Hooks override must exist');
  assert.deepEqual(
    sorted(reactOverride.files),
    sorted(['src/**/*.{ts,tsx}', 'website/src/**/*.{ts,tsx}', 'tests/**/*.{ts,tsx}']),
  );
  assert.equal(reactOverride.rules['react/rules-of-hooks'], 'error');
  assert.equal(reactOverride.rules['react/exhaustive-deps'], 'error');
  assert.equal(reactOverride.rules['react/react-compiler'], 'error');

  const refreshOverride = config.overrides?.find((override) =>
    override.files?.includes('src/**/*.tsx'),
  );
  assert.ok(refreshOverride, 'React Refresh override must exist');
  assert.deepEqual(sorted(refreshOverride.files), sorted(['src/**/*.tsx', 'website/src/**/*.tsx']));
  assert.deepEqual(refreshOverride.rules['react/only-export-components'], [
    'error',
    { allowConstantExport: true },
  ]);
});

test('Oxfmt pins formatting and excludes only generated or parity-owned paths', () => {
  const config = readJson('.oxfmtrc.json');

  assert.equal(config.singleQuote, true);
  assert.equal(config.printWidth, 100);
  assert.deepEqual(
    sorted(config.ignorePatterns),
    sorted([
      'docs/superpowers/**',
      'tests/parity/fixtures/**',
      'tests/parity/legacy/**',
      'tests/visual/__snapshots__/**',
      'tests/visual/fonts/**',
      'website/docs/api/**',
      'website/build/**',
      'website/.docusaurus/**',
    ]),
  );
});

test('generated documentation outputs are ignored', () => {
  const ignorePatterns = readRepositoryFile('.gitignore').split(/\r?\n/);

  for (const path of ['/website/build/', '/website/.docusaurus/', '/website/docs/api/']) {
    assert.ok(ignorePatterns.includes(path), `.gitignore must include ${path}`);
  }
});

test('commit messages use the unweakened conventional configuration', async () => {
  assert.equal(existsSync(new URL('commitlint.config.js', repositoryRoot)), true);
  const { default: config } = await import(new URL('commitlint.config.js', repositoryRoot));

  assert.deepEqual(config, { extends: ['@commitlint/config-conventional'] });
});

test('Husky hooks enforce formatting, linting, and commit policy', () => {
  assert.deepEqual(meaningfulCommands('.husky/pre-commit'), [
    'npm run format:check && npm run lint',
  ]);
  assert.deepEqual(meaningfulCommands('.husky/commit-msg'), ['npx --no -- commitlint --edit "$1"']);
});

test('Playwright uses npm-hosted Vite servers and separated HTML reports', () => {
  const browserConfig = readRepositoryFile('playwright.config.ts').replace(/^\s*\/\/.*$/gm, '');
  const visualConfig = readRepositoryFile('playwright.visual.config.ts').replace(
    /^\s*\/\/.*$/gm,
    '',
  );

  for (const [name, config, reportDirectory] of [
    ['browser', browserConfig, 'playwright-report/browser'],
    ['visual', visualConfig, 'playwright-report/visual'],
  ]) {
    const webServer = configBlock(config, 'webServer', '}');
    const reporters = configBlock(config, 'reporter', ']');
    assert.match(
      webServer,
      /^\s+command:\s*['"]npm exec -- vite(?:\s+[^'"]*)?['"],?\s*$/m,
      `${name} webServer.command must use npm exec -- vite`,
    );
    assert.doesNotMatch(webServer, /\bpnpm\b/, `${name} webServer must not invoke pnpm`);
    assert.match(
      reporters,
      new RegExp(
        `\\[\\s*['"]html['"]\\s*,\\s*\\{[\\s\\S]*?outputFolder:\\s*['"]${reportDirectory}['"]`,
      ),
      `${name} config must write its HTML report to ${reportDirectory}`,
    );
  }
});

test('documentation Playwright lanes build and serve the project subpath with isolated artifacts', () => {
  const docsConfig = readRepositoryFile('playwright.docs.config.ts').replace(/^\s*\/\/.*$/gm, '');
  const docsVisualConfig = readRepositoryFile('playwright.docs-visual.config.ts').replace(
    /^\s*\/\/.*$/gm,
    '',
  );
  const lanes = [
    {
      name: 'docs browser',
      config: docsConfig,
      baseUrl: 'http://127.0.0.1:4175/tego-sheet/',
      command: 'npm run docs:build && npm run docs:serve -- --host 127.0.0.1 --port 4175',
      outputDirectory: 'test-results/playwright-docs',
      reportDirectory: 'playwright-report/docs',
    },
    {
      name: 'docs visual',
      config: docsVisualConfig,
      baseUrl: 'http://127.0.0.1:4176/tego-sheet/',
      command: 'npm run docs:build && npm run docs:serve -- --host 127.0.0.1 --port 4176',
      outputDirectory: 'test-results/playwright-docs-visual',
      reportDirectory: 'playwright-report/docs-visual',
    },
  ];

  for (const lane of lanes) {
    const webServer = configBlock(lane.config, 'webServer', '}');
    assert.match(
      lane.config,
      new RegExp(`baseURL:\\s*['"]${escapeRegExp(lane.baseUrl)}['"]`),
      `${lane.name} must use the deployed project subpath as its base URL`,
    );
    assert.match(
      lane.config,
      new RegExp(`outputDir:\\s*['"]${escapeRegExp(lane.outputDirectory)}['"]`),
      `${lane.name} must use ${lane.outputDirectory}`,
    );
    assert.match(
      webServer,
      new RegExp(`command:\\s*['"]${escapeRegExp(lane.command)}['"]`),
      `${lane.name} must freshly build before serving`,
    );
    assert.match(
      webServer,
      new RegExp(`url:\\s*['"]${escapeRegExp(lane.baseUrl)}['"]`),
      `${lane.name} must wait for the project subpath`,
    );
    assert.match(webServer, /timeout:\s*180_000/);
    assert.match(webServer, /reuseExistingServer:\s*!process\.env\.CI/);
    assert.match(
      lane.config,
      new RegExp(
        `\\[\\s*['"]html['"]\\s*,\\s*\\{[\\s\\S]*?outputFolder:\\s*['"]${escapeRegExp(lane.reportDirectory)}['"]`,
      ),
      `${lane.name} must write its HTML report to ${lane.reportDirectory}`,
    );
  }

  assert.notEqual(lanes[0].outputDirectory, lanes[1].outputDirectory);
  assert.notEqual(lanes[0].reportDirectory, lanes[1].reportDirectory);
});

test('CI covers policy, quality, package, browser, visual, and release lanes', () => {
  const workflow = normalizeNewlines(readRepositoryFile('.github/workflows/ci.yml'));

  const actionReferences = workflow.split('\n').filter((line) => /^\s+(?:-\s+)?uses:/.test(line));
  assert.ok(actionReferences.length > 0, 'workflow must use pinned repository Actions');
  for (const reference of actionReferences) {
    const match = reference.match(/^\s+(?:-\s+)?uses:\s*([^@\s]+)@([0-9a-f]{40})\s+#\s+(v\d+)\s*$/);
    assert.ok(
      match,
      `Action reference must use a 40-character SHA and version comment: ${reference}`,
    );
    const [, action, sha, version] = match;
    assert.deepEqual({ sha, version }, actionPins[action], `${action} must use its approved pin`);
  }

  const triggers = yamlBlock(workflow, 'on', 0);
  const mainBranch = String.raw`(?:\[\s*['"]?main['"]?\s*\]|\n {6}-\s*['"]?main['"]?\s*)`;
  for (const event of ['pull_request', 'push']) {
    const eventBlock = yamlBlock(triggers, event, 2);
    assert.match(
      eventBlock,
      new RegExp(`^ {4}branches:\\s*${mainBranch}`, 'm'),
      `${event} must target main`,
    );
  }
  assert.match(triggers, /^ {2}workflow_dispatch:\s*(?:\{\s*\})?\s*$/m);

  const permissions = yamlBlock(workflow, 'permissions', 0);
  assert.equal(unquote(yamlScalar(permissions, 'contents', 2)), 'read');
  assert.equal(permissions.trim(), 'permissions:\n  contents: read');
  const concurrency = yamlBlock(workflow, 'concurrency', 0);
  assert.equal(unquote(yamlScalar(concurrency, 'cancel-in-progress', 2)), 'true');

  const jobsSection = yamlBlock(workflow, 'jobs', 0);
  const jobs = [...jobsSection.matchAll(/^ {2}([\w-]+):[ \t]*$/gm)].map((match) => match[1]);
  assert.deepEqual(jobs.toSorted(), [
    'browser',
    'commit-policy',
    'deploy-pages',
    'docs-site',
    'package-contract',
    'parity-release',
    'quality',
    'visual',
    'vitest',
  ]);

  const commitPolicy = jobBlock(jobsSection, 'commit-policy');
  const quality = jobBlock(jobsSection, 'quality');
  const vitest = jobBlock(jobsSection, 'vitest');
  const packageContract = jobBlock(jobsSection, 'package-contract');
  const browser = jobBlock(jobsSection, 'browser');
  const visual = jobBlock(jobsSection, 'visual');
  const parityRelease = jobBlock(jobsSection, 'parity-release');
  const docsSite = jobBlock(jobsSection, 'docs-site');
  const deployPages = jobBlock(jobsSection, 'deploy-pages');

  assert.match(
    yamlScalar(commitPolicy, 'if', 4),
    /^(?:\$\{\{\s*)?github\.event_name != 'workflow_dispatch'(?:\s*}})?$/,
  );
  assertJobSetup(commitPolicy, '24');
  assertActionInput(actionStep(commitPolicy, 'actions/checkout'), 'fetch-depth', '0');
  const resolverStep = workflowSteps(commitPolicy).find(
    (step) => stepField(step, 'name')?.value === 'Resolve introduced commit range',
  );
  assert.ok(resolverStep, 'commit-policy must resolve the introduced commit range');
  assert.equal(stepField(resolverStep, 'id')?.value, 'commit-range');
  const resolverEnvironment = yamlBlock(resolverStep, 'env', 8);
  assert.equal(
    yamlScalar(resolverEnvironment, 'BASE_SHA', 10),
    '${{ github.event.pull_request.base.sha || github.event.before }}',
  );
  assert.equal(
    yamlScalar(resolverEnvironment, 'HEAD_SHA', 10),
    '${{ github.event.pull_request.head.sha || github.sha }}',
  );
  assert.equal(
    stepRunScript(resolverStep),
    'node scripts/resolve-commit-range.mjs "$BASE_SHA" "$HEAD_SHA"',
  );
  const commitlintStep = workflowSteps(commitPolicy).find((step) => {
    const run = stepField(step, 'run');
    return run && `${run.value}\n${run.block}`.includes('commitlint');
  });
  assert.ok(commitlintStep, 'commit-policy must run commitlint');
  assert.equal(stepField(commitlintStep, 'name')?.value, 'Validate introduced commits');
  const commitlintEnvironment = yamlBlock(commitlintStep, 'env', 8);
  assert.equal(
    yamlScalar(commitlintEnvironment, 'RANGE_MODE', 10),
    '${{ steps.commit-range.outputs.mode }}',
  );
  assert.equal(
    yamlScalar(commitlintEnvironment, 'RANGE_BASE', 10),
    '${{ steps.commit-range.outputs.base }}',
  );
  assert.equal(
    yamlScalar(commitlintEnvironment, 'RANGE_HEAD', 10),
    '${{ steps.commit-range.outputs.head }}',
  );
  const commitlintScript = stepRunScript(commitlintStep);
  assert.match(
    commitlintScript,
    /npm exec -- commitlint --from "\$RANGE_BASE" --to "\$RANGE_HEAD" --verbose/,
  );
  assert.match(commitlintScript, /npm exec -- commitlint --last --verbose/);
  assert.match(commitlintScript, /case "\$RANGE_MODE" in/);
  assert.doesNotMatch(commitlintScript, /\beval\b|\$\{\{[^}]+}}/);

  assertJobSetup(quality, '24');
  for (const command of [
    'npm run format:check',
    'npm run lint',
    'npm run typecheck',
    'npm run build',
    'npm run build:demo',
  ]) {
    assertRunCommand(quality, command);
  }

  assertJobSetup(vitest, '24');
  assertRunCommand(vitest, 'npm test');

  assertJobSetup(packageContract, '${{ matrix.node }}');
  const strategy = yamlBlock(packageContract, 'strategy', 4);
  const matrix = yamlBlock(strategy, 'matrix', 6);
  assert.match(
    matrix,
    /^ {8}node:\s*\[\s*['"]?20['"]?\s*,\s*['"]?22['"]?\s*,\s*['"]?24['"]?\s*\]\s*$/m,
  );
  assertRunCommand(packageContract, 'npm run test:ssr');
  assertRunCommand(packageContract, 'npm run test:package');

  assertJobSetup(browser, '24');
  assertRunPattern(
    browser,
    /^(?:npm exec -- |npx(?: --no --)? )playwright install --with-deps$/,
    'browser must install all Playwright browsers with dependencies',
  );
  assertRunCommand(browser, 'npm run test:browser');
  assertArtifactUpload(
    browser,
    'browser-report',
    ['playwright-report/browser', 'test-results/playwright'],
    7,
  );

  assertJobSetup(visual, '24', 'macos-14');
  assertRunPattern(
    visual,
    /^(?:npm exec -- |npx(?: --no --)? )playwright install chromium$/,
    'visual must install Chromium on the snapshot baseline platform',
  );
  assertRunCommand(visual, 'npm run test:visual');
  assertRunCommand(visual, 'npm run test:docs-visual');
  assertArtifactUpload(
    visual,
    'visual-report',
    ['playwright-report/visual', 'test-results/playwright-visual'],
    7,
  );
  assertArtifactUpload(
    visual,
    'docs-visual-report',
    ['playwright-report/docs-visual', 'test-results/playwright-docs-visual'],
    7,
    '${{ failure() }}',
  );

  assert.equal(
    yamlScalar(parityRelease, 'if', 4)
      .replace(/^\$\{\{\s*|\s*}}$/g, '')
      .trim(),
    "github.event_name == 'workflow_dispatch' || (github.event_name == 'push' && github.ref == 'refs/heads/main')",
  );
  assertJobSetup(parityRelease, '24', 'macos-14');
  assertRunPattern(
    parityRelease,
    /^(?:npm exec -- |npx(?: --no --)? )playwright install$/,
    'parity-release must install all Playwright browsers on the snapshot baseline platform',
  );
  assertRunCommand(parityRelease, 'npm run test:parity-release');
  assertArtifactUpload(
    parityRelease,
    'parity-release-evidence',
    [
      'test-results/parity',
      'test-results/playwright',
      'test-results/playwright-visual',
      'playwright-report/browser',
      'playwright-report/visual',
    ],
    30,
  );

  assert.doesNotMatch(docsSite, /^ {4}if:/m, 'docs-site must run for pull requests');
  const docsPermissions = yamlBlock(docsSite, 'permissions', 4);
  assert.equal(docsPermissions.trim(), 'permissions:\n      contents: read\n      pages: read');
  assertJobSetup(docsSite, '20');
  assertRunPattern(
    docsSite,
    /^(?:npm exec -- |npx(?: --no --)? )playwright install --with-deps chromium$/,
    'docs-site must install Chromium with system dependencies',
  );
  assertRunCommand(docsSite, 'npm run typecheck:docs');
  assertRunCommand(docsSite, 'npm run test:docs');

  const mainPushCondition = "${{ github.event_name == 'push' && github.ref == 'refs/heads/main' }}";
  const configurePages = actionStep(docsSite, 'actions/configure-pages');
  const uploadPages = actionStep(docsSite, 'actions/upload-pages-artifact');
  assert.equal(stepField(configurePages, 'if')?.value, mainPushCondition);
  assert.equal(stepField(uploadPages, 'if')?.value, mainPushCondition);
  assertActionInput(uploadPages, 'path', 'website/build');
  const docsReport = assertArtifactUpload(
    docsSite,
    'docs-report',
    ['playwright-report/docs', 'test-results/playwright-docs'],
    7,
    '${{ failure() }}',
  );

  const docsSteps = workflowSteps(docsSite);
  const docsBrowserTestIndex = docsSteps.findIndex(
    (step) => stepField(step, 'run')?.value === 'npm run test:docs',
  );
  const uploadPagesIndex = docsSteps.indexOf(uploadPages);
  assert.ok(
    docsBrowserTestIndex < uploadPagesIndex,
    'Pages must upload the website/build directory tested by the docs browser suite',
  );
  assert.deepEqual(
    docsSteps.slice(docsBrowserTestIndex + 1, uploadPagesIndex),
    [docsReport],
    'only the non-mutating failure report may run before uploading tested website/build',
  );
  assert.equal(
    stepField(docsReport, 'run'),
    undefined,
    'the failure report must not rebuild or delete the tested website/build',
  );

  assert.equal(yamlScalar(deployPages, 'if', 4), mainPushCondition);
  assert.deepEqual(yamlList(deployPages, 'needs', 4), [
    'commit-policy',
    'quality',
    'vitest',
    'package-contract',
    'browser',
    'visual',
    'parity-release',
    'docs-site',
  ]);
  const deployPermissions = yamlBlock(deployPages, 'permissions', 4);
  assert.equal(deployPermissions.trim(), 'permissions:\n      pages: write\n      id-token: write');
  assert.deepEqual(
    Object.fromEntries(
      ['pages', 'id-token'].map((permission) => [
        permission,
        unquote(yamlScalar(deployPermissions, permission, 6)),
      ]),
    ),
    { pages: 'write', 'id-token': 'write' },
  );
  for (const job of jobs.filter((job) => !['deploy-pages', 'docs-site'].includes(job))) {
    assert.doesNotMatch(
      jobBlock(jobsSection, job),
      /^ {4}permissions:/m,
      `${job} must inherit read-only workflow permissions`,
    );
  }
  const environment = yamlBlock(deployPages, 'environment', 4);
  assert.equal(unquote(yamlScalar(environment, 'name', 6)), 'github-pages');
  assert.equal(yamlScalar(environment, 'url', 6), '${{ steps.deployment.outputs.page_url }}');
  const deployStep = actionStep(deployPages, 'actions/deploy-pages');
  assert.equal(stepField(deployStep, 'name')?.value, 'Deploy GitHub Pages');
  assert.equal(stepField(deployStep, 'id')?.value, 'deployment');
});

test('README presents the tracked demo and ends with upstream attribution', () => {
  const imagePath = 'docs/assets/tego-sheet-demo.png';
  const readme = readRepositoryFile('readme.md');
  const ownershipHeading = '## Ownership and upstream attribution';
  const headings = [...readme.matchAll(/^## .+$/gm)].map((match) => match[0]);

  assert.equal(existsSync(new URL(imagePath, repositoryRoot)), true);
  assert.doesNotThrow(() =>
    execFileSync('git', ['ls-files', '--error-unmatch', imagePath], {
      cwd: repositoryPath,
      stdio: 'pipe',
    }),
  );
  assert.ok(
    readme.includes('![Tego Sheet interactive workbench](docs/assets/tego-sheet-demo.png)'),
  );
  assert.equal(headings.filter((heading) => heading === ownershipHeading).length, 1);
  assert.equal(headings.at(-1), ownershipHeading);
  assert.ok(
    readme
      .trimEnd()
      .endsWith(
        'Third-party assets that carry their own notices remain subject to their respective licenses.',
      ),
  );
});
