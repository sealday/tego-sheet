import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const repositoryRoot = new URL('../../', import.meta.url);
const repositoryPath = fileURLToPath(repositoryRoot);
const normalizeNewlines = (value) => value.replace(/\r\n?/g, '\n');
const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

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
  const step = workflowSteps(job).find(
    (candidate) => stepField(candidate, 'uses')?.value === action,
  );
  assert.ok(step, `job must use ${action} through an anchored uses field`);
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

function assertJobSetup(job, nodeVersion) {
  assert.equal(unquote(yamlScalar(job, 'runs-on', 4)), 'ubuntu-latest');
  assert.match(yamlScalar(job, 'timeout-minutes', 4), /^[1-9]\d*$/);
  actionStep(job, 'actions/checkout@v6');
  const setupNode = actionStep(job, 'actions/setup-node@v6');
  assertActionInput(setupNode, 'node-version', nodeVersion);
  assertActionInput(setupNode, 'cache', 'npm');
  assertRunCommand(job, 'npm ci');
}

function assertArtifactUpload(job, name, paths) {
  const upload = actionStep(job, 'actions/upload-artifact@v5');
  assert.equal(stepField(upload, 'if')?.value, '${{ !cancelled() }}');
  const inputs = actionInputs(upload);
  assert.equal(unquote(yamlScalar(inputs, 'name', 10)), name);
  assertActionInput(upload, 'if-no-files-found', 'ignore');
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
      ['oxlint', 'oxfmt', 'husky', '@commitlint/cli', '@commitlint/config-conventional'].map(
        (name) => [name, packageJson.devDependencies[name]],
      ),
    ),
    {
      oxlint: '1.74.0',
      oxfmt: '0.59.0',
      husky: '9.1.7',
      '@commitlint/cli': '20.5.3',
      '@commitlint/config-conventional': '20.5.3',
    },
  );
  assert.equal(packageJson.packageManager, 'npm@11.13.0');
  assert.equal(packageJson.engines?.node, '>=20.19.0');
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
});

test('Oxlint enables the required plugins and React correctness rules', () => {
  const config = readJson('.oxlintrc.json');

  assert.deepEqual(config.plugins, ['eslint', 'typescript', 'unicorn', 'oxc', 'react']);
  assert.equal(config.categories.correctness, 'error');

  const jsOverride = config.overrides?.find((override) =>
    override.files?.includes('**/*.{js,mjs,cjs}'),
  );
  assert.ok(jsOverride, 'JavaScript override must exist');
  assert.deepEqual(jsOverride.files, ['**/*.{js,mjs,cjs}']);
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
  assert.deepEqual(tsOverride.files, ['**/*.{ts,tsx}']);
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
  assert.deepEqual(reactOverride.files, ['src/**/*.{ts,tsx}', 'tests/**/*.{ts,tsx}']);
  assert.equal(reactOverride.rules['react/rules-of-hooks'], 'error');
  assert.equal(reactOverride.rules['react/exhaustive-deps'], 'error');
  assert.equal(reactOverride.rules['react/react-compiler'], 'error');

  const refreshOverride = config.overrides?.find((override) =>
    override.files?.includes('src/**/*.tsx'),
  );
  assert.ok(refreshOverride, 'React Refresh override must exist');
  assert.deepEqual(refreshOverride.files, ['src/**/*.tsx']);
  assert.deepEqual(refreshOverride.rules['react/only-export-components'], [
    'error',
    { allowConstantExport: true },
  ]);
});

test('Oxfmt pins formatting and excludes only generated or parity-owned paths', () => {
  const config = readJson('.oxfmtrc.json');

  assert.equal(config.singleQuote, true);
  assert.equal(config.printWidth, 100);
  assert.deepEqual(config.ignorePatterns, [
    'docs/superpowers/**',
    'tests/parity/fixtures/**',
    'tests/parity/legacy/**',
    'tests/visual/__snapshots__/**',
    'tests/visual/fonts/**',
  ]);
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

test('CI covers policy, quality, package, browser, visual, and release lanes', () => {
  const workflow = normalizeNewlines(readRepositoryFile('.github/workflows/ci.yml')).replace(
    /^\s*#.*$/gm,
    '',
  );

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
  const concurrency = yamlBlock(workflow, 'concurrency', 0);
  assert.equal(unquote(yamlScalar(concurrency, 'cancel-in-progress', 2)), 'true');

  const jobsSection = yamlBlock(workflow, 'jobs', 0);
  const jobs = [...jobsSection.matchAll(/^ {2}([\w-]+):[ \t]*$/gm)].map((match) => match[1]);
  assert.deepEqual(jobs.toSorted(), [
    'browser',
    'commit-policy',
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

  assert.match(
    yamlScalar(commitPolicy, 'if', 4),
    /^(?:\$\{\{\s*)?github\.event_name != 'workflow_dispatch'(?:\s*}})?$/,
  );
  assertJobSetup(commitPolicy, '24');
  assertActionInput(actionStep(commitPolicy, 'actions/checkout@v6'), 'fetch-depth', '0');
  const commitlintStep = workflowSteps(commitPolicy).find((step) => {
    const run = stepField(step, 'run');
    return run && `${run.value}\n${run.block}`.includes('commitlint');
  });
  assert.ok(commitlintStep, 'commit-policy must run commitlint');
  const commitlintCommand = stepField(commitlintStep, 'run').block;
  assert.match(
    commitlintCommand,
    /--from[\s\S]*github\.event\.pull_request\.base\.sha\s*\|\|\s*github\.event\.before/,
  );
  assert.match(
    commitlintCommand,
    /--to[\s\S]*github\.event\.pull_request\.head\.sha\s*\|\|\s*github\.sha/,
  );
  assert.match(commitlintCommand, /--verbose\b/);

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
  assertArtifactUpload(browser, 'browser-report', [
    'playwright-report/browser',
    'test-results/playwright',
  ]);

  assertJobSetup(visual, '24');
  assertRunPattern(
    visual,
    /^(?:npm exec -- |npx(?: --no --)? )playwright install --with-deps chromium$/,
    'visual must install Chromium with system dependencies',
  );
  assertRunCommand(visual, 'npm run test:visual');
  assertArtifactUpload(visual, 'visual-report', [
    'playwright-report/visual',
    'test-results/playwright-visual',
  ]);

  assert.equal(
    yamlScalar(parityRelease, 'if', 4)
      .replace(/^\$\{\{\s*|\s*}}$/g, '')
      .trim(),
    "github.event_name == 'workflow_dispatch' || (github.event_name == 'push' && github.ref == 'refs/heads/main')",
  );
  assertJobSetup(parityRelease, '24');
  assertRunPattern(
    parityRelease,
    /^(?:npm exec -- |npx(?: --no --)? )playwright install --with-deps$/,
    'parity-release must install all Playwright browsers with dependencies',
  );
  assertRunCommand(parityRelease, 'npm run test:parity-release');
  assertArtifactUpload(parityRelease, 'parity-release-evidence', [
    'test-results/parity',
    'test-results/playwright',
    'test-results/playwright-visual',
    'playwright-report/browser',
    'playwright-report/visual',
  ]);
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
