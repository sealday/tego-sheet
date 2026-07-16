import { execFileSync } from 'node:child_process';
import { appendFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const commitShaPattern = /^[0-9a-f]{40}$/i;
const zeroShaPattern = /^0{40}$/;

function normalizeSha(value, label, { allowZero }) {
  if (typeof value !== 'string' || !commitShaPattern.test(value)) {
    throw new Error(`${label} SHA must be exactly 40 hexadecimal characters`);
  }
  const normalized = value.toLowerCase();
  if (!allowZero && zeroShaPattern.test(normalized)) {
    throw new Error(`${label} SHA must not be the all-zero SHA`);
  }
  return normalized;
}

function gitSucceeds(args, cwd) {
  try {
    execFileSync('git', args, { cwd, stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
}

export function createGitOperations({ cwd = process.cwd() } = {}) {
  return {
    fetchCommit(sha) {
      return gitSucceeds(['fetch', '--no-tags', '--depth=1', 'origin', sha], cwd);
    },
    isCommitAvailable(sha) {
      return gitSucceeds(['cat-file', '-e', `${sha}^{commit}`], cwd);
    },
    isRangeValid(base, head) {
      try {
        const count = execFileSync('git', ['rev-list', '--count', `${base}..${head}`], {
          cwd,
          encoding: 'utf8',
          stdio: ['ignore', 'pipe', 'ignore'],
        });
        return Number.parseInt(count.trim(), 10) > 0;
      } catch {
        return false;
      }
    },
    parentOf(sha) {
      const revision = execFileSync('git', ['rev-list', '--parents', '--max-count=1', sha], {
        cwd,
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'ignore'],
      })
        .trim()
        .split(/\s+/);
      if (revision[0]?.toLowerCase() !== sha.toLowerCase()) {
        throw new Error(`Git did not resolve the expected head commit ${sha}`);
      }
      return revision[1] ?? null;
    },
  };
}

export function resolveCommitRange({ baseSha, headSha, git = createGitOperations() }) {
  const base = normalizeSha(baseSha, 'base', { allowZero: true });
  const head = normalizeSha(headSha, 'head', { allowZero: false });

  if (!git.isCommitAvailable(head)) {
    throw new Error(`head commit ${head} is not available`);
  }

  if (!zeroShaPattern.test(base)) {
    let baseAvailable = git.isCommitAvailable(base);
    if (!baseAvailable) {
      try {
        git.fetchCommit(base);
      } catch {
        // A server may reject fetch-by-SHA; the parent fallback below remains safe.
      }
      baseAvailable = git.isCommitAvailable(base);
    }
    if (baseAvailable && git.isRangeValid(base, head)) {
      return { base, head, mode: 'range' };
    }
  }

  const parentValue = git.parentOf(head);
  if (parentValue === null) return { base: '', head, mode: 'last' };
  const parent = normalizeSha(parentValue, 'parent', { allowZero: false });
  if (!git.isCommitAvailable(parent) || !git.isRangeValid(parent, head)) {
    throw new Error(`fallback range ${parent}..${head} is not valid`);
  }
  return { base: parent, head, mode: 'range' };
}

export function runCli({
  argv = process.argv.slice(2),
  env = process.env,
  git = createGitOperations(),
} = {}) {
  if (argv.length !== 2) {
    throw new Error('usage: node scripts/resolve-commit-range.mjs <base-sha> <head-sha>');
  }
  if (typeof env.GITHUB_OUTPUT !== 'string' || env.GITHUB_OUTPUT.length === 0) {
    throw new Error('GITHUB_OUTPUT must name the workflow output file');
  }

  const result = resolveCommitRange({ baseSha: argv[0], headSha: argv[1], git });
  appendFileSync(
    env.GITHUB_OUTPUT,
    `mode=${result.mode}\nbase=${result.base}\nhead=${result.head}\n`,
    'utf8',
  );
  return result;
}

const invokedPath = process.argv[1] === undefined ? '' : resolve(process.argv[1]);
if (invokedPath === fileURLToPath(import.meta.url)) {
  try {
    runCli();
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
