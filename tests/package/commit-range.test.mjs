import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

const resolverUrl = new URL('../../scripts/resolve-commit-range.mjs', import.meta.url);

async function loadResolver() {
  assert.equal(existsSync(resolverUrl), true, 'commit-range resolver must exist');
  return import(resolverUrl.href);
}

function fakeGit({ available, fetchSucceeds = false, parent = null, rangeValid = true }) {
  const commits = new Set(available);
  const calls = { fetch: [], ranges: [] };
  return {
    calls,
    operations: {
      fetchCommit(sha) {
        calls.fetch.push(sha);
        if (fetchSucceeds) commits.add(sha);
        return fetchSucceeds;
      },
      isCommitAvailable: (sha) => commits.has(sha),
      isRangeValid(base, head) {
        calls.ranges.push([base, head]);
        return rangeValid;
      },
      parentOf: () => parent,
    },
  };
}

const baseSha = '1'.repeat(40);
const headSha = '2'.repeat(40);
const parentSha = '3'.repeat(40);
const zeroSha = '0'.repeat(40);

test('uses an ordinary available push or pull-request base without fetching', async () => {
  const { resolveCommitRange } = await loadResolver();
  const git = fakeGit({ available: [baseSha, headSha] });

  assert.deepEqual(resolveCommitRange({ baseSha, headSha, git: git.operations }), {
    base: baseSha,
    head: headSha,
    mode: 'range',
  });
  assert.deepEqual(git.calls.fetch, []);
  assert.deepEqual(git.calls.ranges, [[baseSha, headSha]]);
});

test('uses the head parent when the event base is the all-zero SHA', async () => {
  const { resolveCommitRange } = await loadResolver();
  const git = fakeGit({ available: [headSha, parentSha], parent: parentSha });

  assert.deepEqual(resolveCommitRange({ baseSha: zeroSha, headSha, git: git.operations }), {
    base: parentSha,
    head: headSha,
    mode: 'range',
  });
  assert.deepEqual(git.calls.fetch, []);
  assert.deepEqual(git.calls.ranges, [[parentSha, headSha]]);
});

test('fetches an unavailable nonzero base by SHA before using it', async () => {
  const { resolveCommitRange } = await loadResolver();
  const git = fakeGit({ available: [headSha], fetchSucceeds: true, parent: parentSha });

  assert.deepEqual(resolveCommitRange({ baseSha, headSha, git: git.operations }), {
    base: baseSha,
    head: headSha,
    mode: 'range',
  });
  assert.deepEqual(git.calls.fetch, [baseSha]);
  assert.deepEqual(git.calls.ranges, [[baseSha, headSha]]);
});

test('falls back to a validated parent range when fetching the base fails', async () => {
  const { resolveCommitRange } = await loadResolver();
  const git = fakeGit({ available: [headSha, parentSha], parent: parentSha });

  assert.deepEqual(resolveCommitRange({ baseSha, headSha, git: git.operations }), {
    base: parentSha,
    head: headSha,
    mode: 'range',
  });
  assert.deepEqual(git.calls.fetch, [baseSha]);
  assert.deepEqual(git.calls.ranges, [[parentSha, headSha]]);
});

test('uses last mode when the head commit has no parent', async () => {
  const { resolveCommitRange } = await loadResolver();
  const git = fakeGit({ available: [headSha], parent: null });

  assert.deepEqual(resolveCommitRange({ baseSha: zeroSha, headSha, git: git.operations }), {
    base: '',
    head: headSha,
    mode: 'last',
  });
  assert.deepEqual(git.calls.ranges, []);
});

test('rejects malformed and zero head SHAs before invoking Git', async () => {
  const { resolveCommitRange } = await loadResolver();
  const git = fakeGit({ available: [] });

  assert.throws(
    () => resolveCommitRange({ baseSha: 'HEAD~1', headSha, git: git.operations }),
    /base SHA must be exactly 40 hexadecimal characters/,
  );
  assert.throws(
    () => resolveCommitRange({ baseSha, headSha: zeroSha, git: git.operations }),
    /head SHA must not be the all-zero SHA/,
  );
  assert.deepEqual(git.calls.fetch, []);
});

test('CLI writes shell-safe range outputs to GITHUB_OUTPUT', async () => {
  const { runCli } = await loadResolver();
  const temporaryDirectory = mkdtempSync(join(tmpdir(), 'tego-sheet-commit-range-'));
  const outputPath = join(temporaryDirectory, 'github-output');
  const git = fakeGit({ available: [baseSha, headSha] });

  try {
    assert.deepEqual(
      runCli({
        argv: [baseSha, headSha],
        env: { GITHUB_OUTPUT: outputPath },
        git: git.operations,
      }),
      { base: baseSha, head: headSha, mode: 'range' },
    );
    assert.equal(
      readFileSync(outputPath, 'utf8'),
      `mode=range\nbase=${baseSha}\nhead=${headSha}\n`,
    );
  } finally {
    rmSync(temporaryDirectory, { force: true, recursive: true });
  }
});
