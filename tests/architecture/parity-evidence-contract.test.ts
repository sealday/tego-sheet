import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { parityManifest } from '../parity/manifest';
import type { ParityLane } from '../parity/manifest-types';

const laneDirectories: Readonly<Record<ParityLane, readonly string[]>> = {
  unit: ['tests/unit', 'tests/property', 'tests/contract'],
  component: ['tests/component'],
  browser: ['tests/browser'],
  visual: ['tests/visual'],
};

function trackedSource(directories: readonly string[]): string {
  const files = execFileSync('git', ['ls-files', '-z', '--', ...directories], {
    encoding: 'utf8',
  })
    .split('\0')
    .filter((file) => /\.[cm]?[jt]sx?$/u.test(file));
  return files.map((file) => readFileSync(file, 'utf8')).join('\n');
}

it('maps every declared parity assertion to an exact test source token', () => {
  for (const lane of Object.keys(laneDirectories) as ParityLane[]) {
    const source = trackedSource(laneDirectories[lane]);
    const assertionIds = parityManifest.flatMap((row) =>
      'assertions' in row[lane] ? row[lane].assertions : [],
    );

    for (const assertionId of assertionIds) {
      expect(source, `${lane} is missing @parity:${assertionId}`).toContain(
        `@parity:${assertionId}`,
      );
    }
  }
});
