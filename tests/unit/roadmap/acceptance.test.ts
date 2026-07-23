import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createSlugger } from '@docusaurus/utils';
import { expect, expectTypeOf, it } from 'vitest';
import {
  getRoadmapDeliveryState,
  roadmapAcceptance,
  type RoadmapAcceptanceEntry,
} from '../../../docs/roadmap-implementation/acceptance';
import {
  allRoadmapItems,
  selectRoadmapItemsByStatus,
  type RoadmapItemId,
} from '../../../website/src/data/roadmap';

it('maps every roadmap item to acceptance evidence without duplicate ids', () => {
  expect(new Set(roadmapAcceptance.map((entry) => entry.id)).size).toBe(allRoadmapItems.length);
  expect(roadmapAcceptance.map((entry) => entry.id).sort()).toEqual(
    allRoadmapItems.map((entry) => entry.id).sort(),
  );
  expect(roadmapAcceptance.map(({ id, phase }) => ({ id, phase }))).toEqual(
    allRoadmapItems.map(({ id, phase }) => ({ id, phase })),
  );
  expect(
    roadmapAcceptance.filter((entry) => entry.state === 'shipped').map((entry) => entry.id),
  ).toEqual([
    'workbook-2',
    'transactions',
    'formula-format-core',
    'render-semantics',
    'extension-kernel',
  ]);
  expect(getRoadmapDeliveryState('workbook-2')).toBe('shipped');
  expect(getRoadmapDeliveryState('transactions')).toBe('shipped');
  expect(getRoadmapDeliveryState('formula-format-core')).toBe('shipped');
  expect(getRoadmapDeliveryState('render-semantics')).toBe('shipped');
  expect(getRoadmapDeliveryState('extension-kernel')).toBe('shipped');
});

it('keeps canonical ids when an item moves out of the planned projection', () => {
  const deliveryFixture = allRoadmapItems.map((item) =>
    item.id === 'transactions' ? { ...item, status: 'shipped' as const } : item,
  );

  expect(deliveryFixture.map((item) => item.id)).toContain('transactions');
  expect(
    selectRoadmapItemsByStatus(deliveryFixture, 'planned').map((item) => item.id),
  ).not.toContain('transactions');
  expect(selectRoadmapItemsByStatus(deliveryFixture, 'shipped').map((item) => item.id)).toEqual([
    'workbook-2',
    'transactions',
    'formula-format-core',
    'render-semantics',
    'extension-kernel',
  ]);
});

it('keeps acceptance ids typed and derives phase from the canonical record', () => {
  expectTypeOf<RoadmapAcceptanceEntry['id']>().toEqualTypeOf<RoadmapItemId>();

  const canonicalById = new Map(allRoadmapItems.map((item) => [item.id, item]));
  for (const acceptance of roadmapAcceptance) {
    expect(acceptance.phase).toBe(canonicalById.get(acceptance.id)?.phase);
  }
});

it('resolves every spec file and Docusaurus heading anchor', () => {
  for (const acceptance of roadmapAcceptance) {
    const [file, fragment] = acceptance.spec.split('#');
    expect(file).toBeTruthy();
    expect(fragment).toBeTruthy();

    const absoluteFile = join(process.cwd(), file!);
    expect(existsSync(absoluteFile), `${acceptance.id} spec file must exist`).toBe(true);

    const slugger = createSlugger();
    const headingFragments = [
      ...readFileSync(absoluteFile, 'utf8').matchAll(/^#{1,6}\s+(.+?)\s*$/gm),
    ]
      .map((match) => match[1]!.replace(/\s+#+$/, ''))
      .map((heading) => slugger.slug(heading));

    expect(headingFragments, `${acceptance.id} spec anchor must resolve`).toContain(fragment);
  }
});

it('links nonblank tasks to approved plan sections and executable verification commands', () => {
  const plan = readFileSync(
    join(process.cwd(), 'docs/superpowers/plans/2026-07-23-complete-product-roadmap.md'),
    'utf8',
  );
  const vitestConfig = readFileSync(join(process.cwd(), 'vitest.config.ts'), 'utf8');
  const configuredProjects = new Set(
    [...vitestConfig.matchAll(/name:\s*'([^']+)'/g)].map((match) => match[1]),
  );

  for (const acceptance of roadmapAcceptance) {
    expect(acceptance.tasks.length).toBeGreaterThan(0);
    expect(acceptance.tests.length).toBeGreaterThan(0);

    for (const task of acceptance.tasks) {
      expect(task.trim(), `${acceptance.id} task must be nonblank`).not.toBe('');
      const taskNumber = task.match(/^Task (\d+):/)?.[1];
      expect(taskNumber, `${acceptance.id} task must use a stable Task ID`).toBeTruthy();
      expect(plan).toContain(`### Task ${taskNumber}:`);
    }

    for (const command of acceptance.tests) {
      expect(command.trim(), `${acceptance.id} test command must be nonblank`).not.toBe('');
      const taskNumber = acceptance.tasks[0]!.match(/^Task (\d+):/)![1];
      const taskStart = plan.indexOf(`### Task ${taskNumber}:`);
      const nextTask = plan.indexOf('\n### Task ', taskStart + 1);
      const taskSection = plan.slice(taskStart, nextTask < 0 ? undefined : nextTask);
      const commandMarker = `Run: \`${command}\``;
      const commandIndex = taskSection.indexOf(commandMarker);
      expect(
        commandIndex,
        `${acceptance.id} command must be an approved task verification gate`,
      ).toBeGreaterThanOrEqual(0);
      const commandOutcome = taskSection
        .slice(commandIndex + commandMarker.length)
        .split('\nRun: ', 1)[0];
      expect(commandOutcome, `${acceptance.id} command must be a passing final gate`).toMatch(
        /Expected: PASS/,
      );

      for (const project of command.matchAll(/--project\s+(\S+)/g)) {
        expect(
          configuredProjects,
          `${acceptance.id} references Vitest project ${project[1]}`,
        ).toContain(project[1]);
      }

      for (const path of command.matchAll(/\btests\/[^\s&]+/g)) {
        expect(
          existsSync(join(process.cwd(), path[0])) || taskSection.includes(path[0]),
          `${acceptance.id} references an undeclared test path ${path[0]}`,
        ).toBe(true);
      }
    }
  }
});
