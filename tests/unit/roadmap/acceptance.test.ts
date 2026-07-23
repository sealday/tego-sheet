import { expect, it } from 'vitest';
import {
  getRoadmapDeliveryState,
  roadmapAcceptance,
} from '../../../docs/roadmap-implementation/acceptance';
import { roadmapItems } from '../../../website/src/data/roadmap';

it('maps every roadmap item to acceptance evidence without duplicate ids', () => {
  expect(new Set(roadmapAcceptance.map((entry) => entry.id)).size).toBe(33);
  expect(roadmapAcceptance.map((entry) => entry.id).sort()).toEqual(
    roadmapItems.map((entry) => entry.id).sort(),
  );
  expect(
    roadmapAcceptance.every(
      (entry) => entry.spec && entry.tasks.length > 0 && entry.tests.length > 0,
    ),
  ).toBe(true);
  expect(roadmapAcceptance.map(({ id, phase }) => ({ id, phase }))).toEqual(
    roadmapItems.map(({ id, phase }) => ({ id, phase })),
  );
  expect(roadmapAcceptance.every((entry) => entry.state === 'planned')).toBe(true);
  expect(getRoadmapDeliveryState('workbook-2')).toBe('planned');
});
