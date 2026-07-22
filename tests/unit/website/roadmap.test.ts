import { readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { groupRoadmapItems, roadmapItems, roadmapPhases } from '../../../website/src/data/roadmap';

const root = process.cwd();
const roadmapDocument = (name: string): string =>
  readFileSync(join(root, 'website/docs/roadmap', name), 'utf8');

const miniRfcs = {
  'foundation.md': ['F1', 'F2', 'F3', 'F4', 'F5'],
  'template-printing.md': ['TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'TP6'],
  'formulas-data.md': ['FMT-01', 'VAL-01', 'FRM-01', 'VIEW-01', 'DATA-01', 'IO-01'],
  'analysis-visualization.md': [
    'TBL-01',
    'CHT-01',
    'SPK-01',
    'OBJ-01',
    'PVT-01',
    'SLC-01',
    'GSK-01',
    'SLV-01',
  ],
  'extensibility.md': ['E1', 'E2', 'E3'],
  'host-integrations.md': ['H1', 'H2', 'H3', 'H4', 'H5', 'H6'],
} as const;

describe('product roadmap data', () => {
  it('contains the approved capabilities in dependency order', () => {
    expect(roadmapPhases.map((phase) => phase.id)).toEqual([0, 1, 2, 3, 4]);
    expect(roadmapItems).toHaveLength(33);
    expect(new Set(roadmapItems.map((item) => item.id))).toHaveProperty('size', 33);
    expect(roadmapItems.every((item) => item.status === 'planned')).toBe(true);
    expect(roadmapItems[0]?.title).toBe('Workbook 2.0 typed document model');
    expect(
      roadmapItems.some(
        (item) => item.title === 'Safe scalar bindings, repeat rows and conditional ranges',
      ),
    ).toBe(true);
  });

  it('groups every item under one declared phase and a Docusaurus design route', () => {
    const groups = groupRoadmapItems();

    expect([...groups.keys()]).toEqual(roadmapPhases.map((phase) => phase.id));
    expect([...groups.values()].flat()).toHaveLength(roadmapItems.length);
    expect(roadmapItems.every((item) => item.designTo.startsWith('/docs/roadmap/'))).toBe(true);
  });

  it('keeps the public index in exact parity with the typed display data', () => {
    const tableRows = roadmapDocument('index.md')
      .split('\n')
      .filter((line) => /^\| [0-4] /.test(line))
      .map((line) => {
        const [phase, title, status, design] = line
          .split('|')
          .slice(1, -1)
          .map((column) => column.trim());
        return {
          phase: Number(phase),
          title,
          status,
          design: design?.match(/\(([^)]+)\)/)?.[1],
        };
      });

    expect(tableRows).toEqual(
      roadmapItems.map((item) => ({
        phase: item.phase,
        title: item.title,
        status: item.status,
        design: `${basename(item.designTo)}.md`,
      })),
    );
  });

  it('publishes all 34 Mini-RFCs with the required product and technical sections', () => {
    const requiredSections = [
      /产品目标/,
      /范围/,
      /API/,
      /数据模型/,
      /(?:内部模块|内部 API)/,
      /错误.*(?:性能|安全)/,
      /破坏性更新/,
      /(?:分阶段交付|实施阶段|交付阶段)/,
      /验收标准/,
      /依赖.*已决决策/,
    ];
    const ids = Object.values(miniRfcs).flat();

    expect(ids).toHaveLength(34);
    expect(new Set(ids)).toHaveProperty('size', 34);

    for (const [file, fileIds] of Object.entries(miniRfcs)) {
      const document = roadmapDocument(file);
      for (const id of fileIds) {
        const start = document.search(new RegExp(`^## ${id.replace('-', '\\-')}[ .]`, 'm'));
        expect(start, `${file} must define ${id}`).toBeGreaterThanOrEqual(0);
        const end = document.indexOf('\n## ', start + 4);
        const miniRfc = document.slice(start, end < 0 ? undefined : end);

        for (const section of requiredSections) {
          expect(miniRfc, `${file} ${id} must include ${section}`).toMatch(section);
        }

        expect(miniRfc).toContain('planned');
      }
    }
  });
});
