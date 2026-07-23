import type { RoadmapItemId, RoadmapPhaseId } from '../../website/src/data/roadmap';

export interface RoadmapAcceptanceEntry {
  readonly id: string;
  readonly phase: 0 | 1 | 2 | 3 | 4;
  readonly spec: `website/docs/roadmap/${string}.md#${string}`;
  readonly tasks: readonly string[];
  readonly tests: readonly string[];
  readonly state: 'planned' | 'in-progress' | 'shipped';
}

const entry = (
  id: RoadmapItemId,
  phase: RoadmapPhaseId,
  spec: RoadmapAcceptanceEntry['spec'],
  task: string,
  test: string,
): RoadmapAcceptanceEntry & { readonly id: RoadmapItemId } => ({
  id,
  phase,
  spec,
  tasks: [task],
  tests: [test],
  state: 'planned',
});

const documentTests = 'npx vitest run --project unit tests/unit/document';
const extensionTests = 'npx vitest run --project unit tests/unit/extensions';
const formulaTests = 'npx vitest run --project unit tests/unit/formula tests/unit/format';
const presentationTests =
  'npx vitest run --project unit --project component --project architecture tests/unit/presentation tests/component/accessibility-grid.test.tsx tests/architecture/renderer-boundaries.test.ts';
const templateTests =
  'npx vitest run --project unit --project component tests/unit/template tests/unit/print tests/component/template-designer.test.tsx';
const resourceTests =
  'npx vitest run --project unit tests/unit/template/advanced-structures.test.ts tests/unit/resources/resource-pipeline.test.ts';
const outputTests = 'npx vitest run --project unit tests/unit/output';
const spreadsheetTests =
  'npx vitest run --project unit tests/unit/format tests/unit/validation tests/unit/formula/advanced-formulas.test.ts tests/unit/views tests/unit/data-tools tests/unit/interchange tests/unit/objects';
const analysisTests =
  'npx vitest run --project unit --project component tests/unit/analysis tests/component/analysis';
const sdkTests = 'npx vitest run --project unit tests/unit/sdk';
const integrationTests =
  'npx vitest run --project unit --project component tests/contract/integrations tests/component/integrations';

export const roadmapAcceptance = [
  entry(
    'workbook-2',
    0,
    'website/docs/roadmap/foundation.md#f1-workbook-20',
    'Task 2: Deliver F1 Workbook 2.0',
    documentTests,
  ),
  entry(
    'transactions',
    0,
    'website/docs/roadmap/foundation.md#f2-command--transaction',
    'Task 4: Deliver F2 Atomic Command and Transaction',
    'npx vitest run --project unit tests/unit/document/transactions.test.ts tests/unit/document/coordinate-transform.test.ts',
  ),
  entry(
    'formula-format-core',
    0,
    'website/docs/roadmap/foundation.md#f3-formula--format-core',
    'Task 5: Deliver F3 Formula Dependency and Number-Format Core',
    formulaTests,
  ),
  entry(
    'render-semantics',
    0,
    'website/docs/roadmap/foundation.md#f4-render-semantics--accessibility',
    'Task 6: Deliver F4 Shared Presentation, Accessibility, and Print Display List',
    presentationTests,
  ),
  entry(
    'extension-kernel',
    0,
    'website/docs/roadmap/foundation.md#f5-minimal-extension--adapter-kernel',
    'Task 3: Deliver F5 Minimal Extension and Adapter Kernel',
    extensionTests,
  ),
  entry(
    'print-targets',
    1,
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 print targets',
    templateTests,
  ),
  entry(
    'template-bindings',
    1,
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 safe template bindings',
    templateTests,
  ),
  entry(
    'pagination',
    1,
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 deterministic pagination',
    templateTests,
  ),
  entry(
    'print-preview',
    1,
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 preview and isolated browser printing',
    templateTests,
  ),
  entry(
    'advanced-repeats',
    2,
    'website/docs/roadmap/template-printing.md#tp2-advanced-template-structure',
    'Task 8: Deliver TP2 Advanced Template Structures',
    resourceTests,
  ),
  entry(
    'resource-pipeline',
    2,
    'website/docs/roadmap/template-printing.md#tp3-resource-pipeline',
    'Task 8: Deliver TP3 Resource Pipeline',
    resourceTests,
  ),
  entry(
    'pdf-output',
    2,
    'website/docs/roadmap/template-printing.md#tp4-pdf-adapter',
    'Task 9: Deliver TP4 PDF Output',
    outputTests,
  ),
  entry(
    'xlsx-output',
    2,
    'website/docs/roadmap/template-printing.md#tp5-xlsx-template-adapter',
    'Task 9: Deliver TP5 XLSX Output',
    outputTests,
  ),
  entry(
    'image-output',
    2,
    'website/docs/roadmap/template-printing.md#tp6-image-adapter',
    'Task 9: Deliver TP6 Image Output',
    outputTests,
  ),
  entry(
    'conditional-formatting',
    3,
    'website/docs/roadmap/formulas-data.md#fmt-01-数字格式与条件格式',
    'Task 10: Deliver FMT-01 Conditional Formatting',
    spreadsheetTests,
  ),
  entry(
    'advanced-validation',
    3,
    'website/docs/roadmap/formulas-data.md#val-01-数据验证与交互式单元格',
    'Task 10: Deliver VAL-01 Validation and Interactive Cells',
    spreadsheetTests,
  ),
  entry(
    'formula-library',
    3,
    'website/docs/roadmap/formulas-data.md#frm-01-高级公式引擎',
    'Task 10: Deliver FRM-01 expanded formulas and cross-sheet references',
    spreadsheetTests,
  ),
  entry(
    'array-formulas',
    3,
    'website/docs/roadmap/formulas-data.md#frm-01-高级公式引擎',
    'Task 10: Deliver FRM-01 named ranges, arrays, and spill formulas',
    spreadsheetTests,
  ),
  entry(
    'saved-views',
    3,
    'website/docs/roadmap/formulas-data.md#view-01-排序筛选与保存视图',
    'Task 10: Deliver VIEW-01 Saved Views',
    spreadsheetTests,
  ),
  entry(
    'data-cleanup',
    3,
    'website/docs/roadmap/formulas-data.md#data-01-数据整理与清洗命令',
    'Task 10: Deliver DATA-01 Data Cleanup',
    spreadsheetTests,
  ),
  entry(
    'file-interchange',
    3,
    'website/docs/roadmap/formulas-data.md#io-01-csvxlsx-与-ods-文件互操作',
    'Task 10: Deliver IO-01 File Interchange',
    spreadsheetTests,
  ),
  entry(
    'structured-tables',
    3,
    'website/docs/roadmap/analysis-visualization.md#tbl-01-结构化表格',
    'Task 11: Deliver TBL-01 Structured Tables',
    analysisTests,
  ),
  entry(
    'charts',
    3,
    'website/docs/roadmap/analysis-visualization.md#cht-01-图表',
    'Task 11: Deliver CHT-01 Charts and SPK-01 Sparklines',
    analysisTests,
  ),
  entry(
    'objects',
    3,
    'website/docs/roadmap/analysis-visualization.md#obj-01-浮动对象与锚点',
    'Task 10: Deliver OBJ-01 Anchored Objects',
    spreadsheetTests,
  ),
  entry(
    'pivot-slicer',
    3,
    'website/docs/roadmap/analysis-visualization.md#pvt-01-pivottable',
    'Task 11: Deliver PVT-01 PivotTable and SLC-01 Slicer',
    analysisTests,
  ),
  entry(
    'solver',
    3,
    'website/docs/roadmap/analysis-visualization.md#gsk-01-goal-seek',
    'Task 11: Deliver GSK-01 Goal Seek and SLV-01 Solver',
    analysisTests,
  ),
  entry(
    'cell-sdk',
    4,
    'website/docs/roadmap/extensibility.md#e1-cell-extension-sdk',
    'Task 12: Deliver E1 Cell Extension SDK',
    sdkTests,
  ),
  entry(
    'template-sdk',
    4,
    'website/docs/roadmap/extensibility.md#e2-template-module-sdk',
    'Task 12: Deliver E2 Template Module SDK',
    sdkTests,
  ),
  entry(
    'adapter-sdk',
    4,
    'website/docs/roadmap/extensibility.md#e3-adapter-registry',
    'Task 12: Deliver E3 Adapter Registry',
    sdkTests,
  ),
  entry(
    'persistence-history',
    4,
    'website/docs/roadmap/host-integrations.md#h1-persistence-adapter',
    'Task 13: Deliver H1 Persistence and H5 Version History Adapters',
    integrationTests,
  ),
  entry(
    'collaboration',
    4,
    'website/docs/roadmap/host-integrations.md#h2-collaboration-adapter',
    'Task 13: Deliver H2 Collaboration Adapter',
    integrationTests,
  ),
  entry(
    'permission-comments',
    4,
    'website/docs/roadmap/host-integrations.md#h3-permission-adapter',
    'Task 13: Deliver H3 Permission and H4 Comments Adapters',
    integrationTests,
  ),
  entry(
    'ai-commands',
    4,
    'website/docs/roadmap/host-integrations.md#h6-ai-command-adapter',
    'Task 13: Deliver H6 AI Command Adapter',
    integrationTests,
  ),
] as const satisfies readonly RoadmapAcceptanceEntry[];

export type RoadmapAcceptanceId = (typeof roadmapAcceptance)[number]['id'];

export function getRoadmapDeliveryState(id: RoadmapAcceptanceId): RoadmapAcceptanceEntry['state'] {
  const acceptance = roadmapAcceptance.find((candidate) => candidate.id === id);
  if (!acceptance) {
    throw new RangeError(`Unknown Roadmap item: ${id}`);
  }
  return acceptance.state;
}
