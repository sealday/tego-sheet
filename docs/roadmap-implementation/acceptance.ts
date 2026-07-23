import {
  allRoadmapItems,
  type RoadmapItemId,
  type RoadmapPhaseId,
  type RoadmapStatus,
} from '../../website/src/data/roadmap';

export interface RoadmapAcceptanceEntry {
  readonly id: RoadmapItemId;
  readonly phase: RoadmapPhaseId;
  readonly spec: `website/docs/roadmap/${string}.md#${string}`;
  readonly tasks: readonly string[];
  readonly tests: readonly string[];
  readonly state: RoadmapStatus;
}

const canonicalItemsById = new Map(allRoadmapItems.map((item) => [item.id, item]));

const entry = (
  id: RoadmapItemId,
  spec: RoadmapAcceptanceEntry['spec'],
  task: string,
  test: string,
): RoadmapAcceptanceEntry => {
  const canonicalItem = canonicalItemsById.get(id);
  if (!canonicalItem) {
    throw new RangeError(`Unknown Roadmap item: ${id}`);
  }
  return {
    id,
    phase: canonicalItem.phase,
    spec,
    tasks: [task],
    tests: [test],
    state: canonicalItem.status,
  };
};

const documentTests =
  'npx vitest run --project unit tests/unit/document tests/unit/core/schema-validation.test.ts tests/unit/core/serialization.test.ts && npm run typecheck && npm run test:ssr';
const extensionTests =
  'npx vitest run --project unit tests/unit/extensions && npm run typecheck && npm run lint';
const transactionTests =
  'npx vitest run --project unit tests/unit/document tests/unit/core tests/unit/react tests/property/document && npm run typecheck && npm run lint';
const formulaTests =
  'npx vitest run --project unit tests/unit/formula tests/unit/format tests/unit/core/formulas.test.ts tests/unit/core/structure.test.ts && npm run typecheck && npm run lint';
const presentationTests =
  'npx vitest run --project unit --project component --project architecture tests/unit/presentation tests/unit/engine tests/component/accessibility-grid.test.tsx tests/architecture && npm run test:browser && npm run typecheck';
const templateTests =
  'npx vitest run --project unit --project component --project architecture tests/unit/template tests/unit/print tests/component/template-designer.test.tsx tests/architecture && npm run test:browser && npm run test:visual && npm run typecheck && npm run lint';
const resourceTests =
  'npx vitest run --project unit tests/unit/template tests/unit/resources tests/unit/print && npm run typecheck && npm run lint';
const outputTests =
  'npx vitest run --project unit tests/unit/output && npm run test:browser && npm run test:visual && npm run test:package && npm run build';
const spreadsheetTests =
  'npx vitest run --project unit tests/unit/format tests/unit/validation tests/unit/formula tests/unit/views tests/unit/data-tools tests/unit/interchange tests/unit/objects && npm run test:browser && npm run typecheck && npm run lint';
const analysisTests =
  'npx vitest run --project unit --project component tests/unit/analysis tests/component/analysis && npm run test:browser && npm run test:visual && npm run typecheck && npm run lint';
const sdkTests =
  'npx vitest run --project unit tests/unit/sdk && npm run typecheck && npm run test:package && npm run build';
const integrationTests =
  'npx vitest run --project unit --project component tests/contract/integrations tests/component/integrations && npm run test:browser && npm run typecheck && npm run lint';

export const roadmapAcceptance = [
  entry(
    'workbook-2',
    'website/docs/roadmap/foundation.md#f1-workbook-20',
    'Task 2: Deliver F1 Workbook 2.0',
    documentTests,
  ),
  entry(
    'transactions',
    'website/docs/roadmap/foundation.md#f2-command--transaction',
    'Task 4: Deliver F2 Atomic Command and Transaction',
    transactionTests,
  ),
  entry(
    'formula-format-core',
    'website/docs/roadmap/foundation.md#f3-formula--format-core',
    'Task 5: Deliver F3 Formula Dependency and Number-Format Core',
    formulaTests,
  ),
  entry(
    'render-semantics',
    'website/docs/roadmap/foundation.md#f4-render-semantics--accessibility',
    'Task 6: Deliver F4 Shared Presentation, Accessibility, and Print Display List',
    presentationTests,
  ),
  entry(
    'extension-kernel',
    'website/docs/roadmap/foundation.md#f5-minimal-extension--adapter-kernel',
    'Task 3: Deliver F5 Minimal Extension and Adapter Kernel',
    extensionTests,
  ),
  entry(
    'print-targets',
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 print targets',
    templateTests,
  ),
  entry(
    'template-bindings',
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 safe template bindings',
    templateTests,
  ),
  entry(
    'pagination',
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 deterministic pagination',
    templateTests,
  ),
  entry(
    'print-preview',
    'website/docs/roadmap/template-printing.md#tp1-template-print-mvp',
    'Task 7: Deliver TP1 preview and isolated browser printing',
    templateTests,
  ),
  entry(
    'advanced-repeats',
    'website/docs/roadmap/template-printing.md#tp2-advanced-template-structure',
    'Task 8: Deliver TP2 Advanced Template Structures',
    resourceTests,
  ),
  entry(
    'resource-pipeline',
    'website/docs/roadmap/template-printing.md#tp3-resource-pipeline',
    'Task 8: Deliver TP3 Resource Pipeline',
    resourceTests,
  ),
  entry(
    'pdf-output',
    'website/docs/roadmap/template-printing.md#tp4-pdf-adapter',
    'Task 9: Deliver TP4 PDF Output',
    outputTests,
  ),
  entry(
    'xlsx-output',
    'website/docs/roadmap/template-printing.md#tp5-xlsx-template-adapter',
    'Task 9: Deliver TP5 XLSX Output',
    outputTests,
  ),
  entry(
    'image-output',
    'website/docs/roadmap/template-printing.md#tp6-image-adapter',
    'Task 9: Deliver TP6 Image Output',
    outputTests,
  ),
  entry(
    'conditional-formatting',
    'website/docs/roadmap/formulas-data.md#fmt-01-数字格式与条件格式',
    'Task 10: Deliver FMT-01 Conditional Formatting',
    spreadsheetTests,
  ),
  entry(
    'advanced-validation',
    'website/docs/roadmap/formulas-data.md#val-01-数据验证与交互式单元格',
    'Task 10: Deliver VAL-01 Validation and Interactive Cells',
    spreadsheetTests,
  ),
  entry(
    'formula-library',
    'website/docs/roadmap/formulas-data.md#frm-01-高级公式引擎',
    'Task 10: Deliver FRM-01 expanded formulas and cross-sheet references',
    spreadsheetTests,
  ),
  entry(
    'array-formulas',
    'website/docs/roadmap/formulas-data.md#frm-01-高级公式引擎',
    'Task 10: Deliver FRM-01 named ranges, arrays, and spill formulas',
    spreadsheetTests,
  ),
  entry(
    'saved-views',
    'website/docs/roadmap/formulas-data.md#view-01-排序筛选与保存视图',
    'Task 10: Deliver VIEW-01 Saved Views',
    spreadsheetTests,
  ),
  entry(
    'data-cleanup',
    'website/docs/roadmap/formulas-data.md#data-01-数据整理与清洗命令',
    'Task 10: Deliver DATA-01 Data Cleanup',
    spreadsheetTests,
  ),
  entry(
    'file-interchange',
    'website/docs/roadmap/formulas-data.md#io-01-csvxlsx-与-ods-文件互操作',
    'Task 10: Deliver IO-01 File Interchange',
    spreadsheetTests,
  ),
  entry(
    'structured-tables',
    'website/docs/roadmap/analysis-visualization.md#tbl-01-结构化表格',
    'Task 11: Deliver TBL-01 Structured Tables',
    analysisTests,
  ),
  entry(
    'charts',
    'website/docs/roadmap/analysis-visualization.md#cht-01-图表',
    'Task 11: Deliver CHT-01 Charts and SPK-01 Sparklines',
    analysisTests,
  ),
  entry(
    'objects',
    'website/docs/roadmap/analysis-visualization.md#obj-01-浮动对象与锚点',
    'Task 10: Deliver OBJ-01 Anchored Objects',
    spreadsheetTests,
  ),
  entry(
    'pivot-slicer',
    'website/docs/roadmap/analysis-visualization.md#pvt-01-pivottable',
    'Task 11: Deliver PVT-01 PivotTable and SLC-01 Slicer',
    analysisTests,
  ),
  entry(
    'solver',
    'website/docs/roadmap/analysis-visualization.md#gsk-01-goal-seek',
    'Task 11: Deliver GSK-01 Goal Seek and SLV-01 Solver',
    analysisTests,
  ),
  entry(
    'cell-sdk',
    'website/docs/roadmap/extensibility.md#e1-cell-extension-sdk',
    'Task 12: Deliver E1 Cell Extension SDK',
    sdkTests,
  ),
  entry(
    'template-sdk',
    'website/docs/roadmap/extensibility.md#e2-template-module-sdk',
    'Task 12: Deliver E2 Template Module SDK',
    sdkTests,
  ),
  entry(
    'adapter-sdk',
    'website/docs/roadmap/extensibility.md#e3-adapter-registry',
    'Task 12: Deliver E3 Adapter Registry',
    sdkTests,
  ),
  entry(
    'persistence-history',
    'website/docs/roadmap/host-integrations.md#h1-persistence-adapter',
    'Task 13: Deliver H1 Persistence and H5 Version History Adapters',
    integrationTests,
  ),
  entry(
    'collaboration',
    'website/docs/roadmap/host-integrations.md#h2-collaboration-adapter',
    'Task 13: Deliver H2 Collaboration Adapter',
    integrationTests,
  ),
  entry(
    'permission-comments',
    'website/docs/roadmap/host-integrations.md#h3-permission-adapter',
    'Task 13: Deliver H3 Permission and H4 Comments Adapters',
    integrationTests,
  ),
  entry(
    'ai-commands',
    'website/docs/roadmap/host-integrations.md#h6-ai-command-adapter',
    'Task 13: Deliver H6 AI Command Adapter',
    integrationTests,
  ),
] as const satisfies readonly RoadmapAcceptanceEntry[];

export type RoadmapAcceptanceId = (typeof roadmapAcceptance)[number]['id'];

export function getRoadmapDeliveryState(id: RoadmapAcceptanceId): RoadmapAcceptanceEntry['state'] {
  const canonicalItem = canonicalItemsById.get(id);
  if (!canonicalItem) {
    throw new RangeError(`Unknown Roadmap item: ${id}`);
  }
  return canonicalItem.status;
}
