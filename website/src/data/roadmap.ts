export type RoadmapPhaseId = 0 | 1 | 2 | 3 | 4;
export type RoadmapStatus = 'planned';

export interface RoadmapPhase {
  readonly id: RoadmapPhaseId;
  readonly label: string;
  readonly title: string;
  readonly summary: string;
}

export interface RoadmapItem {
  readonly id: string;
  readonly phase: RoadmapPhaseId;
  readonly title: string;
  readonly summary: string;
  readonly status: RoadmapStatus;
  readonly designTo: `/docs/roadmap/${string}`;
}

export const roadmapPhases = [
  {
    id: 0,
    label: 'Foundation',
    title: 'A deterministic document core',
    summary:
      'Replace mutable legacy state with typed documents, atomic commands and shared render semantics.',
  },
  {
    id: 1,
    label: 'Template Print MVP',
    title: 'Print only the sheet or range that matters',
    summary:
      'Turn spreadsheet regions into predictable templates, previews and isolated browser print jobs.',
  },
  {
    id: 2,
    label: 'Document Generation',
    title: 'From one template to production outputs',
    summary:
      'Add advanced repeats, managed resources and deterministic PDF, XLSX and image adapters.',
  },
  {
    id: 3,
    label: 'Spreadsheet Depth',
    title: 'The capabilities expected from a modern spreadsheet',
    summary:
      'Grow formulas, data tooling, interchange, analysis and visualization on the same document model.',
  },
  {
    id: 4,
    label: 'SDK Ecosystem',
    title: 'Stable extension points for host applications',
    summary:
      'Expose versioned plugins and adapters while keeping storage, collaboration and AI host-owned.',
  },
] as const satisfies readonly RoadmapPhase[];

const item = (
  id: string,
  phase: RoadmapPhaseId,
  title: string,
  summary: string,
  designTo: RoadmapItem['designTo'],
): RoadmapItem => ({ id, phase, title, summary, status: 'planned', designTo });

export const roadmapItems = [
  item(
    'workbook-2',
    0,
    'Workbook 2.0 typed document model',
    'Stable IDs, typed cells and deterministic serialization.',
    '/docs/roadmap/foundation',
  ),
  item(
    'transactions',
    0,
    'Atomic Command / Transaction',
    'One validated and reversible mutation boundary for every edit.',
    '/docs/roadmap/foundation',
  ),
  item(
    'formula-format-core',
    0,
    'Formula dependency and number-format core',
    'Excel-aligned values, dependencies, errors and display text.',
    '/docs/roadmap/foundation',
  ),
  item(
    'render-semantics',
    0,
    'Shared render semantics and Canvas accessibility',
    'One presentation model for screen, accessibility and print.',
    '/docs/roadmap/foundation',
  ),
  item(
    'extension-kernel',
    0,
    'Minimal cell-type and adapter registry kernel',
    'Typed internal composition without exposing an unsafe plugin surface.',
    '/docs/roadmap/foundation',
  ),
  item(
    'print-targets',
    1,
    'Sheet, selection and range print targets',
    'Print a chosen sheet or region without editor chrome.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'template-bindings',
    1,
    'Safe scalar bindings, repeat rows and conditional ranges',
    'Bind business data through a constrained template DSL.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'pagination',
    1,
    'Deterministic pagination and print profiles',
    'Make paper, margins, scaling, titles and page breaks explicit.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'print-preview',
    1,
    'Page preview and isolated browser printing',
    'Preview the exact pages rendered inside an isolated print document.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'advanced-repeats',
    2,
    'Nested, horizontal, range and page repeats',
    'Generate complex reports without arbitrary template code.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'resource-pipeline',
    2,
    'Image, font, QR code and async resource pipeline',
    'Resolve bounded external resources before deterministic layout.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'pdf-output',
    2,
    'PDF Blob output',
    'Translate the shared page display list into downloadable PDF data.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'xlsx-output',
    2,
    'XLSX template output',
    'Preserve semantic workbook content and print profiles in OOXML.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'image-output',
    2,
    'SVG and PNG page output',
    'Render selected generated pages as vector or bitmap images.',
    '/docs/roadmap/template-printing',
  ),
  item(
    'conditional-formatting',
    3,
    'Conditional formatting',
    'Apply deterministic visual rules without mutating base styles.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'advanced-validation',
    3,
    'Advanced validation, dropdown and checkbox cells',
    'Combine typed rules with accessible interactive cell controls.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'formula-library',
    3,
    'Expanded function library and cross-sheet references',
    'Broaden Excel-compatible calculation with explicit support levels.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'array-formulas',
    3,
    'Named ranges, array and spill formulas',
    'Model reusable names and dynamic multi-cell results safely.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'saved-views',
    3,
    'Multi-column sort, conditional filter and saved views',
    'Separate reusable data views from destructive row changes.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'data-cleanup',
    3,
    'Grouping, deduplication, text split and data cleanup',
    'Preview data transformations before one atomic command commits them.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'file-interchange',
    3,
    'CSV/TSV, XLSX and ODS interchange',
    'Import and export common spreadsheet formats with diagnostics.',
    '/docs/roadmap/formulas-data',
  ),
  item(
    'structured-tables',
    3,
    'Structured tables and structured references',
    'Give tabular regions stable identity, columns and formula semantics.',
    '/docs/roadmap/analysis-visualization',
  ),
  item(
    'charts',
    3,
    'Charts and Sparklines',
    'Render data-bound visualizations through deterministic series models.',
    '/docs/roadmap/analysis-visualization',
  ),
  item(
    'objects',
    3,
    'Images, shapes, text boxes and anchored objects',
    'Keep floating content attached through structural sheet edits.',
    '/docs/roadmap/analysis-visualization',
  ),
  item(
    'pivot-slicer',
    3,
    'PivotTable and Slicer',
    'Build refreshable aggregations and connected filter controls.',
    '/docs/roadmap/analysis-visualization',
  ),
  item(
    'solver',
    3,
    'Goal Seek and pluggable Solver',
    'Express bounded optimization without embedding a solver service.',
    '/docs/roadmap/analysis-visualization',
  ),
  item(
    'cell-sdk',
    4,
    'Public structured cell renderer/editor plugin SDK',
    'Open the typed cell protocol to trusted host extensions.',
    '/docs/roadmap/extensibility',
  ),
  item(
    'template-sdk',
    4,
    'Versioned Template Module SDK',
    'Extend compilation through validated intermediate-representation hooks.',
    '/docs/roadmap/extensibility',
  ),
  item(
    'adapter-sdk',
    4,
    'Public adapter lifecycle, trust policy and compatibility SDK',
    'Make environment, capabilities and cleanup explicit for integrations.',
    '/docs/roadmap/extensibility',
  ),
  item(
    'persistence-history',
    4,
    'Persistence and version history adapters',
    'Let hosts save snapshots and expose durable revision history.',
    '/docs/roadmap/host-integrations',
  ),
  item(
    'collaboration',
    4,
    'Collaboration and remote selection adapters',
    'Integrate host-owned synchronization without selecting a CRDT backend.',
    '/docs/roadmap/host-integrations',
  ),
  item(
    'permission-comments',
    4,
    'Permission and comment adapters',
    'Connect authorization and discussion while preserving command checks.',
    '/docs/roadmap/host-integrations',
  ),
  item(
    'ai-commands',
    4,
    'Validated AI command proposals',
    'Require schema validation, dry-run and user confirmation before changes.',
    '/docs/roadmap/host-integrations',
  ),
] as const satisfies readonly RoadmapItem[];

export function groupRoadmapItems(): ReadonlyMap<RoadmapPhaseId, readonly RoadmapItem[]> {
  return new Map(
    roadmapPhases.map((phase) => [
      phase.id,
      roadmapItems.filter((roadmapItem) => roadmapItem.phase === phase.id),
    ]),
  );
}
