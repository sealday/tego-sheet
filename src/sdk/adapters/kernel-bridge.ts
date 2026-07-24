import type { WorkbookReader, WorkbookWriter } from '../../interchange/contracts';
import type { ResourceResolver } from '../../template/resources';
import type {
  AiCommandAdapter,
  ChartRendererAdapter,
  CollaborationAdapter,
  CommentsAdapter,
  FormulaFunctionProviderAdapter,
  OutputAdapter,
  PermissionAdapter,
  PersistenceAdapter,
  SolverAdapter,
  VersionHistoryAdapter,
} from './types';

declare module '../../extensions/kernel/capabilities' {
  interface KernelCapabilities {
    'workbook-reader': WorkbookReader;
    'workbook-writer': WorkbookWriter;
    output: OutputAdapter;
    'chart-renderer': ChartRendererAdapter;
    'formula-function-provider': FormulaFunctionProviderAdapter;
    solver: SolverAdapter;
    persistence: PersistenceAdapter;
    collaboration: CollaborationAdapter;
    permission: PermissionAdapter;
    comments: CommentsAdapter;
    'version-history': VersionHistoryAdapter;
    'ai-command': AiCommandAdapter;
    'resource-resolver': ResourceResolver;
  }
}

export {};
