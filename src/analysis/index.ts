export type {
  FormulaGoalSeekApplyProposal,
  FormulaGoalSeekOptions,
  FormulaGoalSeekResult,
} from './goal-seek';
export { createFormulaGoalSeekApplyProposal, solveFormulaGoalSeek } from './goal-seek';
export type {
  PivotAggregate,
  PivotDefinition,
  PivotRefreshLimits,
  PivotRefreshOptions,
  PivotRefreshOutcome,
  PivotResult,
  PivotResultCell,
  PivotScalar,
  PivotSourceSnapshot,
} from './pivot';
export { refreshPivot } from './pivot';
export type {
  CompileSlicerOptions,
  SlicerDefinition,
  SlicerDiagnostic,
  SlicerFilterContext,
  SlicerSessionSelections,
  SlicerValueIndexEntry,
  SlicerValueIndexOptions,
} from './slicer';
export { buildSlicerValueIndex, compileSlicerFilterContext } from './slicer';
export type {
  CompiledSolverModel,
  RunSolverOptions,
  SolverConstraint,
  SolverGoal,
  SolverModelRequest,
  SolverResult,
  SolverStatus,
  SolverVariable,
} from './solver';
export { compileSolverModel, runSolver } from './solver';
export { createStructuredTableResolver, planStructuredTableAutoExpand } from './tables';
export type { StructuredTableAutoExpandResult, StructuredTableDocumentProvider } from './tables';
export { chartAffectedByChanges, chartToDisplayCommands, resolveChart } from './charts';
export type {
  ChartDefinition,
  ChartDisplayOutput,
  ChartResolution,
  ChartResolutionLimits,
  ChartSeriesDefinition,
  ChartType,
  ChartValueSource,
  NormalizedChart,
  NormalizedChartSeries,
} from './charts';
export {
  resolveSparkline,
  sparklineAffectedByChanges,
  sparklineToDisplayCommands,
} from './sparklines';
export type {
  NormalizedSparkline,
  SparklineDefinition,
  SparklineDisplayOutput,
  SparklineResolution,
  SparklineType,
  SparklineValueSource,
} from './sparklines';
