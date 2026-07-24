import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type { CompiledTemplate, TemplateIRBinding, TemplateRegionNode } from '../model';
import { structuralAxis } from './advanced-axis';
import type { ExpansionScope } from './advanced-internals';

export type AdvancedBinding = Extract<
  TemplateIRBinding,
  {
    readonly type:
      | 'repeat-columns'
      | 'repeat-rows'
      | 'repeat-range'
      | 'repeat-page'
      | 'repeat-sheet'
      | 'subtemplate';
  }
>;

export interface PlannedAdvancedBinding {
  readonly binding: AdvancedBinding;
  readonly items: readonly unknown[];
}

export interface ExpansionPlan {
  readonly advanced: readonly AdvancedBinding[];
  readonly tp1Bindings: readonly TemplateIRBinding[];
  readonly handledByNested: (binding: TemplateIRBinding) => boolean;
}

export interface AllocationEstimate {
  readonly cells: number;
  readonly rows: number;
  readonly columns: number;
}

export function createExpansionPlan(compiled: CompiledTemplate): ExpansionPlan {
  const bindingById = new Map(compiled.ir.bindings.map((binding) => [binding.id, binding]));
  const axis = (node: TemplateRegionNode) => structuralAxis(bindingById.get(node.bindingId));
  const horizontalNestedIds = new Set<string>();
  const mixedTreeIds = new Set<string>();
  const collectMixedTree = (node: TemplateRegionNode): void => {
    mixedTreeIds.add(node.bindingId);
    node.children.forEach(collectMixedTree);
  };
  const isMixedTree = (node: TemplateRegionNode): boolean => {
    return axis(node) !== undefined && node.children.every(isMixedTree);
  };
  const hasVerticalOrConditionalDescendant = (node: TemplateRegionNode): boolean =>
    node.children.some(
      (child) => axis(child) !== 'horizontal' || hasVerticalOrConditionalDescendant(child),
    );
  const mixedRoots = (compiled.ir.regionTree ?? []).filter((node) => {
    if (!isMixedTree(node)) return false;
    const binding = bindingById.get(node.bindingId);
    if (binding?.type === 'repeat-range' && node.children.length === 0) return false;
    const type = axis(node);
    return (
      type === 'vertical' || (type === 'horizontal' && hasVerticalOrConditionalDescendant(node))
    );
  });
  mixedRoots.forEach(collectMixedTree);
  const collectHorizontal = (node: TemplateRegionNode): boolean => {
    if (axis(node) !== 'horizontal' || !node.children.every(collectHorizontal)) {
      return false;
    }
    horizontalNestedIds.add(node.bindingId);
    node.children.forEach((child) => horizontalNestedIds.add(child.bindingId));
    return true;
  };
  (compiled.ir.regionTree ?? []).forEach(collectHorizontal);
  const advanced = compiled.ir.bindings.filter(
    (binding): binding is AdvancedBinding =>
      (binding.type === 'repeat-columns' ||
        (binding.type === 'repeat-rows' &&
          binding.objects !== undefined &&
          binding.objects.length > 0) ||
        binding.type === 'repeat-range' ||
        binding.type === 'repeat-page' ||
        binding.type === 'repeat-sheet' ||
        binding.type === 'subtemplate') &&
      !mixedTreeIds.has(binding.id) &&
      !horizontalNestedIds.has(binding.id),
  );
  const tp1Bindings = compiled.ir.bindings.filter(
    (binding) =>
      binding.type === 'value' ||
      (binding.type === 'repeat-rows' &&
        (binding.objects === undefined || binding.objects.length === 0)) ||
      binding.type === 'conditional-range',
  );
  const structuralRanges = compiled.ir.bindings.flatMap((binding) =>
    binding.type === 'repeat-rows' ||
    binding.type === 'repeat-columns' ||
    binding.type === 'repeat-range' ||
    binding.type === 'repeat-page' ||
    binding.type === 'repeat-sheet' ||
    binding.type === 'subtemplate'
      ? [binding.range]
      : [],
  );
  return {
    advanced,
    tp1Bindings,
    handledByNested: (binding) =>
      mixedTreeIds.has(binding.id) ||
      ('target' in binding &&
        structuralRanges.some(
          (range) =>
            binding.target.sheetId === range.sheetId &&
            binding.target.row >= range.start.row &&
            binding.target.row <= range.end.row &&
            binding.target.column >= range.start.column &&
            binding.target.column <= range.end.column,
        )),
  };
}

export function materializeAdvancedBindings(
  bindings: readonly AdvancedBinding[],
  data: unknown,
  formatters: TemplateFormatterRegistry,
): readonly PlannedAdvancedBinding[] {
  const rootScope: ExpansionScope = { root: data };
  return bindings.map((binding) => {
    const value = evaluateTemplateExpression(binding.source, rootScope, formatters);
    return { binding, items: Array.isArray(value) ? value : [] };
  });
}

export function estimateAllocation(
  compiled: CompiledTemplate,
  estimates: readonly PlannedAdvancedBinding[],
): AllocationEstimate {
  const sourceCells = compiled.sourceDocument.workbook.sheets.reduce(
    (count, sheet) => count + sheet.cells.length,
    0,
  );
  let rows = 0;
  let columns = 0;
  const cells = estimates.reduce((count, { binding, items }) => {
    const sourceSheet = compiled.sourceDocument.workbook.sheets.find(
      ({ id }) => id === binding.range.sheetId,
    );
    const sourceCells =
      sourceSheet?.cells.filter(
        ({ row, column }) =>
          row >= binding.range.start.row &&
          row <= binding.range.end.row &&
          column >= binding.range.start.column &&
          column <= binding.range.end.column,
      ) ?? [];
    const sourceCoordinates = new Set(sourceCells.map(({ row, column }) => `${row}:${column}`));
    const createdValueCells = compiled.ir.bindings.filter(
      (candidate) =>
        candidate.type === 'value' &&
        candidate.target.sheetId === binding.range.sheetId &&
        candidate.target.row >= binding.range.start.row &&
        candidate.target.row <= binding.range.end.row &&
        candidate.target.column >= binding.range.start.column &&
        candidate.target.column <= binding.range.end.column &&
        !sourceCoordinates.has(`${candidate.target.row}:${candidate.target.column}`),
    ).length;
    const effectiveSourceCells = sourceCells.length + createdValueCells;
    const matrixRows = items.length;
    const matrixColumns = Math.max(
      0,
      ...items.map((item) => (Array.isArray(item) ? item.length : 1)),
    );
    const multiplier =
      binding.type === 'subtemplate'
        ? 1
        : binding.type === 'repeat-range' && binding.axis === 'both' && items.every(Array.isArray)
          ? matrixRows * matrixColumns
          : items.length;
    const height = binding.range.end.row - binding.range.start.row + 1;
    const width = binding.range.end.column - binding.range.start.column + 1;
    if (
      binding.type === 'repeat-page' ||
      binding.type === 'repeat-rows' ||
      (binding.type === 'repeat-range' && binding.axis === 'vertical')
    ) {
      rows += height * multiplier;
    } else if (
      binding.type === 'repeat-columns' ||
      (binding.type === 'repeat-range' && binding.axis === 'horizontal')
    ) {
      columns += width * multiplier;
    } else if (binding.type === 'repeat-range' && binding.axis === 'both') {
      rows += height * matrixRows;
      columns += width * matrixColumns;
    }
    if (binding.type === 'repeat-sheet') {
      const sheetCells = sourceSheet?.cells.length ?? 0;
      return count + sheetCells * multiplier;
    }
    return count - sourceCells.length + effectiveSourceCells * multiplier;
  }, sourceCells);
  return { cells, rows, columns };
}
