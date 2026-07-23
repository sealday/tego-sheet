import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type { CompiledTemplate, TemplateIRBinding, TemplateRegionNode } from '../model';
import type { ExpansionScope } from './advanced-internals';

export type AdvancedBinding = Extract<
  TemplateIRBinding,
  {
    readonly type:
      | 'repeat-columns'
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
  const bindingType = new Map(compiled.ir.bindings.map((binding) => [binding.id, binding.type]));
  const horizontalNestedIds = new Set<string>();
  const rowTreeIds = new Set<string>();
  const collectRowTree = (node: TemplateRegionNode): void => {
    rowTreeIds.add(node.bindingId);
    node.children.forEach(collectRowTree);
  };
  (compiled.ir.regionTree ?? [])
    .filter((node) => bindingType.get(node.bindingId) === 'repeat-rows')
    .forEach(collectRowTree);
  const collectHorizontal = (node: TemplateRegionNode): boolean => {
    if (
      bindingType.get(node.bindingId) !== 'repeat-columns' ||
      !node.children.every(collectHorizontal)
    ) {
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
        binding.type === 'repeat-range' ||
        binding.type === 'repeat-page' ||
        binding.type === 'repeat-sheet' ||
        binding.type === 'subtemplate') &&
      !rowTreeIds.has(binding.id) &&
      !horizontalNestedIds.has(binding.id),
  );
  const tp1Bindings = compiled.ir.bindings.filter(
    (binding) =>
      binding.type === 'value' ||
      binding.type === 'repeat-rows' ||
      binding.type === 'conditional-range',
  );
  const nestedRanges = compiled.ir.bindings.flatMap((binding) =>
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
      nestedRanges.some((range) =>
        'target' in binding
          ? binding.target.sheetId === range.sheetId &&
            binding.target.row >= range.start.row &&
            binding.target.row <= range.end.row
          : binding.range.sheetId === range.sheetId &&
            binding.range.start.row >= range.start.row &&
            binding.range.end.row <= range.end.row,
      ),
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
    const source =
      compiled.sourceDocument.workbook.sheets
        .find(({ id }) => id === binding.range.sheetId)
        ?.cells.filter(
          ({ row, column }) =>
            row >= binding.range.start.row &&
            row <= binding.range.end.row &&
            column >= binding.range.start.column &&
            column <= binding.range.end.column,
        ).length ?? 0;
    const matrixRows = items.length;
    const matrixColumns = Math.max(
      0,
      ...items.map((item) => (Array.isArray(item) ? item.length : 1)),
    );
    const multiplier =
      binding.type === 'repeat-range' && binding.axis === 'both' && items.every(Array.isArray)
        ? matrixRows * matrixColumns
        : items.length;
    const height = binding.range.end.row - binding.range.start.row + 1;
    const width = binding.range.end.column - binding.range.start.column + 1;
    if (
      binding.type === 'repeat-page' ||
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
    return count + source * Math.max(0, multiplier - 1);
  }, sourceCells);
  return { cells, rows, columns };
}
