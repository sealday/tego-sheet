import type { Diagnostic, DocumentCellRange, SpreadsheetDocument } from '../../document';
import { TemplateExpressionError, compileTemplateExpression } from '../expression';
import { hashSpreadsheetDocument } from '../hash';
import type {
  CompilationResult,
  CompiledTemplate,
  SpreadsheetTemplate,
  TemplateBinding,
  TemplateIRBinding,
  TemplatePrintProfile,
} from '../model';

export const TEMPLATE_COMPILER_VERSION = '1.0.0';

function immutable<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(immutable)) as T;
  if (value !== null && typeof value === 'object') {
    for (const child of Object.values(value as Record<string, unknown>)) immutable(child);
    return Object.freeze(value);
  }
  return value;
}

function diagnostic(code: string, message: string, binding?: TemplateBinding): Diagnostic {
  return immutable({
    code,
    severity: 'error',
    domain: 'template',
    stage: 'compile',
    message,
    ...(binding === undefined ? {} : { location: { bindingId: binding.id } }),
  });
}

function intersects(left: DocumentCellRange, right: DocumentCellRange): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.start.row <= right.end.row &&
    left.end.row >= right.start.row &&
    left.start.column <= right.end.column &&
    left.end.column >= right.start.column
  );
}

function contains(range: DocumentCellRange, row: number, column: number): boolean {
  return (
    row >= range.start.row &&
    row <= range.end.row &&
    column >= range.start.column &&
    column <= range.end.column
  );
}

function compileBinding(binding: TemplateBinding): TemplateIRBinding {
  if (binding.type === 'value') {
    return immutable({
      ...binding,
      expression: compileTemplateExpression(binding.expression),
    });
  }
  if (binding.type === 'repeat-rows') {
    return immutable({ ...binding, source: compileTemplateExpression(binding.source) });
  }
  return immutable({ ...binding, when: compileTemplateExpression(binding.when) });
}

function validateStructuralBindings(
  document: SpreadsheetDocument,
  bindings: readonly TemplateBinding[],
  diagnostics: Diagnostic[],
): void {
  const structural = bindings.filter(
    (binding): binding is Exclude<TemplateBinding, { readonly type: 'value' }> =>
      binding.type !== 'value',
  );
  for (let leftIndex = 0; leftIndex < structural.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < structural.length; rightIndex += 1) {
      const left = structural[leftIndex]!;
      const right = structural[rightIndex]!;
      if (intersects(left.range, right.range)) {
        diagnostics.push(
          diagnostic(
            'OVERLAPPING_REPEAT_REGION',
            `Structural bindings ${left.id} and ${right.id} overlap`,
            right,
          ),
        );
      }
    }
  }
  for (const binding of structural) {
    if (binding.type !== 'repeat-rows') continue;
    const sheet = document.workbook.sheets.find(({ id }) => id === binding.range.sheetId);
    if (sheet === undefined) continue;
    for (const merge of sheet.merges) {
      const startsInside = contains(binding.range, merge.start.row, merge.start.column);
      const endsInside = contains(binding.range, merge.end.row, merge.end.column);
      if (startsInside !== endsInside) {
        diagnostics.push(
          diagnostic(
            'MERGE_CROSSES_REPEAT_BOUNDARY',
            `Merge crosses repeat binding ${binding.id}`,
            binding,
          ),
        );
      }
    }
  }
}

function rangesFromProfile(profile: TemplatePrintProfile): readonly DocumentCellRange[] {
  return profile.targets.flatMap((target) =>
    target.type === 'range' ? [target.range] : target.type === 'ranges' ? target.ranges : [],
  );
}

function validateProfiles(
  document: SpreadsheetDocument,
  profiles: readonly TemplatePrintProfile[],
  diagnostics: Diagnostic[],
): void {
  const sheetIds = new Set(document.workbook.sheets.map(({ id }) => id));
  for (const profile of profiles) {
    if (profile.targets.length === 0) {
      diagnostics.push(diagnostic('INVALID_PRINT_TARGET', `Profile ${profile.id} has no target`));
    }
    for (const target of profile.targets) {
      const targetSheetIds =
        target.type === 'sheet'
          ? [target.sheetId]
          : rangesFromProfile({ ...profile, targets: [target] }).map(({ sheetId }) => sheetId);
      if (targetSheetIds.some((sheetId) => !sheetIds.has(sheetId))) {
        diagnostics.push(
          diagnostic('INVALID_PRINT_TARGET', `Profile ${profile.id} targets an unknown sheet`),
        );
      }
    }
    const { margins } = profile.page;
    if (
      Object.values(margins).some((value) => !Number.isFinite(value) || value < 0) ||
      (profile.page.scale.type === 'fixed' &&
        (!Number.isFinite(profile.page.scale.value) || profile.page.scale.value <= 0))
    ) {
      diagnostics.push(
        diagnostic('INVALID_PRINT_TARGET', `Profile ${profile.id} has invalid page geometry`),
      );
    }
  }
}

/** Compiles explicit binding metadata into immutable, DOM-free template IR. */
export function compileSpreadsheetTemplate(
  document: SpreadsheetDocument,
  template: SpreadsheetTemplate,
): CompilationResult {
  const diagnostics: Diagnostic[] = [];
  const bindings: TemplateIRBinding[] = [];
  for (const binding of template.bindings) {
    try {
      bindings.push(compileBinding(binding));
    } catch (cause) {
      const message =
        cause instanceof TemplateExpressionError ? cause.message : 'Expression compilation failed';
      diagnostics.push(diagnostic('INVALID_EXPRESSION', message, binding));
    }
  }
  validateStructuralBindings(document, template.bindings, diagnostics);
  validateProfiles(document, template.printProfiles, diagnostics);
  const frozenDiagnostics = immutable(diagnostics);
  if (diagnostics.some(({ severity }) => severity === 'error')) {
    return immutable({ diagnostics: frozenDiagnostics, hasErrors: true });
  }
  const compiled: CompiledTemplate = immutable({
    templateId: template.id,
    sourceDocumentHash: hashSpreadsheetDocument(document),
    sourceDocument: document,
    ir: {
      template,
      bindings,
      profiles: template.printProfiles,
    },
    diagnostics: frozenDiagnostics,
    compilerVersion: TEMPLATE_COMPILER_VERSION,
  });
  return immutable({ template: compiled, diagnostics: frozenDiagnostics, hasErrors: false });
}
