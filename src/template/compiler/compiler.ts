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

/** Version of the serialized template IR contract. */
export const TEMPLATE_COMPILER_VERSION = '1.0.0';

function immutableClone<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(immutableClone)) as T;
  if (value !== null && typeof value === 'object') {
    const clone = Object.fromEntries(
      Object.entries(value as Record<string, unknown>).map(([key, child]) => [
        key,
        immutableClone(child),
      ]),
    );
    return Object.freeze(clone) as T;
  }
  return value;
}

function diagnostic(code: string, message: string, binding?: TemplateBinding): Diagnostic {
  return immutableClone({
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
    return immutableClone({
      ...binding,
      expression: compileTemplateExpression(binding.expression),
    });
  }
  if (binding.type === 'repeat-rows') {
    return immutableClone({ ...binding, source: compileTemplateExpression(binding.source) });
  }
  return immutableClone({ ...binding, when: compileTemplateExpression(binding.when) });
}

function normalized(range: DocumentCellRange): boolean {
  return (
    Number.isInteger(range.start.row) &&
    Number.isInteger(range.start.column) &&
    Number.isInteger(range.end.row) &&
    Number.isInteger(range.end.column) &&
    range.start.row >= 0 &&
    range.start.column >= 0 &&
    range.end.row >= range.start.row &&
    range.end.column >= range.start.column
  );
}

function duplicateIds(
  values: readonly { readonly id: string }[],
  code: string,
  label: string,
  diagnostics: Diagnostic[],
): void {
  const seen = new Set<string>();
  for (const value of values) {
    if (seen.has(value.id)) {
      diagnostics.push(diagnostic(code, `Duplicate ${label} ID ${value.id}`));
    }
    seen.add(value.id);
  }
}

function validateBindings(
  document: SpreadsheetDocument,
  bindings: readonly TemplateBinding[],
  diagnostics: Diagnostic[],
): void {
  const sheetIds = new Set(document.workbook.sheets.map(({ id }) => id));
  duplicateIds(bindings, 'DUPLICATE_BINDING_ID', 'binding', diagnostics);
  for (const binding of bindings) {
    if (binding.type === 'value') {
      if (
        !sheetIds.has(binding.target.sheetId) ||
        !Number.isInteger(binding.target.row) ||
        !Number.isInteger(binding.target.column) ||
        binding.target.row < 0 ||
        binding.target.column < 0
      ) {
        diagnostics.push(
          diagnostic(
            'INVALID_BINDING_TARGET',
            `Binding ${binding.id} has an invalid target`,
            binding,
          ),
        );
      }
      continue;
    }
    if (!sheetIds.has(binding.range.sheetId) || !normalized(binding.range)) {
      diagnostics.push(
        diagnostic('INVALID_BINDING_RANGE', `Binding ${binding.id} has an invalid range`, binding),
      );
    }
  }
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
  duplicateIds(profiles, 'DUPLICATE_PRINT_PROFILE_ID', 'print profile', diagnostics);
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
    if (Object.values(margins).some((value) => !Number.isFinite(value) || value < 0)) {
      diagnostics.push(
        diagnostic('INVALID_PAGE_GEOMETRY', `Profile ${profile.id} has invalid margins`),
      );
    }
    if (
      profile.page.paper.type === 'custom' &&
      (!Number.isFinite(profile.page.paper.width) ||
        profile.page.paper.width <= 0 ||
        !Number.isFinite(profile.page.paper.height) ||
        profile.page.paper.height <= 0)
    ) {
      diagnostics.push(
        diagnostic('INVALID_PAGE_GEOMETRY', `Profile ${profile.id} has invalid paper geometry`),
      );
    }
    if (
      (profile.page.scale.type === 'fixed' &&
        (!Number.isFinite(profile.page.scale.value) || profile.page.scale.value <= 0)) ||
      (profile.page.scale.type === 'fit-width' &&
        (!Number.isInteger(profile.page.scale.pages) || profile.page.scale.pages <= 0))
    ) {
      diagnostics.push(
        diagnostic('INVALID_PRINT_SCALE', `Profile ${profile.id} has an invalid scale`),
      );
    }
    for (const range of [profile.repeatRows, profile.repeatColumns]) {
      if (range !== undefined && (!sheetIds.has(range.sheetId) || !normalized(range))) {
        diagnostics.push(
          diagnostic(
            'INVALID_REPEAT_TITLE_RANGE',
            `Profile ${profile.id} has an invalid repeat title range`,
          ),
        );
      }
    }
    for (const pageBreak of profile.manualBreaks) {
      if (
        !sheetIds.has(pageBreak.sheetId) ||
        !Number.isInteger(pageBreak.beforeRow) ||
        pageBreak.beforeRow < 0
      ) {
        diagnostics.push(
          diagnostic('INVALID_PAGE_BREAK', `Profile ${profile.id} has an invalid page break`),
        );
      }
    }
    for (const range of rangesFromProfile(profile)) {
      if (!normalized(range)) {
        diagnostics.push(
          diagnostic('INVALID_PRINT_TARGET', `Profile ${profile.id} has an invalid target range`),
        );
      }
    }
  }
}

function resolveTemplate(
  document: SpreadsheetDocument,
  templateOrId: SpreadsheetTemplate | SpreadsheetTemplate['id'],
  diagnostics: Diagnostic[],
): { readonly document: SpreadsheetDocument; readonly template?: SpreadsheetTemplate } {
  if (typeof templateOrId !== 'string') {
    const existing = document.templates.find(({ id }) => id === templateOrId.id);
    return existing === undefined
      ? {
          document: { ...document, templates: [...document.templates, templateOrId] },
          template: templateOrId,
        }
      : { document, template: existing };
  }
  const template = document.templates.find(({ id }) => id === templateOrId);
  if (template === undefined) {
    diagnostics.push(diagnostic('TEMPLATE_NOT_FOUND', `Template ${templateOrId} does not exist`));
  }
  return { document, ...(template === undefined ? {} : { template }) };
}

/** Compiles one persisted template into immutable, DOM-free template IR. */
export function compileSpreadsheetTemplate(
  document: SpreadsheetDocument,
  templateOrId: SpreadsheetTemplate | SpreadsheetTemplate['id'],
): CompilationResult {
  const diagnostics: Diagnostic[] = [];
  const resolved = resolveTemplate(document, templateOrId, diagnostics);
  const template = resolved.template;
  if (template === undefined) {
    const frozenDiagnostics = immutableClone(diagnostics);
    return immutableClone({ diagnostics: frozenDiagnostics, hasErrors: true });
  }
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
  validateBindings(resolved.document, template.bindings, diagnostics);
  validateStructuralBindings(resolved.document, template.bindings, diagnostics);
  validateProfiles(resolved.document, template.printProfiles, diagnostics);
  const frozenDiagnostics = immutableClone(diagnostics);
  if (diagnostics.some(({ severity }) => severity === 'error')) {
    return immutableClone({ diagnostics: frozenDiagnostics, hasErrors: true });
  }
  const sourceDocument = immutableClone(resolved.document);
  const sourceTemplate = sourceDocument.templates.find(({ id }) => id === template.id)!;
  const compiled: CompiledTemplate = immutableClone({
    templateId: template.id,
    sourceDocumentHash: hashSpreadsheetDocument(sourceDocument),
    sourceDocument,
    ir: {
      template: sourceTemplate,
      bindings,
      profiles: sourceTemplate.printProfiles,
    },
    diagnostics: frozenDiagnostics,
    compilerVersion: TEMPLATE_COMPILER_VERSION,
  });
  return immutableClone({ template: compiled, diagnostics: frozenDiagnostics, hasErrors: false });
}
