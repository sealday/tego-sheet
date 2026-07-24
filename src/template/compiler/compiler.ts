import type { Diagnostic, DocumentCellRange, SpreadsheetDocument } from '../../document';
import { TemplateExpressionError, compileTemplateExpression } from '../expression';
import { hashSpreadsheetDocument } from '../hash';
import type {
  CompilationResult,
  CompiledTemplate,
  AdvancedCompileOptions,
  SpreadsheetTemplate,
  TemplateBinding,
  TemplateIRBinding,
  TemplatePrintProfile,
  TemplateRegionNode,
  ValueBinding,
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

function containsRange(outer: DocumentCellRange, inner: DocumentCellRange): boolean {
  return (
    outer.sheetId === inner.sheetId &&
    contains(outer, inner.start.row, inner.start.column) &&
    contains(outer, inner.end.row, inner.end.column)
  );
}

function sameRange(left: DocumentCellRange, right: DocumentCellRange): boolean {
  return (
    left.sheetId === right.sheetId &&
    left.start.row === right.start.row &&
    left.start.column === right.start.column &&
    left.end.row === right.end.row &&
    left.end.column === right.end.column
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
  if (binding.type === 'conditional-range') {
    return immutableClone({ ...binding, when: compileTemplateExpression(binding.when) });
  }
  if (binding.type === 'repeat-sheet') {
    return immutableClone({
      ...binding,
      source: compileTemplateExpression(binding.source),
      name: compileTemplateExpression(binding.name),
    }) as TemplateIRBinding;
  }
  return immutableClone({
    ...binding,
    source: compileTemplateExpression(binding.source),
  }) as TemplateIRBinding;
}

function objectAnchorRange(
  object: SpreadsheetDocument['workbook']['sheets'][number]['objects'][number],
): DocumentCellRange | undefined {
  if (object.anchor.type === 'absolute') return undefined;
  return object.anchor.type === 'one-cell'
    ? {
        sheetId: object.anchor.cell.sheetId,
        start: { row: object.anchor.cell.row, column: object.anchor.cell.column },
        end: { row: object.anchor.cell.row, column: object.anchor.cell.column },
      }
    : {
        sheetId: object.anchor.from.sheetId,
        start: { row: object.anchor.from.row, column: object.anchor.from.column },
        end: { row: object.anchor.to.row, column: object.anchor.to.column },
      };
}

type ObjectRepeatBinding = Extract<
  TemplateBinding,
  { readonly type: 'repeat-rows' | 'repeat-range' | 'repeat-page' | 'repeat-sheet' }
>;

function isObjectRepeatBinding(binding: TemplateBinding): binding is ObjectRepeatBinding {
  return (
    binding.type === 'repeat-rows' ||
    binding.type === 'repeat-range' ||
    binding.type === 'repeat-page' ||
    binding.type === 'repeat-sheet'
  );
}

function detectRepeatedObjects(
  document: SpreadsheetDocument,
  template: SpreadsheetTemplate,
  diagnostics: Diagnostic[],
): SpreadsheetTemplate {
  const repeats = template.bindings.filter(isObjectRepeatBinding);
  const persistent = document.workbook.sheets.flatMap((sheet) =>
    (sheet.objects ?? []).flatMap((object) => {
      const anchor = objectAnchorRange(object);
      if (anchor === undefined) return [];
      return [
        {
          key: `${String(sheet.id)}:${String(object.id)}`,
          reference: {
            id: String(object.id),
            anchor,
            anchorMode:
              object.templateRepeat === 'shared' ? ('absolute' as const) : ('range' as const),
            ...(object.kind === 'image' ? { resourceId: String(object.resourceId) } : {}),
          },
        },
      ];
    }),
  );
  const anchorByKey = new Map(persistent.map(({ key, reference }) => [key, reference.anchor]));
  for (const binding of repeats) {
    for (const reference of binding.objects ?? []) {
      const key = `${String(binding.range.sheetId)}:${reference.id}`;
      if (!anchorByKey.has(key)) anchorByKey.set(key, reference.anchor);
    }
  }
  const ownerByKey = new Map<string, string>();
  for (const [key, anchor] of anchorByKey) {
    const candidates = repeats
      .filter(
        (binding) => binding.range.sheetId === anchor.sheetId && intersects(binding.range, anchor),
      )
      .sort((left, right) => {
        const leftArea =
          (left.range.end.row - left.range.start.row + 1) *
          (left.range.end.column - left.range.start.column + 1);
        const rightArea =
          (right.range.end.row - right.range.start.row + 1) *
          (right.range.end.column - right.range.start.column + 1);
        return leftArea - rightArea || String(left.id).localeCompare(String(right.id));
      });
    const owner = candidates[0];
    if (owner === undefined) continue;
    const ambiguous = candidates
      .slice(1)
      .find((candidate) => !containsRange(candidate.range, owner.range));
    if (
      ambiguous !== undefined ||
      candidates.some(
        (candidate) => sameRange(candidate.range, owner.range) && candidate.id !== owner.id,
      )
    ) {
      diagnostics.push(
        diagnostic(
          'INVALID_OBJECT_ANCHOR',
          `Object ${key.slice(key.indexOf(':') + 1)} intersects repeats without a unique deepest owner`,
          owner,
        ),
      );
    }
    ownerByKey.set(key, String(owner.id));
  }
  const persistentByOwner = new Map<string, (typeof persistent)[number]['reference'][]>();
  for (const { key, reference } of persistent) {
    const owner = ownerByKey.get(key);
    if (owner === undefined) continue;
    const values = persistentByOwner.get(owner) ?? [];
    values.push(reference);
    persistentByOwner.set(owner, values);
  }
  return {
    ...template,
    bindings: template.bindings.map((binding) => {
      if (!isObjectRepeatBinding(binding)) return binding;
      const ownedDeclared = (binding.objects ?? []).filter(
        ({ id }) => ownerByKey.get(`${String(binding.range.sheetId)}:${id}`) === String(binding.id),
      );
      const declared = new Set(ownedDeclared.map(({ id }) => id));
      const owned = [
        ...ownedDeclared,
        ...(persistentByOwner.get(String(binding.id)) ?? []).filter(({ id }) => !declared.has(id)),
      ];
      const { objects: _objects, ...withoutObjects } = binding;
      return owned.length === 0 ? withoutObjects : { ...withoutObjects, objects: owned };
    }),
  };
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
  advanced: boolean,
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
        if (advanced && sameRange(left.range, right.range)) {
          diagnostics.push(
            diagnostic(
              'INVALID_NESTING',
              `Structural bindings ${left.id} and ${right.id} occupy the same region`,
              right,
            ),
          );
          continue;
        }
        const leftContainsRight = containsRange(left.range, right.range);
        const rightContainsLeft = containsRange(right.range, left.range);
        if (advanced && (leftContainsRight || rightContainsLeft)) continue;
        diagnostics.push(
          diagnostic(
            advanced ? 'PARTIALLY_OVERLAPPING_REGION' : 'OVERLAPPING_REPEAT_REGION',
            `Structural bindings ${left.id} and ${right.id} overlap`,
            right,
          ),
        );
      }
    }
  }
  for (const binding of structural) {
    if ('objects' in binding && binding.objects !== undefined && binding.objects.length > 0) {
      if (binding.objectPolicy === undefined) {
        diagnostics.push(
          diagnostic(
            'OBJECT_REPEAT_POLICY_REQUIRED',
            `Binding ${binding.id} intersects floating objects but has no copy policy`,
            binding,
          ),
        );
      } else if (binding.objectPolicy === 'forbidden') {
        diagnostics.push(
          diagnostic(
            'OBJECT_REPEAT_POLICY_REQUIRED',
            `Binding ${binding.id} requires per-item or shared object behavior`,
            binding,
          ),
        );
        diagnostics.push(
          diagnostic(
            'OBJECT_REPEAT_FORBIDDEN',
            `Binding ${binding.id} forbids repeating intersecting objects`,
            binding,
          ),
        );
      }
      const objectIds = new Set<string>();
      for (const object of binding.objects) {
        if (objectIds.has(object.id)) {
          diagnostics.push(
            diagnostic(
              'INVALID_OBJECT_ANCHOR',
              `Binding ${binding.id} repeats duplicate object ${object.id}`,
              binding,
            ),
          );
        }
        objectIds.add(object.id);
        const persistentObject = document.workbook.sheets
          .find(({ id }) => id === binding.range.sheetId)
          ?.objects.find(({ id }) => id === object.id);
        if (
          persistentObject !== undefined &&
          binding.objectPolicy !== undefined &&
          binding.objectPolicy !== 'forbidden' &&
          persistentObject.templateRepeat !== binding.objectPolicy
        ) {
          diagnostics.push(
            diagnostic(
              'OBJECT_REPEAT_POLICY_REQUIRED',
              `Object ${object.id} policy ${persistentObject.templateRepeat} does not match ${binding.objectPolicy}`,
              binding,
            ),
          );
        }
        if (
          persistentObject?.kind === 'image' &&
          object.resourceId !== undefined &&
          object.resourceId !== persistentObject.resourceId
        ) {
          diagnostics.push(
            diagnostic(
              'INVALID_OBJECT_ANCHOR',
              `Object ${object.id} resource does not match the persistent object`,
              binding,
            ),
          );
        }
        if (
          object.anchor === undefined ||
          !normalized(object.anchor) ||
          object.anchor.sheetId !== binding.range.sheetId ||
          !intersects(binding.range, object.anchor) ||
          (object.anchorMode !== 'range' && object.anchorMode !== 'absolute') ||
          (object.resourceId !== undefined && typeof object.resourceId !== 'string')
        ) {
          diagnostics.push(
            diagnostic(
              'INVALID_OBJECT_ANCHOR',
              `Object ${object.id} has an invalid repeat anchor`,
              binding,
            ),
          );
        } else if (
          (binding.objectPolicy === 'per-item' && object.anchorMode !== 'range') ||
          (binding.objectPolicy === 'shared' && object.anchorMode !== 'absolute')
        ) {
          diagnostics.push(
            diagnostic(
              'INVALID_OBJECT_ANCHOR',
              `Object ${object.id} anchor mode does not match ${binding.objectPolicy}`,
              binding,
            ),
          );
        }
      }
    }
    if (
      binding.type !== 'repeat-rows' &&
      binding.type !== 'repeat-range' &&
      binding.type !== 'repeat-page'
    )
      continue;
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

function buildRegionTree(
  bindings: readonly TemplateBinding[],
  maxDepth: number,
  diagnostics: Diagnostic[],
): readonly TemplateRegionNode[] {
  const structural = bindings.filter(
    (binding): binding is Exclude<TemplateBinding, ValueBinding> => binding.type !== 'value',
  );
  const nodes = new Map(
    structural.map((binding) => [
      binding.id,
      {
        bindingId: binding.id,
        range: binding.range,
        depth: 0,
        children: [] as TemplateRegionNode[],
      },
    ]),
  );
  const roots: TemplateRegionNode[] = [];
  for (const binding of structural) {
    const parents = structural
      .filter(
        (candidate) =>
          candidate.id !== binding.id &&
          !sameRange(candidate.range, binding.range) &&
          containsRange(candidate.range, binding.range),
      )
      .sort((left, right) => {
        const leftArea =
          (left.range.end.row - left.range.start.row + 1) *
          (left.range.end.column - left.range.start.column + 1);
        const rightArea =
          (right.range.end.row - right.range.start.row + 1) *
          (right.range.end.column - right.range.start.column + 1);
        return leftArea - rightArea || String(left.id).localeCompare(String(right.id));
      });
    const node = nodes.get(binding.id)!;
    const parent = parents[0];
    if (parent === undefined) roots.push(node);
    else nodes.get(parent.id)!.children.push(node);
  }
  const finalize = (node: TemplateRegionNode, depth: number): TemplateRegionNode => {
    if (depth > maxDepth) {
      diagnostics.push(
        diagnostic(
          'INVALID_NESTING',
          `Binding ${node.bindingId} exceeds nesting depth ${maxDepth}`,
        ),
      );
    }
    return immutableClone({
      ...node,
      depth,
      children: [...node.children]
        .sort(
          (left, right) =>
            left.range.start.row - right.range.start.row ||
            left.range.start.column - right.range.start.column ||
            String(left.bindingId).localeCompare(String(right.bindingId)),
        )
        .map((child) => finalize(child, depth + 1)),
    });
  };
  return roots
    .sort(
      (left, right) =>
        left.range.start.row - right.range.start.row ||
        left.range.start.column - right.range.start.column ||
        String(left.bindingId).localeCompare(String(right.bindingId)),
    )
    .map((root) => finalize(root, 1));
}

function subtemplateDiagnostics(
  root: SpreadsheetTemplate,
  subtemplates: ReadonlyMap<SpreadsheetTemplate['id'], SpreadsheetTemplate>,
  diagnostics: Diagnostic[],
): void {
  const visiting: string[] = [];
  const visited = new Set<string>();
  const visit = (template: SpreadsheetTemplate): void => {
    const cycleAt = visiting.indexOf(template.id);
    if (cycleAt >= 0) {
      const chain = [...visiting.slice(cycleAt), template.id].join(' -> ');
      diagnostics.push(diagnostic('SUBTEMPLATE_CYCLE', `Subtemplate cycle: ${chain}`));
      return;
    }
    if (visited.has(template.id)) return;
    visiting.push(template.id);
    for (const binding of template.bindings) {
      if (binding.type !== 'subtemplate') continue;
      const child = subtemplates.get(binding.templateId);
      if (child === undefined) {
        diagnostics.push(
          diagnostic(
            'SUBTEMPLATE_NOT_FOUND',
            `Subtemplate ${binding.templateId} is not registered`,
            binding,
          ),
        );
      } else {
        visit(child);
      }
    }
    visiting.pop();
    visited.add(template.id);
  };
  visit(root);
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function hasPointShape(value: unknown): boolean {
  return isRecord(value) && typeof value.row === 'number' && typeof value.column === 'number';
}

function hasRangeShape(value: unknown): boolean {
  return (
    isRecord(value) &&
    typeof value.sheetId === 'string' &&
    hasPointShape(value.start) &&
    hasPointShape(value.end)
  );
}

function hasPageShape(value: unknown): boolean {
  if (
    !isRecord(value) ||
    !isRecord(value.paper) ||
    !isRecord(value.margins) ||
    !isRecord(value.scale) ||
    (value.orientation !== 'portrait' && value.orientation !== 'landscape')
  ) {
    return false;
  }
  const margins = value.margins;
  const paper = value.paper;
  const scale = value.scale;
  const validPaper =
    paper.type === 'A4' ||
    paper.type === 'A5' ||
    paper.type === 'Letter' ||
    (paper.type === 'custom' &&
      typeof paper.width === 'number' &&
      typeof paper.height === 'number');
  const validScale =
    (scale.type === 'fixed' && typeof scale.value === 'number') ||
    (scale.type === 'fit-width' && typeof scale.pages === 'number') ||
    scale.type === 'fit-page';
  return (
    validPaper &&
    validScale &&
    ['top', 'right', 'bottom', 'left'].every((side) => typeof margins[side] === 'number')
  );
}

function hasBandShape(value: unknown): boolean {
  return (
    isRecord(value) &&
    ['left', 'center', 'right'].every(
      (slot) => value[slot] === undefined || typeof value[slot] === 'string',
    )
  );
}

function hasRuntimeTemplateShape(value: unknown): value is SpreadsheetTemplate {
  if (
    !isRecord(value) ||
    typeof value.id !== 'string' ||
    typeof value.name !== 'string' ||
    !Array.isArray(value.bindings) ||
    !Array.isArray(value.printProfiles)
  ) {
    return false;
  }
  const bindingsAreValid = value.bindings.every((binding) => {
    if (!isRecord(binding) || typeof binding.id !== 'string' || typeof binding.type !== 'string') {
      return false;
    }
    if (binding.type === 'value') {
      return (
        isRecord(binding.target) &&
        typeof binding.target.sheetId === 'string' &&
        hasPointShape(binding.target) &&
        typeof binding.expression === 'string' &&
        (binding.formatter === undefined || typeof binding.formatter === 'string')
      );
    }
    if (binding.type === 'repeat-rows') {
      return (
        hasRangeShape(binding.range) &&
        typeof binding.source === 'string' &&
        (binding.empty === 'remove' || binding.empty === 'keep-template-row') &&
        (binding.pageBreak === 'auto' || binding.pageBreak === 'before-each-item')
      );
    }
    if (
      binding.type === 'repeat-columns' ||
      binding.type === 'repeat-range' ||
      binding.type === 'repeat-page'
    ) {
      return (
        hasRangeShape(binding.range) &&
        typeof binding.source === 'string' &&
        (binding.empty === 'remove' || binding.empty === 'keep-template-row') &&
        (binding.type !== 'repeat-range' ||
          binding.axis === 'vertical' ||
          binding.axis === 'horizontal' ||
          binding.axis === 'both')
      );
    }
    if (binding.type === 'repeat-sheet') {
      return (
        hasRangeShape(binding.range) &&
        typeof binding.source === 'string' &&
        typeof binding.name === 'string'
      );
    }
    if (binding.type === 'subtemplate') {
      return (
        hasRangeShape(binding.range) &&
        typeof binding.templateId === 'string' &&
        typeof binding.source === 'string'
      );
    }
    return (
      binding.type === 'conditional-range' &&
      hasRangeShape(binding.range) &&
      typeof binding.when === 'string'
    );
  });
  return (
    bindingsAreValid &&
    value.printProfiles.every(
      (profile) =>
        isRecord(profile) &&
        typeof profile.id === 'string' &&
        typeof profile.name === 'string' &&
        Array.isArray(profile.targets) &&
        profile.targets.every(
          (target) =>
            isRecord(target) &&
            ((target.type === 'sheet' && typeof target.sheetId === 'string') ||
              (target.type === 'range' && hasRangeShape(target.range)) ||
              (target.type === 'ranges' &&
                Array.isArray(target.ranges) &&
                target.ranges.every(hasRangeShape))),
        ) &&
        hasPageShape(profile.page) &&
        Array.isArray(profile.manualBreaks) &&
        profile.manualBreaks.every(
          (pageBreak) =>
            isRecord(pageBreak) &&
            typeof pageBreak.sheetId === 'string' &&
            typeof pageBreak.beforeRow === 'number',
        ) &&
        (profile.repeatRows === undefined || hasRangeShape(profile.repeatRows)) &&
        (profile.repeatColumns === undefined || hasRangeShape(profile.repeatColumns)) &&
        (profile.header === undefined || hasBandShape(profile.header)) &&
        (profile.footer === undefined || hasBandShape(profile.footer)) &&
        typeof profile.showGridlines === 'boolean' &&
        typeof profile.showHeadings === 'boolean',
    )
  );
}

function hasUnsafeCompilationGraph(roots: readonly unknown[]): boolean {
  interface Frame {
    readonly object: object;
    readonly values: readonly unknown[];
    readonly depth: number;
    index: number;
  }

  const active = new WeakSet<object>();
  const stack: Frame[] = [];
  let entries = 0;
  let stringUnits = 0;
  const enter = (value: unknown, depth: number): boolean => {
    if (typeof value === 'string') {
      stringUnits += value.length;
      return stringUnits <= 32_000_000;
    }
    if (value === null || typeof value !== 'object') return true;
    if (active.has(value) || depth > 512) return false;
    let descriptors: readonly PropertyDescriptor[];
    try {
      descriptors = Object.values(Object.getOwnPropertyDescriptors(value));
    } catch {
      return false;
    }
    if (descriptors.some((descriptor) => !('value' in descriptor))) return false;
    entries += descriptors.length;
    if (entries > 1_000_000) return false;
    active.add(value);
    stack.push({
      object: value,
      values: descriptors.map((descriptor) =>
        'value' in descriptor ? descriptor.value : undefined,
      ),
      depth,
      index: 0,
    });
    return true;
  };
  for (const root of roots) {
    if (!enter(root, 0)) return true;
    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!;
      if (frame.index >= frame.values.length) {
        active.delete(frame.object);
        stack.pop();
        continue;
      }
      if (!enter(frame.values[frame.index++], frame.depth + 1)) return true;
    }
  }
  return false;
}

function exceedsCompilationBudget(document: SpreadsheetDocument, template: SpreadsheetTemplate) {
  const sourceCollections =
    document.workbook.sheets.reduce(
      (sum, sheet) =>
        sum + sheet.cells.length + sheet.rows.length + sheet.columns.length + sheet.merges.length,
      0,
    ) +
    document.workbook.styles.length +
    document.workbook.validations.length +
    document.resources.items.length;
  const templateBindings = document.templates.reduce(
    (sum, sourceTemplate) => sum + sourceTemplate.bindings.length,
    0,
  );
  const templateProfiles = document.templates.reduce(
    (sum, sourceTemplate) => sum + sourceTemplate.printProfiles.length,
    0,
  );
  const pending: unknown[] = [document];
  const seen = new WeakSet<object>();
  let graphEntries = 0;
  let stringUnits = 0;
  while (pending.length > 0) {
    const value = pending.pop();
    if (typeof value === 'string') {
      stringUnits += value.length;
      if (stringUnits > 32_000_000) return true;
      continue;
    }
    if (value === null || typeof value !== 'object' || seen.has(value)) continue;
    seen.add(value);
    const descriptors = Object.values(Object.getOwnPropertyDescriptors(value));
    graphEntries += descriptors.length;
    if (graphEntries > 1_000_000) return true;
    for (const descriptor of descriptors) {
      if (!('value' in descriptor)) return true;
      pending.push(descriptor.value);
    }
  }
  return (
    sourceCollections > 1_000_000 ||
    templateBindings > 10_000 ||
    templateProfiles > 1_000 ||
    template.bindings.length > 10_000 ||
    template.printProfiles.length > 1_000
  );
}

/** Compiles one persisted template into immutable, DOM-free template IR. */
export function compileSpreadsheetTemplate(
  document: SpreadsheetDocument,
  templateOrId: SpreadsheetTemplate | SpreadsheetTemplate['id'],
  options?: AdvancedCompileOptions,
): CompilationResult {
  const diagnostics: Diagnostic[] = [];
  if (
    hasUnsafeCompilationGraph([
      document,
      templateOrId,
      options?.limits,
      ...(options === undefined ? [] : [...options.subtemplates.values()]),
    ])
  ) {
    const frozenDiagnostics = immutableClone([
      diagnostic('COMPILATION_RESOURCE_LIMIT', 'Template source exceeds compilation limits'),
    ]);
    return immutableClone({ diagnostics: frozenDiagnostics, hasErrors: true });
  }
  const resolved = resolveTemplate(document, templateOrId, diagnostics);
  const unresolvedTemplate = resolved.template;
  if (unresolvedTemplate === undefined) {
    const frozenDiagnostics = immutableClone(diagnostics);
    return immutableClone({ diagnostics: frozenDiagnostics, hasErrors: true });
  }
  if (!hasRuntimeTemplateShape(unresolvedTemplate)) {
    const frozenDiagnostics = immutableClone([
      ...diagnostics,
      diagnostic('INVALID_TEMPLATE_STRUCTURE', 'Template structure is malformed'),
    ]);
    return immutableClone({ diagnostics: frozenDiagnostics, hasErrors: true });
  }
  const template = detectRepeatedObjects(resolved.document, unresolvedTemplate, diagnostics);
  if (exceedsCompilationBudget(resolved.document, template)) {
    const frozenDiagnostics = immutableClone([
      ...diagnostics,
      diagnostic('COMPILATION_RESOURCE_LIMIT', 'Template source exceeds compilation limits'),
    ]);
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
  if (options !== undefined) {
    for (const child of options.subtemplates.values()) {
      if (!hasRuntimeTemplateShape(child)) {
        diagnostics.push(
          diagnostic(
            'INVALID_TEMPLATE_STRUCTURE',
            `Subtemplate ${String((child as unknown as { readonly id?: unknown }).id)} is malformed`,
          ),
        );
        continue;
      }
      for (const childBinding of child.bindings) {
        try {
          compileBinding(childBinding);
        } catch (cause) {
          diagnostics.push(
            diagnostic(
              'INVALID_EXPRESSION',
              cause instanceof TemplateExpressionError
                ? `Subtemplate ${String(child.id)}: ${cause.message}`
                : `Subtemplate ${String(child.id)} expression compilation failed`,
              childBinding,
            ),
          );
        }
      }
    }
  }
  validateBindings(resolved.document, template.bindings, diagnostics);
  validateStructuralBindings(
    resolved.document,
    template.bindings,
    diagnostics,
    options !== undefined,
  );
  const regionTree =
    options === undefined
      ? undefined
      : buildRegionTree(template.bindings, options.limits.maxNestingDepth ?? 8, diagnostics);
  if (options !== undefined) subtemplateDiagnostics(template, options.subtemplates, diagnostics);
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
      ...(regionTree === undefined ? {} : { regionTree }),
      ...(options === undefined
        ? {}
        : {
            subtemplates: [...options.subtemplates.values()].sort((left, right) =>
              String(left.id).localeCompare(String(right.id)),
            ),
          }),
    },
    diagnostics: frozenDiagnostics,
    compilerVersion: TEMPLATE_COMPILER_VERSION,
  });
  return immutableClone({ template: compiled, diagnostics: frozenDiagnostics, hasErrors: false });
}
