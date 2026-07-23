import type { Cell, Diagnostic } from '../../document';
import { parseFormula, renderFormula, translateFormula } from '../../formula';
import { evaluateTemplateExpression, type TemplateFormatterRegistry } from '../expression';
import type { TemplateIRBinding } from '../model';

export interface ExpansionScope {
  readonly root: unknown;
  readonly item?: unknown;
  readonly parent?: unknown;
  readonly index?: number;
  readonly first?: boolean;
  readonly last?: boolean;
}

export function freeze<T>(value: T): T {
  if (Array.isArray(value)) return Object.freeze(value.map(freeze)) as T;
  if (value !== null && typeof value === 'object' && !(value instanceof Map)) {
    for (const child of Object.values(value as Record<string, unknown>)) freeze(child);
    return Object.freeze(value);
  }
  return value;
}

export function expansionError(code: string, message: string, bindingId?: string): Diagnostic {
  return freeze({
    code,
    severity: 'error',
    domain: 'template',
    stage: 'expand',
    message,
    ...(bindingId === undefined ? {} : { location: { bindingId: bindingId as never } }),
  });
}

export function translatedCell(cell: Cell, rowDelta: number, columnDelta: number): Cell {
  if (cell.input.type !== 'formula' || (rowDelta === 0 && columnDelta === 0)) return cell;
  try {
    return {
      ...cell,
      input: {
        type: 'formula',
        source: renderFormula(
          translateFormula(parseFormula(cell.input.source), { rowDelta, columnDelta }),
        ),
      },
    };
  } catch {
    return cell;
  }
}

export function collection(
  binding: Extract<
    TemplateIRBinding,
    {
      readonly type:
        | 'repeat-columns'
        | 'repeat-range'
        | 'repeat-page'
        | 'repeat-sheet'
        | 'subtemplate';
    }
  >,
  scope: ExpansionScope,
  formatters: TemplateFormatterRegistry,
): readonly unknown[] {
  const value = evaluateTemplateExpression(binding.source, scope, formatters);
  return Array.isArray(value) ? value : [];
}

export function safeSheetName(value: unknown, fallback: string): string {
  const sanitized = String(value ?? fallback)
    .replace(/[:\\/?*[\]]/gu, '_')
    .trim()
    .slice(0, 31);
  return sanitized || fallback;
}

export function valueInput(value: unknown): Cell['input'] {
  if (value === undefined || value === null) return { type: 'blank' };
  if (typeof value === 'string') return { type: 'string', value };
  if (typeof value === 'number' && Number.isFinite(value)) return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  return { type: 'string', value: JSON.stringify(value) };
}
