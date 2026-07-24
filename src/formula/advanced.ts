import type { DocumentCellAddress, DocumentCellRange } from '../document';
import type { FormulaValue, ScalarFormulaValue } from './ast';

/** Stable named-range definition. */
export interface FormulaNameDefinition {
  /** Stable identifier retained across renames. */
  readonly id: string;
  /** Case-insensitive formula name. */
  readonly name: string;
  /** Workbook-wide or sheet-local visibility. */
  readonly scope:
    | 'workbook'
    | {
        /** Sheet that owns this local name. */
        readonly sheetId: string;
      };
  /** Range referenced by the name. */
  readonly refersTo: DocumentCellRange;
}

/** Isolated named-range registry. */
export interface FormulaNameRegistry {
  /** Registers one definition and returns its disposer. */
  register(definition: FormulaNameDefinition): () => void;
  /** Resolves a visible definition by name. */
  resolve(name: string, currentSheetId: string): FormulaNameDefinition | undefined;
}

/** Stable registration failure for conflicting formula names. */
export class FormulaNameConflictError extends TypeError {
  /** Machine-readable registration failure. */
  readonly code = 'FORMULA_NAME_CONFLICT';

  /** Creates one same-name, same-scope conflict. */
  constructor(name: string) {
    super(`Duplicate formula name ${name} in the same scope`);
    this.name = 'FormulaNameConflictError';
  }
}

/** Creates a case-insensitive stable-ID name registry. */
export function createFormulaNameRegistry(): FormulaNameRegistry {
  const workbookDefinitions = new Map<string, FormulaNameDefinition>();
  const localDefinitions = new Map<string, Map<string, FormulaNameDefinition>>();
  return {
    register(definition) {
      const nameKey = definition.name.toLocaleLowerCase('en-US');
      const definitions =
        definition.scope === 'workbook'
          ? workbookDefinitions
          : (localDefinitions.get(definition.scope.sheetId) ??
            (() => {
              const scoped = new Map<string, FormulaNameDefinition>();
              localDefinitions.set(definition.scope.sheetId, scoped);
              return scoped;
            })());
      if (definitions.has(nameKey)) throw new FormulaNameConflictError(definition.name);
      const snapshot = Object.freeze({
        ...definition,
        scope:
          definition.scope === 'workbook'
            ? definition.scope
            : Object.freeze({ sheetId: definition.scope.sheetId }),
        refersTo: Object.freeze({
          sheetId: definition.refersTo.sheetId,
          start: Object.freeze({ ...definition.refersTo.start }),
          end: Object.freeze({ ...definition.refersTo.end }),
        }),
      });
      definitions.set(nameKey, snapshot);
      return () => {
        if (definitions.get(nameKey) === snapshot) definitions.delete(nameKey);
      };
    },
    resolve(name, currentSheetId) {
      const key = name.toLocaleLowerCase('en-US');
      return localDefinitions.get(currentSheetId)?.get(key) ?? workbookDefinitions.get(key);
    },
  };
}

/** Stable request passed to an injected structured-reference binding provider. */
export interface FormulaTableBindingRequest {
  /** Formula-facing table token. */
  readonly tableName: string;
  /** Formula-facing column token. */
  readonly columnName: string;
  /** Sheet containing the formula. */
  readonly currentSheetId: string;
}

/** Result of resolving one structured table-column reference. */
export type FormulaTableBindingResult =
  | {
      /** The stable table and column were resolved. */
      readonly status: 'resolved';
      /** Stable table identifier retained across renames. */
      readonly tableId: string;
      /** Stable column identifier retained across renames. */
      readonly columnId: string;
      /** Current rectangular cells represented by the column. */
      readonly range: DocumentCellRange;
    }
  | {
      /** The provider could not bind the requested reference. */
      readonly status: 'invalid';
      /** Stable human-readable diagnostic detail. */
      readonly message: string;
    };

/** Injectable bridge used until persistent structured tables land in TBL-01. */
export interface FormulaTableBindingResolver {
  /** Resolves a display reference to stable identifiers and its current range. */
  resolve(request: FormulaTableBindingRequest): FormulaTableBindingResult;
}

/** Stable identifiers available while binding an advanced formula. */
export interface AdvancedFormulaBindingContext {
  /** Sheet used to resolve local names. */
  readonly currentSheetId: string;
  /** Registered workbook and sheet names. */
  readonly names: FormulaNameRegistry;
  /** Structured tables available to the formula. */
  readonly tables: readonly {
    /** Stable table identifier. */
    readonly id: string;
    /** Case-insensitive formula-facing table name. */
    readonly name: string;
    /** Ordered columns available through structured references. */
    readonly columns: readonly {
      /** Stable column identifier. */
      readonly id: string;
      /** Formula-facing column name. */
      readonly name: string;
    }[];
  }[];
}

/** Binds named and structured references to stable identifiers without evaluating them. */
export function bindAdvancedFormula(
  source: string,
  context: AdvancedFormulaBindingContext,
): {
  /** Stable identifiers referenced by the source formula, in source order. */
  readonly references: readonly (
    | {
        /** Named-range reference discriminator. */
        readonly kind: 'name';
        /** Stable named-range identifier. */
        readonly id: string;
      }
    | {
        /** Structured table-column reference discriminator. */
        readonly kind: 'table-column';
        /** Stable table identifier. */
        readonly tableId: string;
        /** Stable column identifier. */
        readonly columnId: string;
      }
  )[];
  /** Recoverable binding failures discovered in the source formula. */
  readonly diagnostics: readonly {
    /** Stable invalid-reference diagnostic code. */
    readonly code: 'FORMULA_REFERENCE_INVALID';
    /** Human-readable diagnostic detail. */
    readonly message: string;
  }[];
} {
  if (!source.startsWith('=')) throw new TypeError('Formula source must start with =');
  const references: {
    readonly index: number;
    readonly reference:
      | { kind: 'name'; id: string }
      | { kind: 'table-column'; tableId: string; columnId: string };
  }[] = [];
  const diagnostics: { code: 'FORMULA_REFERENCE_INVALID'; message: string }[] = [];
  const quotedMask = maskQuotedFormulaText(source);
  const structured = /\b([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]/gu;
  const masked = quotedMask.replace(
    structured,
    (match, tableName: string, columnName: string, offset: number) => {
      const table = context.tables.find(
        ({ name }) => name.toLocaleLowerCase('en-US') === tableName.toLocaleLowerCase('en-US'),
      );
      const column = table?.columns.find(
        ({ name }) => name.toLocaleLowerCase('en-US') === columnName.toLocaleLowerCase('en-US'),
      );
      if (table === undefined || column === undefined) {
        diagnostics.push({
          code: 'FORMULA_REFERENCE_INVALID',
          message: `Unknown structured reference ${tableName}[${columnName}]`,
        });
      } else {
        references.push({
          index: offset,
          reference: {
            kind: 'table-column',
            tableId: table.id,
            columnId: column.id,
          },
        });
      }
      return ' '.repeat(match.length);
    },
  );
  for (const match of masked.matchAll(/\b[A-Za-z_][A-Za-z0-9_.]*\b/gu)) {
    const token = match[0];
    if (['TRUE', 'FALSE'].includes(token.toUpperCase())) continue;
    const name = context.names.resolve(token, context.currentSheetId);
    if (name !== undefined) {
      references.push({
        index: match.index,
        reference: { kind: 'name', id: name.id },
      });
    }
  }
  return {
    references: Object.freeze(
      references.sort((left, right) => left.index - right.index).map(({ reference }) => reference),
    ),
    diagnostics: Object.freeze(diagnostics),
  };
}

function maskQuotedFormulaText(source: string): string {
  const output = source.split('');
  let index = 0;
  while (index < source.length) {
    const quote = source[index];
    if (quote !== '"' && quote !== "'") {
      index += 1;
      continue;
    }
    output[index] = ' ';
    index += 1;
    while (index < source.length) {
      output[index] = ' ';
      if (source[index] !== quote) {
        index += 1;
        continue;
      }
      if (source[index + 1] === quote) {
        output[index + 1] = ' ';
        index += 2;
        continue;
      }
      index += 1;
      break;
    }
  }
  return output.join('');
}

function columnName(column: number): string {
  let value = column + 1;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function addressKey(address: DocumentCellAddress): string {
  return `${address.sheetId}!${columnName(address.column)}${address.row + 1}`;
}

/** Atomic projection plan for one array formula result. */
/** Stable blocked dynamic-array value. */
export interface FormulaSpillError {
  /** Formula error discriminator. */
  readonly type: 'error';
  /** Stable blocked-spill error value. */
  readonly value: '#SPILL!';
}

/** Atomic projection plan for one array formula result. */
export type FormulaSpillPlan =
  | {
      /** Indicates every target cell is available. */
      readonly status: 'ready';
      /** Qualified cell keys mapped to scalar results. */
      readonly cells: ReadonlyMap<string, ScalarFormulaValue>;
    }
  | {
      /** Indicates the spill cannot be projected. */
      readonly status: 'blocked';
      /** Stable spreadsheet error exposed at the anchor. */
      readonly value: FormulaSpillError;
      /** First occupied or invalid target cell. */
      readonly blocker: DocumentCellAddress;
    };

/** Plans an array spill atomically before any cell projection is exposed. */
export function planFormulaSpill(input: {
  readonly anchor: DocumentCellAddress;
  readonly value: Extract<FormulaValue, { readonly type: 'array' }>;
  readonly occupied: ReadonlySet<string>;
  readonly limits: { readonly maxCells: number };
}): FormulaSpillPlan {
  const width = input.value.rows[0]?.length ?? 0;
  const cellCount = input.value.rows.length * width;
  if (cellCount > input.limits.maxCells || input.value.rows.some((row) => row.length !== width)) {
    return {
      status: 'blocked',
      value: { type: 'error', value: '#SPILL!' },
      blocker: input.anchor,
    };
  }
  const cells = new Map<string, ScalarFormulaValue>();
  for (const [rowOffset, row] of input.value.rows.entries()) {
    for (const [columnOffset, value] of row.entries()) {
      const address = {
        sheetId: input.anchor.sheetId,
        row: input.anchor.row + rowOffset,
        column: input.anchor.column + columnOffset,
      };
      if (input.occupied.has(addressKey(address))) {
        return {
          status: 'blocked',
          value: { type: 'error', value: '#SPILL!' },
          blocker: address,
        };
      }
      cells.set(addressKey(address), value);
    }
  }
  return { status: 'ready', cells };
}
