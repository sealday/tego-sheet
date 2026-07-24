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

/** Creates a case-insensitive stable-ID name registry. */
export function createFormulaNameRegistry(): FormulaNameRegistry {
  const definitions = new Map<string, FormulaNameDefinition>();
  return {
    register(definition) {
      const key = definition.name.toLocaleLowerCase('en-US');
      if (definitions.has(key)) throw new TypeError(`Duplicate formula name ${definition.name}`);
      const snapshot = Object.freeze({
        ...definition,
        refersTo: Object.freeze(definition.refersTo),
      });
      definitions.set(key, snapshot);
      return () => {
        definitions.delete(key);
      };
    },
    resolve(name, currentSheetId) {
      const definition = definitions.get(name.toLocaleLowerCase('en-US'));
      if (definition?.scope !== 'workbook' && definition?.scope.sheetId !== currentSheetId) {
        return undefined;
      }
      return definition;
    },
  };
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
  const structured = /\b([A-Za-z_][A-Za-z0-9_.]*)\[([^\]]+)\]/gu;
  const masked = source.replace(
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
