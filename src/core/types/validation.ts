import type { CellAddress, SheetId } from './coordinates';

/** Supported semantic value categories for cell validation. */
export type ValidationType = 'date' | 'number' | 'list' | 'phone' | 'email';

/** Supported validation comparison operators. */
export type ValidationOperator = 'be' | 'nbe' | 'eq' | 'neq' | 'lt' | 'lte' | 'gt' | 'gte';

/** Rule used to validate cell text. */
export interface ValidationRule {
  /** Validation scope; currently validation is applied per cell. */
  readonly mode: 'cell';
  /** Semantic value category checked by the rule. */
  readonly type: ValidationType;
  /** Whether blank cell text is invalid. */
  readonly required: boolean;
  /** Optional comparison applied after type validation. */
  readonly operator?: ValidationOperator;
  /** Comparison operand or inclusive pair of operands for range operators. */
  readonly value?: string | readonly [string, string];
}

/** One cell that failed a validation rule. */
export interface ValidationIssue {
  /** Worksheet containing the invalid cell. */
  readonly sheet: SheetId;
  /** Zero-based position of the worksheet in the workbook. */
  readonly sheetIndex: number;
  /** Zero-based address of the invalid cell. */
  readonly address: CellAddress;
  /** Rule that the cell failed. */
  readonly rule: ValidationRule;
  /** Human-readable reason the value is invalid. */
  readonly message: string;
}

/** Aggregate result returned by imperative workbook validation. */
export interface ValidationResult {
  /** Whether every configured validation rule passed. */
  readonly valid: boolean;
  /** All validation failures; empty when `valid` is `true`. */
  readonly issues: readonly ValidationIssue[];
}
