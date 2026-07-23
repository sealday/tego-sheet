import type { DocumentCellAddress } from '../document';
import type { FormulaValue } from '../formula';

/** Bounded list source used by a validation rule. */
export type ValidationListSource =
  | { readonly type: 'static'; readonly values: readonly string[] }
  | { readonly type: 'resolver'; readonly id: string };

/** Typed Phase 3A validation rule. */
export type ValidationRule =
  | {
      readonly id: string;
      readonly type: 'number';
      readonly predicate: {
        readonly operator: 'between';
        readonly minimum: number;
        readonly maximum: number;
      };
      readonly behavior: 'reject' | 'warn';
      readonly allowBlank: boolean;
    }
  | {
      readonly id: string;
      readonly type: 'list';
      readonly predicate: { readonly source: ValidationListSource };
      readonly behavior: 'reject' | 'warn';
      readonly allowBlank: boolean;
    };

/** One immutable edit candidate. */
export interface ValidationRequest {
  readonly address: DocumentCellAddress;
  readonly value: FormulaValue;
  readonly rule: ValidationRule;
  readonly signal?: AbortSignal;
}

/** Stable validation result returned before any controller mutation. */
export type ValidationResult =
  | { readonly status: 'accepted'; readonly diagnostics: readonly [] }
  | {
      readonly status: 'rejected' | 'warning';
      readonly code: 'VALIDATION_REJECTED';
      readonly diagnostics: readonly [
        { readonly code: 'VALIDATION_REJECTED'; readonly ruleId: string },
      ];
    }
  | {
      readonly status: 'error';
      readonly code: 'VALIDATION_SOURCE_ERROR' | 'VALIDATION_SOURCE_TOO_LARGE';
      readonly diagnostics: readonly [{ readonly code: string; readonly ruleId: string }];
    };
