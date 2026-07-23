import type { DocumentCellAddress } from '../document';
import type { FormulaValue } from '../formula';

/** Bounded list source used by a validation rule. */
export type ValidationListSource =
  | {
      /** Uses values embedded in the rule. */
      readonly type: 'static';
      /** Allowed string values. */
      readonly values: readonly string[];
    }
  | {
      /** Uses a host-registered resolver. */
      readonly type: 'resolver';
      /** Resolver registry identifier. */
      readonly id: string;
    };

/** Comparison operators shared by scalar validation rules. */
export type ValidationComparisonOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

/** Comparison against one scalar boundary. */
export interface ValidationScalarComparison<Value> {
  /** Scalar comparison operator. */
  readonly operator: Exclude<ValidationComparisonOperator, 'between' | 'notBetween'>;
  /** Comparison boundary. */
  readonly value: Value;
}

/** Comparison against an inclusive pair of scalar boundaries. */
export interface ValidationRangeComparison<Value> {
  /** Inclusive range comparison operator. */
  readonly operator: 'between' | 'notBetween';
  /** Inclusive minimum. */
  readonly minimum: Value;
  /** Inclusive maximum. */
  readonly maximum: Value;
}

/** Scalar or inclusive-range comparison definition. */
export type ValidationComparison<Value> =
  | ValidationScalarComparison<Value>
  | ValidationRangeComparison<Value>;

/** Enforcement and blank handling shared by validation rules. */
export interface ValidationRuleBase {
  /** Stable rule identifier. */
  readonly id: string;
  /** Whether invalid edits are rejected or warned. */
  readonly behavior: 'reject' | 'warn';
  /** Whether a blank value bypasses validation. */
  readonly allowBlank: boolean;
}

/** Typed Phase 3A validation rule. */
export type ValidationRule =
  | (ValidationRuleBase & {
      /** Validates integral numeric values. */
      readonly type: 'whole';
      /** Numeric comparison definition. */
      readonly predicate: ValidationComparison<number>;
    })
  | (ValidationRuleBase & {
      /** Validates finite numeric values. */
      readonly type: 'decimal' | 'number';
      /** Numeric comparison definition. */
      readonly predicate: ValidationComparison<number>;
    })
  | (ValidationRuleBase & {
      /** Validates membership in a list. */
      readonly type: 'list';
      /** List source definition. */
      readonly predicate: {
        /** Static or host-resolved allowed values. */
        readonly source: ValidationListSource;
      };
    })
  | (ValidationRuleBase & {
      /** Validates an ISO calendar date or Excel date serial. */
      readonly type: 'date';
      /** Date comparison definition. */
      readonly predicate: ValidationComparison<string>;
    })
  | (ValidationRuleBase & {
      /** Validates an ISO local time or Excel day fraction. */
      readonly type: 'time';
      /** Time comparison definition. */
      readonly predicate: ValidationComparison<string>;
    })
  | (ValidationRuleBase & {
      /** Validates the displayed scalar text length. */
      readonly type: 'text-length';
      /** Character-count comparison definition. */
      readonly predicate: ValidationComparison<number>;
    })
  | (ValidationRuleBase & {
      /** Validates with a host-injected restricted formula evaluator. */
      readonly type: 'custom-formula';
      /** Formula source evaluated against the immutable request. */
      readonly predicate: {
        /** Formula source beginning with `=`. */
        readonly formula: string;
      };
    });

/** One immutable edit candidate. */
export interface ValidationRequest {
  /** Cell receiving the proposed value. */
  readonly address: DocumentCellAddress;
  /** Proposed typed value. */
  readonly value: FormulaValue;
  /** Rule to evaluate. */
  readonly rule: ValidationRule;
  /** Optional caller cancellation. */
  readonly signal?: AbortSignal;
}

/** Stable validation result returned before any controller mutation. */
export type ValidationResult =
  | {
      /** Indicates the edit satisfies the rule. */
      readonly status: 'accepted';
      /** Empty diagnostics for an accepted edit. */
      readonly diagnostics: readonly [];
    }
  | {
      /** Indicates an invalid edit and its configured enforcement. */
      readonly status: 'rejected' | 'warning';
      /** Stable validation failure code. */
      readonly code: 'VALIDATION_REJECTED';
      /** Rule-specific diagnostics. */
      readonly diagnostics: readonly [
        {
          /** Stable validation failure code. */
          readonly code: 'VALIDATION_REJECTED';
          /** Rule that rejected or warned. */
          readonly ruleId: string;
        },
      ];
    }
  | {
      /** Indicates the validation source could not be evaluated safely. */
      readonly status: 'error';
      /** Stable source failure code. */
      readonly code:
        | 'VALIDATION_SOURCE_ERROR'
        | 'VALIDATION_SOURCE_TOO_LARGE'
        | 'VALIDATION_SOURCE_TIMEOUT'
        | 'VALIDATION_SOURCE_ABORTED';
      /** Rule-specific source diagnostics. */
      readonly diagnostics: readonly [
        {
          /** Stable source failure code. */
          readonly code: string;
          /** Rule whose source failed. */
          readonly ruleId: string;
        },
      ];
    };
