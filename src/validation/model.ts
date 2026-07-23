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

/** Typed Phase 3A validation rule. */
export type ValidationRule =
  | {
      /** Stable rule identifier. */
      readonly id: string;
      /** Validates numeric values. */
      readonly type: 'number';
      /** Numeric comparison definition. */
      readonly predicate: {
        /** Inclusive range comparison. */
        readonly operator: 'between';
        /** Inclusive minimum. */
        readonly minimum: number;
        /** Inclusive maximum. */
        readonly maximum: number;
      };
      /** Whether invalid edits are rejected or warned. */
      readonly behavior: 'reject' | 'warn';
      /** Whether a blank value bypasses validation. */
      readonly allowBlank: boolean;
    }
  | {
      /** Stable rule identifier. */
      readonly id: string;
      /** Validates membership in a list. */
      readonly type: 'list';
      /** List source definition. */
      readonly predicate: {
        /** Static or host-resolved allowed values. */
        readonly source: ValidationListSource;
      };
      /** Whether invalid edits are rejected or warned. */
      readonly behavior: 'reject' | 'warn';
      /** Whether a blank value bypasses validation. */
      readonly allowBlank: boolean;
    };

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
      readonly code: 'VALIDATION_SOURCE_ERROR' | 'VALIDATION_SOURCE_TOO_LARGE';
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
