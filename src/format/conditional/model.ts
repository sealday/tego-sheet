import type { DocumentCellAddress, DocumentCellRange } from '../../document';
import type { FormulaValue } from '../../formula';
import type { JsonValue } from '../../core/types/json';

/** Stable conditional-format identifier. */
export type ConditionalRuleId = string;

/** Conditions supported by the Phase 3A evaluator. */
export type ConditionalExpression =
  | {
      readonly type: 'cell-is';
      readonly operator:
        | 'equal'
        | 'notEqual'
        | 'greaterThan'
        | 'greaterThanOrEqual'
        | 'lessThan'
        | 'lessThanOrEqual';
      readonly value: number | string;
    }
  | { readonly type: 'blank' }
  | { readonly type: 'not-blank' }
  | { readonly type: 'text-contains'; readonly text: string }
  | { readonly type: 'formula'; readonly source: string };

/** A derived style patch; it is never written back to the base cell style. */
export type ConditionalStylePatch = Readonly<Record<string, JsonValue>>;

/** Presentation effect produced by a conditional rule. */
export type ConditionalEffect =
  | { readonly type: 'style'; readonly patch: ConditionalStylePatch }
  | {
      readonly type: 'color-scale';
      readonly minimumColor: string;
      readonly midpointColor?: string;
      readonly maximumColor: string;
    };

/** Ordered, range-scoped conditional-format rule. */
export interface ConditionalRule {
  readonly id: ConditionalRuleId;
  readonly priority: number;
  readonly stopIfTrue: boolean;
  readonly ranges: readonly DocumentCellRange[];
  readonly condition: ConditionalExpression;
  readonly effect: ConditionalEffect;
}

/** One cell evaluated against ordered conditional rules. */
export interface ConditionalEvaluationInput {
  readonly address: DocumentCellAddress;
  readonly value: FormulaValue;
  readonly text: string;
  readonly baseStyle: Readonly<Record<string, JsonValue>>;
  readonly rules: readonly ConditionalRule[];
  readonly lookup?: (address: DocumentCellAddress) => FormulaValue | undefined;
}

/** Deterministic conditional evaluation output. */
export interface ConditionalEvaluationResult {
  readonly matchedRuleIds: readonly ConditionalRuleId[];
  readonly stylePatch: ConditionalStylePatch;
  readonly diagnostics: readonly {
    readonly code: 'UNSUPPORTED_FORMAT_FEATURE';
    readonly ruleId: ConditionalRuleId;
  }[];
}
