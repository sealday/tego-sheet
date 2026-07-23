import type { DocumentCellAddress, DocumentCellRange } from '../../document';
import type { FormulaValue } from '../../formula';
import type { JsonValue } from '../../core/types/json';

/** Stable conditional-format identifier. */
export type ConditionalRuleId = string;

/** Conditions supported by the Phase 3A evaluator. */
export type ConditionalExpression =
  | {
      /** Compares the cell scalar to a fixed value. */
      readonly type: 'cell-is';
      /** Comparison applied to the cell scalar. */
      readonly operator:
        | 'equal'
        | 'notEqual'
        | 'between'
        | 'notBetween'
        | 'greaterThan'
        | 'greaterThanOrEqual'
        | 'lessThan'
        | 'lessThanOrEqual';
      /** Fixed comparison value. */
      readonly value: number | string;
      /** Inclusive upper bound for between comparisons. */
      readonly value2?: number | string;
    }
  | {
      /** Matches blank values. */
      readonly type: 'blank';
    }
  | {
      /** Matches non-blank values. */
      readonly type: 'not-blank';
    }
  | {
      /** Matches formatted text containing a substring. */
      readonly type: 'text-contains';
      /** Required substring. */
      readonly text: string;
    }
  | {
      /** Evaluates a restricted formula condition. */
      readonly type: 'formula';
      /** Formula source beginning with an equals sign. */
      readonly source: string;
    };

/** A derived style patch; it is never written back to the base cell style. */
export type ConditionalStylePatch = Readonly<Record<string, JsonValue>>;

/** Presentation effect produced by a conditional rule. */
export type ConditionalEffect =
  | {
      /** Applies a derived style patch. */
      readonly type: 'style';
      /** Presentation properties to overlay. */
      readonly patch: ConditionalStylePatch;
    }
  | {
      /** Requests a color-scale presentation. */
      readonly type: 'color-scale';
      /** Low-end color. */
      readonly minimumColor: string;
      /** Optional middle color. */
      readonly midpointColor?: string;
      /** High-end color. */
      readonly maximumColor: string;
    };

/** Ordered, range-scoped conditional-format rule. */
export interface ConditionalRule {
  /** Stable rule identifier. */
  readonly id: ConditionalRuleId;
  /** Ascending evaluation order. */
  readonly priority: number;
  /** Whether a match prevents later rules from running. */
  readonly stopIfTrue: boolean;
  /** Target regions. */
  readonly ranges: readonly DocumentCellRange[];
  /** Match expression. */
  readonly condition: ConditionalExpression;
  /** Derived presentation effect. */
  readonly effect: ConditionalEffect;
}

/** One cell evaluated against ordered conditional rules. */
export interface ConditionalEvaluationInput {
  /** Cell being presented. */
  readonly address: DocumentCellAddress;
  /** Typed calculated value. */
  readonly value: FormulaValue;
  /** Already formatted cell text. */
  readonly text: string;
  /** Immutable base style. */
  readonly baseStyle: Readonly<object>;
  /** Candidate rules. */
  readonly rules: readonly ConditionalRule[];
  /** Optional side-effect-free referenced-cell lookup. */
  readonly lookup?: (address: DocumentCellAddress) => FormulaValue | undefined;
}

/** Deterministic conditional evaluation output. */
export interface ConditionalEvaluationResult {
  /** Matched rule identifiers in evaluation order. */
  readonly matchedRuleIds: readonly ConditionalRuleId[];
  /** Combined derived style overlay. */
  readonly stylePatch: ConditionalStylePatch;
  /** Recoverable unsupported-effect diagnostics. */
  readonly diagnostics: readonly {
    readonly code: 'UNSUPPORTED_FORMAT_FEATURE';
    readonly ruleId: ConditionalRuleId;
  }[];
}
