import type { DocumentSheetId } from '../document';

/** Half-open character offsets into formula source. */
export interface SourceSpan {
  /** Inclusive starting offset. */
  readonly start: number;
  /** Exclusive ending offset. */
  readonly end: number;
}

/** Parsed A1 reference with optional resolved stable sheet identity. */
export interface FormulaReference {
  /** Display sheet name written in the source. */
  readonly sheetToken?: string;
  /** Stable sheet identity resolved from the document. */
  readonly sheetId?: DocumentSheetId;
  /** Zero-based row. */
  readonly row: number;
  /** Zero-based column. */
  readonly column: number;
  /** Whether copy translation preserves the row. */
  readonly rowAbsolute: boolean;
  /** Whether copy translation preserves the column. */
  readonly columnAbsolute: boolean;
}

/** Source location shared by every typed formula AST node. */
export interface FormulaNodeBase {
  /** Half-open source span occupied by the node. */
  readonly span: SourceSpan;
}

/** Typed, non-executable formula syntax tree. */
export type FormulaAst =
  | (FormulaNodeBase & {
      /** Numeric literal node. */
      readonly kind: 'number';
      /** Parsed finite number. */
      readonly value: number;
    })
  | (FormulaNodeBase & {
      /** String literal node. */
      readonly kind: 'string';
      /** Unescaped string content. */
      readonly value: string;
    })
  | (FormulaNodeBase & {
      /** Boolean literal node. */
      readonly kind: 'boolean';
      /** Boolean content. */
      readonly value: boolean;
    })
  | (FormulaNodeBase & {
      /** Standard spreadsheet error literal node. */
      readonly kind: 'error';
      /** Stable spreadsheet error. */
      readonly value: FormulaError;
    })
  | (FormulaNodeBase & {
      /** Single-cell reference node. */
      readonly kind: 'reference';
      /** Parsed reference. */
      readonly reference: FormulaReference;
    })
  | (FormulaNodeBase & {
      /** Inclusive rectangular range node. */
      readonly kind: 'range';
      /** First range endpoint. */
      readonly start: FormulaReference;
      /** Second range endpoint. */
      readonly end: FormulaReference;
    })
  | (FormulaNodeBase & {
      /** Unary operator node. */
      readonly kind: 'unary';
      /** Supported unary operator. */
      readonly operator: '-';
      /** Operand expression. */
      readonly operand: FormulaAst;
    })
  | (FormulaNodeBase & {
      /** Binary operator node. */
      readonly kind: 'binary';
      /** Supported arithmetic, concatenation, or comparison operator. */
      readonly operator:
        | '+'
        | '-'
        | '*'
        | '/'
        | '&'
        | '='
        | '=='
        | '<>'
        | '!='
        | '>'
        | '>='
        | '<'
        | '<=';
      /** Left operand. */
      readonly left: FormulaAst;
      /** Right operand. */
      readonly right: FormulaAst;
    })
  | (FormulaNodeBase & {
      /** Declared function call node. */
      readonly kind: 'call';
      /** Uppercase function name. */
      readonly name: string;
      /** Ordered argument expressions. */
      readonly arguments: readonly FormulaAst[];
    });

/** Stable spreadsheet errors supported by the formula core. */
export type FormulaError =
  | '#REF!'
  | '#VALUE!'
  | '#DIV/0!'
  | '#NAME?'
  | '#N/A'
  | '#NUM!'
  | '#SPILL!';

/** One non-array calculation value. */
export type ScalarFormulaValue =
  | {
      /** Blank value discriminator. */
      readonly type: 'blank';
    }
  | {
      /** Number value discriminator. */
      readonly type: 'number';
      /** Finite calculated number or Excel serial. */
      readonly value: number;
    }
  | {
      /** String value discriminator. */
      readonly type: 'string';
      /** Calculated string. */
      readonly value: string;
    }
  | {
      /** Boolean value discriminator. */
      readonly type: 'boolean';
      /** Calculated boolean. */
      readonly value: boolean;
    }
  | {
      /** Error value discriminator. */
      readonly type: 'error';
      /** Stable spreadsheet error. */
      readonly value: FormulaError;
    };

/** Typed calculation result, including reserved array results. */
export type FormulaValue =
  | ScalarFormulaValue
  | {
      /** Array value discriminator. */
      readonly type: 'array';
      /** Immutable rectangular scalar rows. */
      readonly rows: readonly (readonly ScalarFormulaValue[])[];
    };

/** Stable diagnostic emitted while compiling or calculating formulas. */
export interface FormulaDiagnostic {
  /** Machine-readable diagnostic code. */
  readonly code:
    | 'FORMULA_PARSE_ERROR'
    | 'FORMULA_UNKNOWN_FUNCTION'
    | 'FORMULA_CIRCULAR_REFERENCE'
    | 'FORMULA_REFERENCE_INVALID'
    | 'FORMULA_EVALUATION_LIMIT_EXCEEDED'
    | 'VOLATILE_FORMULA_NOT_RESOLVED'
    | 'ASYNC_FORMULA_NOT_RESOLVED';
  /** Human-readable diagnostic detail. */
  readonly message: string;
  /** Optional source location. */
  readonly span?: SourceSpan;
}

/** Syntax error raised for invalid restricted formula source. */
export class FormulaSyntaxError extends SyntaxError {
  /** Stable parse-error code. */
  readonly code = 'FORMULA_PARSE_ERROR';
  /** Source location that caused the error. */
  readonly span: SourceSpan;

  /** Creates a syntax error at a half-open source span. */
  constructor(message: string, span: SourceSpan) {
    super(message);
    this.name = 'FormulaSyntaxError';
    this.span = span;
  }
}

/** @internal */
export function freezeFormulaAst(ast: FormulaAst): FormulaAst {
  if (ast.kind === 'reference') Object.freeze(ast.reference);
  else if (ast.kind === 'range') {
    Object.freeze(ast.start);
    Object.freeze(ast.end);
  } else if (ast.kind === 'unary') {
    freezeFormulaAst(ast.operand);
  } else if (ast.kind === 'binary') {
    freezeFormulaAst(ast.left);
    freezeFormulaAst(ast.right);
  } else if (ast.kind === 'call') {
    for (const argument of ast.arguments) freezeFormulaAst(argument);
    Object.freeze(ast.arguments);
  }
  Object.freeze(ast.span);
  return Object.freeze(ast);
}
