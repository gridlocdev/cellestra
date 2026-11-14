import type { CellAddress } from '../../model/cellAddress.ts';
import type { CellRange } from '../../model/cellRange.ts';
import type { CellValue } from '../../values/cellValue.ts';

export type UnaryOperator = 'negate' | 'plus' | 'percent';

export type BinaryOperator =
  | 'add'
  | 'subtract'
  | 'multiply'
  | 'divide'
  | 'power'
  | 'concat'
  | 'equal'
  | 'notEqual'
  | 'less'
  | 'lessEqual'
  | 'greater'
  | 'greaterEqual'
  /** The reference union `~`: every area of both sides (Calc's reference list). */
  | 'union'
  /** The reference intersection `!`: the cells both sides share. */
  | 'intersect';

/** A literal number, string, or boolean. */
export interface LiteralExpr {
  readonly kind: 'literal';
  readonly value: CellValue;
}

/**
 * An inline array of constants, `{1,2;3,4}`: rows of equal length, each element a number, text,
 * boolean or error value.
 */
export interface ArrayExpr {
  readonly kind: 'array';
  readonly rows: readonly (readonly CellValue[])[];
}

/** A single cell reference such as `A1`, `$A$1`, or `Sheet2.A1`. */
export interface ReferenceExpr {
  readonly kind: 'reference';
  readonly address: CellAddress;
}

/** A rectangular range reference such as `A1:B3`. */
export interface RangeExpr {
  readonly kind: 'range';
  readonly range: CellRange;
}

/**
 * A range through several sheets (a 3D reference), `Sheet1.A1:Sheet3.B2`: the same cells on every
 * sheet from `range`'s sheet to `lastSheet` in tab order (either may come first). Both corners of
 * `range` carry the first sheet.
 */
export interface SheetSpanExpr {
  readonly kind: 'sheetSpan';
  readonly range: CellRange;
  readonly lastSheet: string;
}

/** A defined name / named range used in a formula (resolved at evaluation). */
export interface NameExpr {
  readonly kind: 'name';
  readonly name: string;
}

/**
 * A reference broken by an insert/delete (its target was deleted). Written as `#REF!` and
 * evaluates to a #REF! error — the formula structure is preserved so the user can fix it.
 */
export interface RefErrorExpr {
  readonly kind: 'refError';
}

export interface UnaryExpr {
  readonly kind: 'unary';
  readonly operator: UnaryOperator;
  readonly operand: Expr;
}

export interface BinaryExpr {
  readonly kind: 'binary';
  readonly operator: BinaryOperator;
  readonly left: Expr;
  readonly right: Expr;
}

/** A function call; `name` is upper-cased by the parser. */
export interface FunctionCallExpr {
  readonly kind: 'call';
  readonly name: string;
  readonly args: readonly Expr[];
}

/** A parsed formula expression node. */
export type Expr =
  | LiteralExpr
  | ArrayExpr
  | ReferenceExpr
  | RangeExpr
  | SheetSpanExpr
  | NameExpr
  | RefErrorExpr
  | UnaryExpr
  | BinaryExpr
  | FunctionCallExpr;

export type ExprKind = Expr['kind'];

const REF_ERROR: RefErrorExpr = Object.freeze({ kind: 'refError' });

/** Node constructors. */
export const Expr = {
  literal: (value: CellValue): Expr => ({ kind: 'literal', value }),
  array: (rows: readonly (readonly CellValue[])[]): Expr => ({ kind: 'array', rows }),
  reference: (address: CellAddress): Expr => ({ kind: 'reference', address }),
  range: (range: CellRange): Expr => ({ kind: 'range', range }),
  sheetSpan: (range: CellRange, lastSheet: string): Expr => ({ kind: 'sheetSpan', range, lastSheet }),
  name: (name: string): Expr => ({ kind: 'name', name }),
  refError: REF_ERROR as Expr,
  unary: (operator: UnaryOperator, operand: Expr): Expr => ({ kind: 'unary', operator, operand }),
  binary: (operator: BinaryOperator, left: Expr, right: Expr): Expr => ({
    kind: 'binary',
    operator,
    left,
    right,
  }),
  call: (name: string, args: readonly Expr[]): Expr => ({ kind: 'call', name, args }),
} as const;
