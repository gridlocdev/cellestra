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
  | ReferenceExpr
  | RangeExpr
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
  reference: (address: CellAddress): Expr => ({ kind: 'reference', address }),
  range: (range: CellRange): Expr => ({ kind: 'range', range }),
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
