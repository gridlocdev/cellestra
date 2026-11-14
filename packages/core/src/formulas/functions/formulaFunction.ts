import type { Expr } from '../ast/expr.ts';
import type { EvaluationContext } from '../evaluation/evaluationContext.ts';
import type { CellRange } from '../../model/cellRange.ts';
import type { CellValue } from '../../values/cellValue.ts';
import type { FunctionMetadata } from './functionMetadata.ts';
import type { RangeValues } from './rangeValues.ts';

/**
 * The slice of the evaluator a function may use. Arguments arrive un-evaluated so functions can
 * short-circuit (IF), trap errors (IFERROR), or read whole ranges (SUM, VLOOKUP).
 */
export interface FunctionEvaluator {
  readonly context: EvaluationContext;
  /** Evaluate an argument to a single scalar value. */
  evaluate(expr: Expr): CellValue;
  /**
   * Flatten an argument into scalars: a range/named range yields each cell (row-major); any other
   * expression (including a single reference) yields its one evaluated value. Errors are yielded as-is.
   */
  enumerateValues(expr: Expr): Iterable<CellValue>;
  /** Resolve an argument to a 2D range view; null when it is not a reference/range/named range. */
  tryEvaluateRange(expr: Expr): RangeValues | null;
  /**
   * A range used where one value is needed: its one cell, or the cell in the formula's row (of a
   * one-column range) or column (of a one-row range); otherwise #VALUE!, as in Calc.
   */
  intersect(range: CellRange): CellValue;
}

/**
 * A spreadsheet function. Its metadata is the single source of truth: the evaluator checks arity
 * against it, autocomplete reads its hints, the Function Reference renders it.
 */
export interface FormulaFunction {
  readonly metadata: FunctionMetadata;
  /** Compute the result; return an error value rather than throwing for user-facing problems. */
  invoke(args: readonly Expr[], evaluator: FunctionEvaluator): CellValue;
  /**
   * For a function that returns a reference (INDEX): the range it refers to, or the error value it
   * gives instead. A range-taking argument such as SUM's reads the range through this.
   */
  readonly reference?: (args: readonly Expr[], evaluator: FunctionEvaluator) => CellRange | CellValue;
}

export type FunctionBody = (args: readonly Expr[], evaluator: FunctionEvaluator) => CellValue;

/** A function built from metadata plus a body. */
export function defineFunction(metadata: FunctionMetadata, body: FunctionBody): FormulaFunction {
  return Object.freeze({ metadata, invoke: body });
}

/** A function that returns a reference, plus the body that gives its value in a single-value place. */
export function defineReferenceFunction(
  metadata: FunctionMetadata,
  reference: NonNullable<FormulaFunction['reference']>,
  body: FunctionBody,
): FormulaFunction {
  return Object.freeze({ metadata, invoke: body, reference });
}
