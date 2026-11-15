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
   * Flatten an argument into scalars: a range/named range yields each cell (row-major), a union each
   * cell of each area in turn; any other
   * expression (including a single reference) yields its one evaluated value. Errors are yielded as-is.
   */
  enumerateValues(expr: Expr): Iterable<CellValue>;
  /** Resolve an argument to a 2D range view; null when it is not a reference/range/named range. */
  tryEvaluateRange(expr: Expr): RangeValues | null;
  /**
   * Resolve an argument to every area it names: one for a range, several for a union (`A1:A2~C1`).
   * Null when it is not a reference, or when it is a reference error (evaluate it for the error).
   */
  tryEvaluateAreas(expr: Expr): readonly RangeValues[] | null;
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
  /**
   * Recomputed every recalculation pass, whatever its arguments: TODAY and NOW, and SUBTOTAL, whose
   * result moves when rows are hidden or filtered without any cell it reads changing.
   */
  readonly volatile?: boolean;
  /**
   * Takes error arguments as values to inspect (IFERROR), so an argument on a deleted sheet is
   * passed in rather than turned into the call's #REF!.
   */
  readonly trapsErrors?: boolean;
  /**
   * Cells the function reads beyond the ones its arguments name, for the dependency graph: SUMIF's
   * sum range, read from its top-left at the size of the criteria range.
   */
  readonly extraReads?: (
    args: readonly Expr[],
    resolveName: (name: string) => CellRange | null,
  ) => readonly CellRange[];
}

/** What a function does beyond computing a value from its arguments. */
export type FunctionTraits = Pick<FormulaFunction, 'volatile' | 'trapsErrors' | 'extraReads'>;

export type FunctionBody = (args: readonly Expr[], evaluator: FunctionEvaluator) => CellValue;

/** A function built from metadata plus a body. */
export function defineFunction(metadata: FunctionMetadata, body: FunctionBody): FormulaFunction {
  return Object.freeze({ metadata, invoke: body });
}

/** The same function with `traits` added. */
export function withTraits(fn: FormulaFunction, traits: FunctionTraits): FormulaFunction {
  return Object.freeze({ ...fn, ...traits });
}

/** A function that returns a reference, plus the body that gives its value in a single-value place. */
export function defineReferenceFunction(
  metadata: FunctionMetadata,
  reference: NonNullable<FormulaFunction['reference']>,
  body: FunctionBody,
): FormulaFunction {
  return Object.freeze({ metadata, invoke: body, reference });
}
