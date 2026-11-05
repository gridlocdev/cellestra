import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import { Coercion } from '../evaluation/coercion.ts';
import type { FunctionEvaluator } from './formulaFunction.ts';

/**
 * Shared argument readers for the built-in functions. Each returns the coerced value, or a
 * `CellValue` error the function should return directly (check with `isErrorValue`).
 */
export type ArgResult<T> = T | ErrorResult;

/** An error to return from the function as-is. */
export interface ErrorResult {
  readonly failed: CellValue;
}

const fail = (value: CellValue): ErrorResult => ({ failed: value });

export const isFailure = (r: unknown): r is ErrorResult =>
  typeof r === 'object' && r !== null && 'failed' in r && !(r instanceof CellError);

export function readNumber(ev: FunctionEvaluator, arg: Expr): ArgResult<number> {
  const value = ev.evaluate(arg);
  if (value.kind === 'error') return fail(value);
  const n = Coercion.toNumber(value);
  return n instanceof CellError ? fail(CellValue.error(n)) : n;
}

/** A number truncated toward zero, clamped to the int32 range like a C# `(int)` cast. */
export function readInt(ev: FunctionEvaluator, arg: Expr): ArgResult<number> {
  const n = readNumber(ev, arg);
  if (isFailure(n)) return n;
  return toInt32(n);
}

export function toInt32(n: number): number {
  if (Number.isNaN(n)) return 0;
  const t = Math.trunc(n);
  return Math.max(-2147483648, Math.min(2147483647, t));
}

export function readText(ev: FunctionEvaluator, arg: Expr): ArgResult<string> {
  const value = ev.evaluate(arg);
  if (value.kind === 'error') return fail(value);
  const t = Coercion.toText(value);
  return t instanceof CellError ? fail(CellValue.error(t)) : t;
}

export function readBoolean(ev: FunctionEvaluator, arg: Expr): ArgResult<boolean> {
  const value = ev.evaluate(arg);
  if (value.kind === 'error') return fail(value);
  const b = Coercion.toBoolean(value);
  return b instanceof CellError ? fail(CellValue.error(b)) : b;
}

/**
 * Collect numeric values from arguments (flattening ranges). Values pulled from a reference/range
 * follow the LibreOffice rule — only numbers count — but a value typed directly as an argument is
 * coerced like arithmetic, so SUM(TRUE,FALSE,5)=6 and SUM(1,"x")=#VALUE!. Returns an error value
 * to propagate, else null.
 */
export function collectNumbers(
  ev: FunctionEvaluator,
  args: readonly Expr[],
  destination: number[],
): CellValue | null {
  for (const arg of args) {
    const fromReference = ev.tryEvaluateRange(arg) !== null;
    for (const value of ev.enumerateValues(arg)) {
      if (value.kind === 'error') return value;
      if (fromReference) {
        if (value.kind === 'number') destination.push(value.value);
        continue;
      }
      const n = Coercion.toNumber(value);
      if (n instanceof CellError) return CellValue.error(n);
      destination.push(n);
    }
  }
  return null;
}

/** .NET `Math.Round(x, MidpointRounding.AwayFromZero)`. */
export function roundAwayFromZero(x: number): number {
  const r = Math.round(Math.abs(x));
  // Math.round rounds .5 up; on the absolute value that is away from zero.
  return x < 0 ? -r : r;
}

/** Round half away from zero (Excel/Calc ROUND) to the given decimal places. */
export function roundHalfAway(value: number, digits: number): number {
  const factor = Math.pow(10, digits);
  return roundAwayFromZero(value * factor) / factor;
}
