import type { Expr } from '../ast/expr.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import {
  defineFunction,
  withTraits,
  type FormulaFunction,
  type FunctionEvaluator,
} from './formulaFunction.ts';
import { Arg, createMetadata } from './functionMetadata.ts';
import { collectNumbers, isFailure, readInt } from './helpers.ts';
import { KahanSum } from './kahanSum.ts';

/** A SUBTOTAL or AGGREGATE call anywhere in a formula: such cells are left out of a subtotal. */
const NESTED = /\b(SUBTOTAL|AGGREGATE)\s*\(/i;

const sum = (nums: readonly number[]): number => {
  const total = new KahanSum();
  for (const n of nums) total.add(n);
  return total.get();
};

/** The sum of squared differences from the mean. */
const squares = (nums: readonly number[]): number => {
  const mean = sum(nums) / nums.length;
  return sum(nums.map((n) => (n - mean) ** 2));
};

const divideByZero = () => CellValue.error(CellError.divideByZero('SUBTOTAL has too few numbers.'));

/** Function numbers 1–11, each from the numbers it reads (COUNTA reads its own count). */
const REDUCERS: Record<number, (nums: readonly number[], filled: number) => CellValue> = {
  1: (n) => (n.length === 0 ? divideByZero() : CellValue.number(sum(n) / n.length)),
  2: (n) => CellValue.number(n.length),
  3: (_, filled) => CellValue.number(filled),
  4: (n) => CellValue.number(n.length === 0 ? 0 : Math.max(...n)),
  5: (n) => CellValue.number(n.length === 0 ? 0 : Math.min(...n)),
  6: (n) => CellValue.number(n.length === 0 ? 0 : n.reduce((a, b) => a * b, 1)),
  7: (n) => (n.length < 2 ? divideByZero() : CellValue.number(Math.sqrt(squares(n) / (n.length - 1)))),
  8: (n) => (n.length < 1 ? divideByZero() : CellValue.number(Math.sqrt(squares(n) / n.length))),
  9: (n) => CellValue.number(sum(n)),
  10: (n) => (n.length < 2 ? divideByZero() : CellValue.number(squares(n) / (n.length - 1))),
  11: (n) => (n.length < 1 ? divideByZero() : CellValue.number(squares(n) / n.length)),
};

/**
 * SUBTOTAL(function; ref; …), as in Calc: rows a filter hid and cells that hold a subtotal of their
 * own are always left out; 101–111 also leave out rows hidden by hand.
 */
function subtotal(args: readonly Expr[], ev: FunctionEvaluator): CellValue {
  const which = readInt(ev, args[0]!);
  if (isFailure(which)) return which.failed;
  const skipHidden = which > 100;
  const reduce = REDUCERS[skipHidden ? which - 100 : which];
  if (!reduce) return CellValue.error(CellError.value(`SUBTOTAL has no function ${which}.`));

  const nums: number[] = [];
  let filled = 0;
  for (const arg of args.slice(1)) {
    const areas = ev.tryEvaluateAreas(arg);
    if (!areas) {
      const before = nums.length;
      const error = collectNumbers(ev, [arg], nums);
      if (error) return error;
      filled += nums.length - before;
      continue;
    }
    for (const area of areas) {
      const { top, left, sheet } = area.range;
      for (let r = 0; r < area.usedRowCount; r++) {
        if (!area.isArray) {
          const hidden = ev.context.rowHiddenBy?.(sheet, top + r) ?? null;
          if (hidden === 'filtered' || (hidden === 'hidden' && skipHidden)) continue;
        }
        for (let c = 0; c < area.usedColumnCount; c++) {
          if (!area.isArray && NESTED.test(ev.context.cellFormula?.(sheet, top + r, left + c) ?? ''))
            continue;
          const value = area.get(r, c);
          if (value.kind === 'error') return value;
          if (value.kind !== 'empty') filled++;
          if (value.kind === 'number') nums.push(value.value);
          else if (value.kind === 'boolean' && !area.isArray) nums.push(value.value ? 1 : 0);
        }
      }
    }
  }
  const result = reduce(nums, filled);
  return result.kind === 'number' && !Number.isFinite(result.value)
    ? CellValue.error(CellError.number('The result is not a finite number.'))
    : result;
}

/** SUBTOTAL. */
export function subtotalFunctions(): FormulaFunction[] {
  return [
    withTraits(
      defineFunction(
        createMetadata(
          'SUBTOTAL',
          'math',
          'Totals a range with one of 11 functions (9 is SUM), leaving out filtered rows and other subtotals; 101–111 also leave out hidden rows.',
          'number',
          'SUBTOTAL(9, A2:A20)',
          Arg.number('function'),
          Arg.any('reference', { repeatable: true }),
        ),
        subtotal,
      ),
      // Hiding or filtering rows changes the result without changing any cell it reads.
      { volatile: true },
    ),
  ];
}
