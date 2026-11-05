import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import { Coercion } from '../evaluation/coercion.ts';
import { CriteriaMatcher } from './criteriaMatcher.ts';
import { defineFunction, type FormulaFunction, type FunctionEvaluator } from './formulaFunction.ts';
import { Arg, createMetadata, type FunctionCategory } from './functionMetadata.ts';
import { collectNumbers, isFailure, readInt, readNumber, roundHalfAway } from './helpers.ts';
import type { RangeValues } from './rangeValues.ts';

function aggregate(
  name: string,
  category: FunctionCategory,
  summary: string,
  example: string,
  reduce: (nums: number[]) => number,
): FormulaFunction {
  return defineFunction(
    createMetadata(name, category, summary, 'number', example, Arg.number('number', { repeatable: true })),
    (args, ev) => {
      const nums: number[] = [];
      const error = collectNumbers(ev, args, nums);
      return error ?? CellValue.number(reduce(nums));
    },
  );
}

function unary(name: string, summary: string, example: string, op: (n: number) => number): FormulaFunction {
  return defineFunction(
    createMetadata(name, 'math', summary, 'number', example, Arg.number('number')),
    (args, ev) => {
      const n = readNumber(ev, args[0]!);
      return isFailure(n) ? n.failed : CellValue.number(op(n));
    },
  );
}

function rounding(
  name: string,
  summary: string,
  example: string,
  round: (value: number, digits: number) => number,
): FormulaFunction {
  return defineFunction(
    createMetadata(
      name,
      'math',
      summary,
      'number',
      example,
      Arg.number('number'),
      Arg.number('digits', { optional: true }),
    ),
    (args, ev) => {
      const value = readNumber(ev, args[0]!);
      if (isFailure(value)) return value.failed;
      let digits = 0;
      if (args.length > 1) {
        const d = readInt(ev, args[1]!);
        if (isFailure(d)) return d.failed;
        digits = d;
      }
      return CellValue.number(round(value, digits));
    },
  );
}

function directional(value: number, digits: number, away: boolean): number {
  const factor = Math.pow(10, digits);
  const scaled = value * factor;
  const adjusted = away ? (scaled < 0 ? Math.floor(scaled) : Math.ceil(scaled)) : Math.trunc(scaled);
  return adjusted / factor;
}

function countValues(args: readonly Expr[], ev: FunctionEvaluator, onlyNumbers: boolean): number {
  let count = 0;
  for (const arg of args) {
    // COUNT counts numbers from references, but a value typed directly as an argument also counts
    // when it is a boolean or a numeric string — COUNT("x",1,TRUE)=2 (Excel/Calc).
    const fromReference = ev.tryEvaluateRange(arg) !== null;
    for (const value of ev.enumerateValues(arg)) {
      if (!onlyNumbers) {
        if (value.kind !== 'empty') count++;
      } else if (value.kind === 'number') count++;
      else if (!fromReference && value.kind !== 'empty' && typeof Coercion.toNumber(value) === 'number')
        count++;
    }
  }
  return count;
}

function sumOrCountIf(args: readonly Expr[], ev: FunctionEvaluator, sum: boolean): CellValue {
  const range = ev.tryEvaluateRange(args[0]!);
  if (!range) return CellValue.error(CellError.value('The first argument must be a range.'));
  const criteriaValue = ev.evaluate(args[1]!);
  if (criteriaValue.kind === 'error') return criteriaValue;
  const matcher = CriteriaMatcher.build(criteriaValue, ev.context.locale);

  let sumRange: RangeValues | null = null;
  if (sum && args.length > 2) {
    sumRange = ev.tryEvaluateRange(args[2]!);
    if (!sumRange) return CellValue.error(CellError.value('The sum range must be a range.'));
  }

  let total = 0;
  let count = 0;
  for (let r = 0; r < range.rowCount; r++) {
    for (let c = 0; c < range.columnCount; c++) {
      const cell = range.get(r, c);
      if (cell.kind === 'error') return cell;
      if (!matcher.matches(cell)) continue;
      count++;
      if (!sum) continue;
      // Excel/Calc anchor sum_range at its top-left and reshape it to the criteria range, so its own
      // size is ignored: read offset (r, c) from its top-left through the context.
      const target = sumRange
        ? ev.context.getCell(sumRange.range.sheet, sumRange.range.top + r, sumRange.range.left + c)
        : cell;
      if (target.kind === 'error') return target;
      if (target.kind === 'number') total += target.value;
    }
  }
  return CellValue.number(sum ? total : count);
}

/** The math & statistics functions. */
export function mathFunctions(): FormulaFunction[] {
  return [
    aggregate('SUM', 'math', 'Adds all the numbers in a range or list.', 'SUM(A1:A10)', (nums) =>
      nums.reduce((a, b) => a + b, 0),
    ),
    aggregate('PRODUCT', 'math', 'Multiplies all the numbers given.', 'PRODUCT(A1:A5)', (nums) =>
      nums.length === 0 ? 0 : nums.reduce((a, b) => a * b, 1),
    ),
    defineFunction(
      createMetadata(
        'AVERAGE',
        'statistical',
        'Returns the average (arithmetic mean) of the numbers.',
        'number',
        'AVERAGE(A1:A10)',
        Arg.number('number', { repeatable: true }),
      ),
      (args, ev) => {
        const nums: number[] = [];
        const error = collectNumbers(ev, args, nums);
        if (error) return error;
        return nums.length === 0
          ? CellValue.error(CellError.divideByZero('AVERAGE has no numeric values.'))
          : CellValue.number(nums.reduce((a, b) => a + b, 0) / nums.length);
      },
    ),
    aggregate('MIN', 'statistical', 'Returns the smallest number.', 'MIN(A1:A10)', (nums) =>
      nums.length === 0 ? 0 : nums.reduce((a, b) => (b < a ? b : a)),
    ),
    aggregate('MAX', 'statistical', 'Returns the largest number.', 'MAX(A1:A10)', (nums) =>
      nums.length === 0 ? 0 : nums.reduce((a, b) => (b > a ? b : a)),
    ),
    defineFunction(
      createMetadata(
        'COUNT',
        'statistical',
        'Counts how many cells contain numbers.',
        'number',
        'COUNT(A1:A10)',
        Arg.any('value', { repeatable: true }),
      ),
      (args, ev) => CellValue.number(countValues(args, ev, true)),
    ),
    defineFunction(
      createMetadata(
        'COUNTA',
        'statistical',
        'Counts how many cells are not empty.',
        'number',
        'COUNTA(A1:A10)',
        Arg.any('value', { repeatable: true }),
      ),
      (args, ev) => CellValue.number(countValues(args, ev, false)),
    ),
    unary('ABS', 'Returns the absolute value of a number.', 'ABS(-5)', Math.abs),
    unary('INT', 'Rounds a number down to the nearest integer.', 'INT(3.7)', Math.floor),
    defineFunction(
      createMetadata(
        'SQRT',
        'math',
        'Returns the positive square root of a number.',
        'number',
        'SQRT(16)',
        Arg.number('number'),
      ),
      (args, ev) => {
        const n = readNumber(ev, args[0]!);
        if (isFailure(n)) return n.failed;
        return n < 0
          ? CellValue.error(CellError.value('SQRT of a negative number.'))
          : CellValue.number(Math.sqrt(n));
      },
    ),
    defineFunction(
      createMetadata(
        'POWER',
        'math',
        'Raises a number to a power.',
        'number',
        'POWER(2, 10)',
        Arg.number('base'),
        Arg.number('exponent'),
      ),
      (args, ev) => {
        const a = readNumber(ev, args[0]!);
        if (isFailure(a)) return a.failed;
        const b = readNumber(ev, args[1]!);
        if (isFailure(b)) return b.failed;
        const result = Math.pow(a, b);
        return Number.isFinite(result)
          ? CellValue.number(result)
          : CellValue.error(CellError.value('POWER result is not a finite number.'));
      },
    ),
    defineFunction(
      createMetadata(
        'MOD',
        'math',
        'Returns the remainder after division (sign follows the divisor).',
        'number',
        'MOD(10, 3)',
        Arg.number('number'),
        Arg.number('divisor'),
      ),
      (args, ev) => {
        const n = readNumber(ev, args[0]!);
        if (isFailure(n)) return n.failed;
        const d = readNumber(ev, args[1]!);
        if (isFailure(d)) return d.failed;
        if (d === 0) return CellValue.error(CellError.divideByZero('MOD by zero.'));
        // Excel/Calc MOD: the result takes the sign of the divisor.
        return CellValue.number(n - d * Math.floor(n / d));
      },
    ),
    rounding(
      'ROUND',
      'Rounds a number to a given number of digits (half away from zero).',
      'ROUND(3.14159, 2)',
      roundHalfAway,
    ),
    rounding('ROUNDUP', 'Rounds a number away from zero.', 'ROUNDUP(3.14159, 2)', (v, d) =>
      directional(v, d, true),
    ),
    rounding('ROUNDDOWN', 'Rounds a number toward zero.', 'ROUNDDOWN(3.14159, 2)', (v, d) =>
      directional(v, d, false),
    ),
    defineFunction(
      createMetadata(
        'SUMIF',
        'math',
        'Adds the cells that meet a criteria.',
        'number',
        'SUMIF(A1:A10, ">5")',
        Arg.range('range'),
        Arg.any('criteria'),
        Arg.range('sum_range', { optional: true }),
      ),
      (args, ev) => sumOrCountIf(args, ev, true),
    ),
    defineFunction(
      createMetadata(
        'COUNTIF',
        'statistical',
        'Counts the cells that meet a criteria.',
        'number',
        'COUNTIF(A1:A10, ">5")',
        Arg.range('range'),
        Arg.any('criteria'),
      ),
      (args, ev) => sumOrCountIf(args, ev, false),
    ),
  ];
}
