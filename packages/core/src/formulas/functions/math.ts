import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import { Coercion } from '../evaluation/coercion.ts';
import { approxFloor, approxSub } from '../evaluation/valueComparer.ts';
import { CriteriaMatcher } from './criteriaMatcher.ts';
import { defineFunction, type FormulaFunction, type FunctionEvaluator } from './formulaFunction.ts';
import { Arg, createMetadata, type FunctionCategory } from './functionMetadata.ts';
import { KahanSum } from './kahanSum.ts';
import { collectNumbers, isFailure, readInt, readNumber, roundHalfAway, roundScaled } from './helpers.ts';
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
      if (error) return error;
      const total = reduce(nums);
      return Number.isFinite(total)
        ? CellValue.number(total)
        : CellValue.error(CellError.number('The result is not a finite number.'));
    },
  );
}

/**
 * The numbers of SUM/AVERAGE-style arguments, compensated-summed in Calc's order: the arguments
 * from last to first (Calc pops them off its stack) and each range column by column, top to bottom.
 * Order matters: SUM(0.2,-1E16,1E16) is 0.2 in Calc, SUM(1E16,-1E16,0.2) is 0.
 */
function sumArguments(
  ev: FunctionEvaluator,
  args: readonly Expr[],
): { sum: number; count: number } | CellValue {
  const groups: number[][] = [];
  for (const arg of args) {
    const nums: number[] = [];
    const areas = arg.kind === 'reference' ? null : ev.tryEvaluateAreas(arg);
    if (areas) {
      for (const range of areas) {
        for (let c = 0; c < range.usedColumnCount; c++) {
          for (let r = 0; r < range.usedRowCount; r++) {
            const value = range.get(r, c);
            if (value.kind === 'error') return value;
            if (value.kind === 'number') nums.push(value.value);
            else if (value.kind === 'boolean' && !range.isArray) nums.push(value.value ? 1 : 0);
          }
        }
      }
    } else {
      const error = collectNumbers(ev, [arg], nums);
      if (error) return error;
    }
    groups.push(nums);
  }
  const total = new KahanSum();
  let count = 0;
  for (let g = groups.length - 1; g >= 0; g--) {
    for (const n of groups[g]!) total.add(n);
    count += groups[g]!.length;
  }
  return { sum: total.get(), count };
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
      const rounded = round(value, digits);
      // A non-finite result surfaces as #NUM!, like the evaluator's.
      return Number.isFinite(rounded)
        ? CellValue.number(rounded)
        : CellValue.error(CellError.number('The result is not a finite number.'));
    },
  );
}

function directional(value: number, digits: number, away: boolean): number {
  return roundScaled(value, digits, (s) => (away ? (s < 0 ? Math.floor(s) : Math.ceil(s)) : Math.trunc(s)));
}

function countValues(args: readonly Expr[], ev: FunctionEvaluator, onlyNumbers: boolean): number {
  let count = 0;
  for (const arg of args) {
    // COUNT counts numbers from references, but a value typed directly as an argument also counts
    // when it is a boolean or a numeric string — COUNT("x",1,TRUE)=2 (Excel/Calc).
    const areas = ev.tryEvaluateAreas(arg);
    const fromReference = areas !== null;
    // A boolean cell is a number to Calc; one in an inline array is not.
    const fromCells = fromReference && !areas.some((a) => a.isArray);
    for (const value of ev.enumerateValues(arg)) {
      if (!onlyNumbers) {
        if (value.kind !== 'empty') count++;
      } else if (value.kind === 'number' || (fromCells && value.kind === 'boolean')) count++;
      else if (!fromReference && value.kind !== 'empty' && typeof Coercion.toNumber(value) === 'number')
        count++;
    }
  }
  return count;
}

/** The cell at (r, c) from `sumRange`'s top-left, read past its own size; an array has nothing past it. */
function reshapedCell(sumRange: RangeValues, r: number, c: number, ev: FunctionEvaluator): CellValue {
  if (!sumRange.isArray)
    return ev.context.getCell(sumRange.range.sheet, sumRange.range.top + r, sumRange.range.left + c);
  return r < sumRange.rowCount && c < sumRange.columnCount ? sumRange.get(r, c) : CellValue.empty;
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

  // Past the data of both ranges every criteria cell and every sum cell is empty: read only up to
  // it, then count the empty remainder in one step if the criteria match an empty cell.
  const rows = Math.min(range.rowCount, Math.max(range.reach.rows, sumRange?.reach.rows ?? 0));
  const columns = Math.min(range.columnCount, Math.max(range.reach.columns, sumRange?.reach.columns ?? 0));
  const total = new KahanSum();
  let count = 0;
  // Column by column, as Calc walks a range: the order decides what compensated summation cancels.
  for (let c = 0; c < columns; c++) {
    for (let r = 0; r < rows; r++) {
      const cell = range.get(r, c);
      // An error in the criteria range never matches (Calc), so it does not poison the result.
      if (cell.kind === 'error' || !matcher.matches(cell)) continue;
      count++;
      if (!sum) continue;
      // Excel/Calc anchor sum_range at its top-left and reshape it to the criteria range, so its own
      // size is ignored: read offset (r, c) from its top-left through the context.
      const target = sumRange ? reshapedCell(sumRange, r, c, ev) : cell;
      if (target.kind === 'error') return target;
      // A TRUE in the sum range adds 1, as in Calc, where a boolean cell is a number.
      if (target.kind === 'number') total.add(target.value);
      else if (target.kind === 'boolean') total.add(target.value ? 1 : 0);
    }
  }
  if (matcher.matches(CellValue.empty)) count += range.rowCount * range.columnCount - rows * columns;
  return CellValue.number(sum ? total.get() : count);
}

/**
 * SUMPRODUCT over ranges, inline arrays or single values, all the same size (else #VALUE!). As in
 * Calc, text and empty cells are 0, a boolean cell is 1 or 0, and an error anywhere is the result.
 */
function sumProduct(args: readonly Expr[], ev: FunctionEvaluator): CellValue {
  type Grid = {
    rows: number;
    columns: number;
    /** How many rows from the top can hold a value. */
    reach: number;
    get: (r: number, c: number) => CellValue;
    /** Cells on a sheet, where a boolean is a number (not an inline array or a typed value). */
    cells: boolean;
  };
  const grids: Grid[] = [];
  for (const arg of args) {
    const range = ev.tryEvaluateRange(arg);
    if (range) {
      grids.push({
        rows: range.rowCount,
        columns: range.columnCount,
        reach: range.reach.rows,
        get: (r, c) => range.get(r, c),
        cells: !range.isArray,
      });
    } else {
      const value = ev.evaluate(arg);
      grids.push({ rows: 1, columns: 1, reach: 1, get: () => value, cells: true });
    }
  }
  const { rows, columns } = grids[0]!;
  if (grids.some((g) => g.rows !== rows || g.columns !== columns))
    return CellValue.error(CellError.value('SUMPRODUCT needs ranges of the same size.'));
  // Rows past every range's data are all empty, so their products are 0.
  const reach = Math.min(rows, ...grids.map((g) => g.reach));
  const total = new KahanSum();
  for (let r = 0; r < reach; r++) {
    for (let c = 0; c < columns; c++) {
      let product = 1;
      for (const grid of grids) {
        const value = grid.get(r, c);
        if (value.kind === 'error') return value;
        product *=
          value.kind === 'number'
            ? value.value
            : value.kind === 'boolean' && grid.cells && value.value
              ? 1
              : 0;
      }
      total.add(product);
    }
  }
  const sum = total.get();
  return Number.isFinite(sum)
    ? CellValue.number(sum)
    : CellValue.error(CellError.number('The result is not a finite number.'));
}

/** The math & statistics functions. */
export function mathFunctions(): FormulaFunction[] {
  return [
    defineFunction(
      createMetadata(
        'SUM',
        'math',
        'Adds all the numbers in a range or list.',
        'number',
        'SUM(A1:A10)',
        Arg.number('number', { repeatable: true }),
      ),
      (args, ev) => {
        const result = sumArguments(ev, args);
        if (!('sum' in result)) return result;
        return Number.isFinite(result.sum)
          ? CellValue.number(result.sum)
          : CellValue.error(CellError.number('The result is not a finite number.'));
      },
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
        const result = sumArguments(ev, args);
        if (!('sum' in result)) return result;
        return result.count === 0
          ? CellValue.error(CellError.divideByZero('AVERAGE has no numeric values.'))
          : CellValue.number(result.sum / result.count);
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
          ? CellValue.error(CellError.number('SQRT of a negative number.'))
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
          : CellValue.error(CellError.number('POWER result is not a finite number.'));
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
        // Calc's MOD (ScInterpreter::ScMod): the result takes the sign of the divisor, and the
        // quotient's floor and the subtraction both strip binary noise, so MOD(0.3, 0.1) is 0. When
        // precision runs out the remainder can land on the divisor itself (MOD(-1E-20, 1) = 1);
        // Calc reports #VALUE! there.
        const result = approxSub(n, approxFloor(n / d) * d);
        if (d > 0 ? result >= 0 && result < d : result <= 0 && result > d) return CellValue.number(result);
        return CellValue.error(CellError.value('MOD has no exact remainder at this precision.'));
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
        'SUMPRODUCT',
        'math',
        'Multiplies matching cells of same-sized ranges and adds up the products.',
        'number',
        'SUMPRODUCT(A1:A5, B1:B5)',
        Arg.any('array', { repeatable: true }),
      ),
      sumProduct,
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
