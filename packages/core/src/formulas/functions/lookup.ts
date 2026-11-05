import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import { ValueComparer } from '../evaluation/valueComparer.ts';
import { CriteriaMatcher } from './criteriaMatcher.ts';
import { defineFunction, type FormulaFunction, type FunctionEvaluator } from './formulaFunction.ts';
import { Arg, createMetadata } from './functionMetadata.ts';
import { isFailure, readBoolean, readInt } from './helpers.ts';
import type { RangeValues } from './rangeValues.ts';

/** Position of `lookup` in `values` (−1 when absent), or the error value to return. */
function search(
  values: readonly CellValue[],
  lookup: CellValue,
  matchType: number,
  locale: string,
): number | CellValue {
  if (matchType === 0) {
    // Exact match, honouring text wildcards.
    const matcher = lookup.kind === 'text' ? CriteriaMatcher.build(lookup, locale) : null;
    for (let i = 0; i < values.length; i++) {
      const v = values[i]!;
      if (v.kind === 'error') return v;
      const isMatch = matcher ? matcher.matches(v) : ValueComparer.compare(v, lookup) === 0;
      if (isMatch) return i;
    }
    return -1;
  }
  // Approximate: assume sorted (ascending for 1, descending for -1); take the last value on the
  // correct side of the lookup.
  let best = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i]!;
    if (v.kind === 'error') return v;
    if (v.kind === 'empty') continue; // Calc skips empty cells in a sorted lookup
    const cmp = ValueComparer.compare(v, lookup);
    if (cmp instanceof CellError) return CellValue.error(cmp);
    const onSide = matchType > 0 ? cmp <= 0 : cmp >= 0;
    if (!onSide) break; // sorted assumption: once past the boundary, stop
    best = i;
  }
  return best;
}

/**
 * `search` over the first `used` cells of a vector of `length`, where every cell past `used` is
 * empty (beyond the sheet's data). An exact match for an empty cell lands on the first of those.
 */
function searchUsed(
  read: (i: number) => CellValue,
  used: number,
  length: number,
  lookup: CellValue,
  matchType: number,
  locale: string,
): number | CellValue {
  const position = search(
    Array.from({ length: used }, (_, i) => read(i)),
    lookup,
    matchType,
    locale,
  );
  if (position !== -1 || used >= length || matchType !== 0) return position;
  return search([CellValue.empty], lookup, 0, locale) === 0 ? used : -1;
}

function lookupTable(args: readonly Expr[], ev: FunctionEvaluator, vertical: boolean): CellValue {
  const lookup = ev.evaluate(args[0]!);
  if (lookup.kind === 'error') return lookup;
  const table = ev.tryEvaluateRange(args[1]!);
  if (!table) return CellValue.error(CellError.value('The lookup table must be a range.'));
  const index = readInt(ev, args[2]!);
  if (isFailure(index)) return index.failed;
  let approximate = true;
  if (args.length > 3) {
    const a = readBoolean(ev, args[3]!);
    if (isFailure(a)) return a.failed;
    approximate = a;
  }
  // The first row/column is the search vector.
  const position = searchUsed(
    (i) => (vertical ? table.get(i, 0) : table.get(0, i)),
    vertical ? table.usedRowCount : table.usedColumnCount,
    vertical ? table.rowCount : table.columnCount,
    lookup,
    approximate ? 1 : 0,
    ev.context.locale,
  );
  if (typeof position !== 'number') return position;
  if (position < 0) return CellValue.error(CellError.notAvailable('No match found.'));
  const otherCount = vertical ? table.columnCount : table.rowCount;
  if (index < 1 || index > otherCount) {
    return CellValue.error(CellError.reference(`Index ${index} is outside the table.`));
  }
  return vertical ? table.get(position, index - 1) : table.get(index - 1, position);
}

function index(range: RangeValues, rowArg: number, colArg: number | null): CellValue {
  let rowWithin: number;
  let colWithin: number;
  if (colArg !== null) {
    rowWithin = rowArg - 1;
    colWithin = colArg - 1;
  } else if (range.rowCount === 1) {
    rowWithin = 0;
    colWithin = rowArg - 1;
  } else if (range.columnCount === 1) {
    rowWithin = rowArg - 1;
    colWithin = 0;
  } else {
    return CellValue.error(CellError.reference('INDEX needs a column number for a 2D range.'));
  }
  if (rowWithin < 0 || rowWithin >= range.rowCount || colWithin < 0 || colWithin >= range.columnCount) {
    return CellValue.error(CellError.reference('INDEX position is outside the range.'));
  }
  return range.get(rowWithin, colWithin);
}

/** The lookup & reference functions. */
export function lookupFunctions(): FormulaFunction[] {
  return [
    defineFunction(
      createMetadata(
        'VLOOKUP',
        'lookup',
        'Looks up a value in the first column of a table and returns a value in the same row.',
        'any',
        'VLOOKUP(A1, D1:F20, 3, FALSE)',
        Arg.any('lookup'),
        Arg.range('table'),
        Arg.number('column'),
        Arg.boolean('approximate', { optional: true }),
      ),
      (args, ev) => lookupTable(args, ev, true),
    ),
    defineFunction(
      createMetadata(
        'HLOOKUP',
        'lookup',
        'Looks up a value in the first row of a table and returns a value in the same column.',
        'any',
        'HLOOKUP(A1, D1:Z2, 2, FALSE)',
        Arg.any('lookup'),
        Arg.range('table'),
        Arg.number('row'),
        Arg.boolean('approximate', { optional: true }),
      ),
      (args, ev) => lookupTable(args, ev, false),
    ),
    defineFunction(
      createMetadata(
        'MATCH',
        'lookup',
        'Returns the position of a value within a row or column.',
        'number',
        'MATCH(A1, D1:D20, 0)',
        Arg.any('lookup'),
        Arg.range('range'),
        Arg.number('match_type', { optional: true }),
      ),
      (args, ev) => {
        const lookup = ev.evaluate(args[0]!);
        if (lookup.kind === 'error') return lookup;
        const range = ev.tryEvaluateRange(args[1]!);
        if (!range) return CellValue.error(CellError.value('MATCH needs a range to search.'));
        let matchType = 1;
        if (args.length > 2) {
          const m = readInt(ev, args[2]!);
          if (isFailure(m)) return m.failed;
          matchType = m;
        }
        if (range.rowCount !== 1 && range.columnCount !== 1) {
          return CellValue.error(CellError.notAvailable('MATCH needs a single row or column.'));
        }
        const vertical = range.columnCount === 1;
        const position = searchUsed(
          (i) => (vertical ? range.get(i, 0) : range.get(0, i)),
          vertical ? range.usedRowCount : range.usedColumnCount,
          vertical ? range.rowCount : range.columnCount,
          lookup,
          matchType,
          ev.context.locale,
        );
        if (typeof position !== 'number') return position;
        return position < 0
          ? CellValue.error(CellError.notAvailable('No match found.'))
          : CellValue.number(position + 1);
      },
    ),
    defineFunction(
      createMetadata(
        'INDEX',
        'lookup',
        'Returns the value at a position within a range.',
        'any',
        'INDEX(A1:C10, 2, 3)',
        Arg.range('range'),
        Arg.number('row'),
        Arg.number('column', { optional: true }),
      ),
      (args, ev) => {
        const range = ev.tryEvaluateRange(args[0]!);
        if (!range) return CellValue.error(CellError.value('INDEX needs a range.'));
        const rowArg = readInt(ev, args[1]!);
        if (isFailure(rowArg)) return rowArg.failed;
        let colArg: number | null = null;
        if (args.length > 2) {
          const c = readInt(ev, args[2]!);
          if (isFailure(c)) return c.failed;
          colArg = c;
        }
        return index(range, rowArg, colArg);
      },
    ),
  ];
}
