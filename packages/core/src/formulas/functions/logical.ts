import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import { Coercion } from '../evaluation/coercion.ts';
import { defineFunction, type FormulaFunction, type FunctionEvaluator } from './formulaFunction.ts';
import { Arg, createMetadata } from './functionMetadata.ts';
import { isFailure, readBoolean } from './helpers.ts';

function reduce(args: readonly Expr[], ev: FunctionEvaluator, requireAll: boolean): CellValue {
  let any = false;
  let accumulator = requireAll;
  for (const arg of args) {
    const fromReference = ev.tryEvaluateAreas(arg) !== null;
    for (const value of ev.enumerateValues(arg)) {
      if (value.kind === 'error') return value;
      if (value.kind === 'empty') continue;
      const b = Coercion.toBoolean(value);
      if (b instanceof CellError) {
        if (fromReference) continue; // non-logical text in a range is ignored, per Calc
        return CellValue.error(b); // a direct non-logical argument errors
      }
      any = true;
      accumulator = requireAll ? accumulator && b : accumulator || b;
    }
  }
  return any
    ? CellValue.boolean(accumulator)
    : CellValue.error(CellError.value('No logical values were supplied.'));
}

/** The logical functions. */
export function logicalFunctions(): FormulaFunction[] {
  return [
    defineFunction(
      createMetadata(
        'IF',
        'logical',
        'Returns one value if a condition is true, another if false.',
        'any',
        'IF(A1>0, "yes", "no")',
        Arg.boolean('condition'),
        Arg.any('value_if_true'),
        Arg.any('value_if_false', { optional: true }),
      ),
      (args, ev) => {
        const condition = readBoolean(ev, args[0]!);
        if (isFailure(condition)) return condition.failed;
        if (condition) return ev.evaluate(args[1]!);
        return args.length > 2 ? ev.evaluate(args[2]!) : CellValue.boolean(false);
      },
    ),
    defineFunction(
      createMetadata(
        'IFERROR',
        'logical',
        'Returns a fallback value when an expression is an error.',
        'any',
        'IFERROR(A1/B1, 0)',
        Arg.any('value'),
        Arg.any('value_if_error'),
      ),
      (args, ev) => {
        const value = ev.evaluate(args[0]!);
        return value.kind === 'error' ? ev.evaluate(args[1]!) : value;
      },
    ),
    defineFunction(
      createMetadata(
        'AND',
        'logical',
        'TRUE if every argument is true.',
        'boolean',
        'AND(A1>0, B1>0)',
        Arg.boolean('logical', { repeatable: true }),
      ),
      (args, ev) => reduce(args, ev, true),
    ),
    defineFunction(
      createMetadata(
        'OR',
        'logical',
        'TRUE if any argument is true.',
        'boolean',
        'OR(A1>0, B1>0)',
        Arg.boolean('logical', { repeatable: true }),
      ),
      (args, ev) => reduce(args, ev, false),
    ),
    defineFunction(
      createMetadata(
        'NOT',
        'logical',
        'Reverses a logical value.',
        'boolean',
        'NOT(A1)',
        Arg.boolean('logical'),
      ),
      (args, ev) => {
        const b = readBoolean(ev, args[0]!);
        return isFailure(b) ? b.failed : CellValue.boolean(!b);
      },
    ),
    defineFunction(
      createMetadata('TRUE', 'logical', 'Returns the logical value TRUE.', 'boolean', 'TRUE()'),
      () => CellValue.boolean(true),
    ),
    defineFunction(
      createMetadata('FALSE', 'logical', 'Returns the logical value FALSE.', 'boolean', 'FALSE()'),
      () => CellValue.boolean(false),
    ),
  ];
}
