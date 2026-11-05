import { Evaluator } from '../../src/formulas/evaluation/evaluator.ts';
import type { CivilDateTime } from '../../src/formulas/evaluation/dateConversion.ts';
import { WorkbookEvaluationContext } from '../../src/formulas/evaluation/workbookEvaluationContext.ts';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import type { CellValue } from '../../src/values/cellValue.ts';

const registry = FunctionRegistry.createDefault();

/** The test stand-in for .NET's InvariantCulture. */
export const INVARIANT = 'en-US';

/** A civil date-time for injected clocks. */
export function civil(
  year: number,
  month: number,
  day: number,
  hour = 0,
  minute = 0,
  second = 0,
): CivilDateTime {
  return { year, month, day, hour, minute, second, millisecond: 0 };
}

/** Evaluate a formula string against a fresh in-memory workbook. */
export function evaluate(
  formula: string,
  setup?: (sheet: Worksheet) => void,
  now?: CivilDateTime,
  locale?: string,
): CellValue {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  setup?.(sheet);
  const context = new WorkbookEvaluationContext(
    workbook,
    sheet.name,
    locale ?? INVARIANT,
    now === undefined ? undefined : () => now,
  );
  return new Evaluator(context, registry).evaluate(Parser.parse(formula));
}

export function num(formula: string, setup?: (sheet: Worksheet) => void, now?: CivilDateTime): number {
  const value = evaluate(formula, setup, now);
  if (value.kind === 'error') {
    throw new Error(
      `Expected a number from '${formula}' but got ${value.error.token}: ${value.error.reason}`,
    );
  }
  if (value.kind !== 'number') throw new Error(`Expected a number from '${formula}' but got ${value.kind}.`);
  return value.value;
}

export function text(formula: string, setup?: (sheet: Worksheet) => void): string {
  const value = evaluate(formula, setup);
  if (value.kind !== 'text') throw new Error(`Expected text from '${formula}' but got ${value.kind}.`);
  return value.value;
}

export function bool(formula: string, setup?: (sheet: Worksheet) => void): boolean {
  const value = evaluate(formula, setup);
  if (value.kind !== 'boolean')
    throw new Error(`Expected a boolean from '${formula}' but got ${value.kind}.`);
  return value.value;
}

export function errorToken(formula: string, setup?: (sheet: Worksheet) => void): string {
  const value = evaluate(formula, setup);
  if (value.kind !== 'error') throw new Error(`Expected an error from '${formula}' but got ${value.kind}.`);
  return value.error.token;
}
